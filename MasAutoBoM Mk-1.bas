Attribute VB_Name = "Module1"
Option Explicit

'=============================
' Global / Parent configuration
'=============================

' Source & Target sheets
Public Const SRC_SHEET As String = "PLA Pipe & Ftgs"
Public Const TGT_SHEET As String = "OrderSheet"

' Target content rows
Public Const TGT_STARTROW As Long = 12
Public Const TGT_ENDROW As Long = 78   ' inclusive

' Placeholder for missing source values
Public Const NULL_PLACEHOLDER As String = "N/A"

' Target columns (numeric indexes)
Public Const TGT_QTY_COL As Long = 6          ' F
Public Const TGT_LEAD_COL As Long = 2         ' B (description column for duplicate detection)

' Columns to clear when archiving/clearing (content band)
Public Const TGT_FIRST_CONTENT_COL As Long = 2 ' B
Public Const TGT_LAST_CONTENT_COL As Long = 10 ' J

' Capacity warnings / sheet duplication
Public Const WARN_THRESHOLD As Long = 5
Public Const ENABLE_CAPACITY_WARNINGS As Boolean = True
Public Const ENABLE_AUTO_DUPLICATE_AND_CLEAR As Boolean = True

' Optional features toggles
Public Const ENABLE_ASK_QTY_FROM_FLAG As Boolean = True   ' reads a named range AskQtyFlag if present
Public Const ENABLE_MERGE_DUPLICATES As Boolean = True

' Named range (Boolean TRUE/FALSE) to control quantity prompt; if missing, defaults to False
Public Const NAME_ASK_QTY_FLAG As String = "AskQtyFlag"

Public ASK As Boolean

' Scan safety limit for row searching
Public Const MAX_SCAN As Long = 5000

'-----------------------------
' Utility - convert col letter to number (optional helper)
'-----------------------------
Public Function ColNumber(ByVal colLetter As String) As Long
    colLetter = UCase$(Trim$(colLetter))
    Dim i As Long, n As Long
    For i = 1 To Len(colLetter)
        n = n * 26 + (Asc(Mid$(colLetter, i, 1)) - 64)
    Next i
    ColNumber = n
End Function



'==========================================================
' Resolve caller row, next empty row, duplicate detection,
' capacity checks, and the single-point data copy kernel.
'==========================================================

' Safely get the row of the clicked button.
' Works for Forms/Shapes. For ActiveX, pass CallerRow explicitly to the top-level sequence.
Public Function ResolveCallerRow() As Long
    On Error GoTo Fail
    Dim nm As String
    nm = Application.Caller
    If LenB(nm) > 0 Then
        ResolveCallerRow = ActiveSheet.Shapes(nm).TopLeftCell.Row
        Exit Function
    End If
Fail:
    ResolveCallerRow = 0
End Function

' Return remaining slots in the target grid
Public Function RemainingSlots(wsTgt As Worksheet, ByVal leadCol As Long, _
                               ByVal startRow As Long, ByVal endRow As Long) As Long
    Dim r As Long, used As Long
    For r = startRow To endRow
        If LenB(wsTgt.Cells(r, leadCol).Value) > 0 Then used = used + 1
    Next r
    RemainingSlots = (endRow - startRow + 1) - used
End Function







Public Sub CheckForSpacesRemaining(ByVal wsTgt As Worksheet)

    ' If both features are off, no work to do.
    If (Not ENABLE_CAPACITY_WARNINGS) And (Not ENABLE_AUTO_DUPLICATE_AND_CLEAR) Then Exit Sub

    Dim freeRows As Long
    freeRows = RemainingSlots(wsTgt, TGT_LEAD_COL, TGT_STARTROW, TGT_ENDROW)

    If ENABLE_CAPACITY_WARNINGS Then
        If freeRows <= WARN_THRESHOLD And freeRows > 0 Then
            MsgBox "Warning: only " & freeRows & " rows remain on " & wsTgt.Name & ".", vbInformation
        End If
    End If

    If ENABLE_AUTO_DUPLICATE_AND_CLEAR And freeRows <= 0 Then
        DuplicateAndArchive wsTgt
        ClearContentBand wsTgt, TGT_STARTROW, TGT_ENDROW, TGT_FIRST_CONTENT_COL, TGT_LAST_CONTENT_COL
        MsgBox "The sheet has been flipped and Please continue.", vbInformation
    End If

End Sub





' Find next empty row in the lead column within [startRow, endRow], or 0 if none.
Public Function NextEmptyRow(wsTgt As Worksheet, ByVal leadCol As Long, _
                             ByVal startRow As Long, ByVal endRow As Long) As Long
    Dim r As Long
    For r = startRow To endRow
        If LenB(wsTgt.Cells(r, leadCol).Value) = 0 Then
            NextEmptyRow = r
            Exit Function
        End If
    Next r
    NextEmptyRow = 0
End Function

' Find duplicate row by exact match on lead datapoint; 0 if not found.
Public Function FindDuplicateRow(wsTgt As Worksheet, ByVal leadValue As String, _
                                 ByVal leadCol As Long, ByVal startRow As Long, _
                                 ByVal endRow As Long) As Long
    Dim r As Long
    If LenB(leadValue) = 0 Then Exit Function
    For r = startRow To endRow
        If CStr(wsTgt.Cells(r, leadCol).Value) = leadValue Then
            FindDuplicateRow = r
            Exit Function
        End If
    Next r
    FindDuplicateRow = 0
End Function

Public Sub CheckBox_Ask_Click()


    Dim cb As CheckBox
    Set cb = ActiveSheet.CheckBoxes(Application.Caller)

    ASK = (cb.Value = 1) ' 1 = checked; -4146 = unchecked

    If ASK Then
        MsgBox "Multi Select On", vbInformation, "Ask State"
    Else
        MsgBox "Multi Select Off", vbCritical, "Ask State"
    End If
End Sub

' Ask quantity if enabled by named flag; default = 1 if not enabled or not numeric.
Public Function ResolveQuantity(Optional ByVal defaultQty As Double = 1#) As Double
    On Error Resume Next
   Dim cb As CheckBox
    Set cb = ActiveSheet.CheckBoxes(Application.Caller)
    ASK = (cb.Value = 1) ' 1 = checked; -4146 = unchecked
    If ENABLE_ASK_QTY_FROM_FLAG Then
        ' If named range exists and is TRUE, ask the user
        Dim rng As Range
        Set rng = Nothing
        On Error Resume Next
        Set rng = Range(NAME_ASK_QTY_FLAG)
        On Error GoTo 0
        If Not rng Is Nothing Then
            ASK = CBool(rng.Value)
        End If
    End If

    If ASK Then
        Dim v As Variant
        v = Application.InputBox("Enter quantity:", "Quantity", defaultQty, Type:=1)
        If VarType(v) = vbBoolean And v = False Then
            ' Cancel pressed -> use default
            ResolveQuantity = defaultQty
            Exit Function
        End If
        If IsNumeric(v) Then
            ResolveQuantity = CDbl(v)
        Else
            ResolveQuantity = defaultQty
        End If
    Else
        ResolveQuantity = defaultQty
    End If
End Function

' Duplicate the sheet and give it a suffix OrderSheet_001, _002, ...
Private Sub DuplicateAndArchive(wsTgt As Worksheet)
    Dim baseName As String, newName As String, idx As Long
    baseName = wsTgt.Name & "_"
    idx = 1
    Do
        newName = baseName & Format$(idx, "000")
        If SheetExists(newName) = False Then Exit Do
        idx = idx + 1
    Loop
    wsTgt.Copy After:=wsTgt
    ActiveWorkbook.ActiveSheet.Name = newName
End Sub

Private Function SheetExists(ByVal sheetName As String) As Boolean
    On Error Resume Next
    SheetExists = Not Worksheets(sheetName) Is Nothing
    On Error GoTo 0
End Function

' Clear the main content band on the original sheet (so it can start fresh)
Private Sub ClearContentBand(wsTgt As Worksheet, ByVal startRow As Long, ByVal endRow As Long, _
                             ByVal firstCol As Long, ByVal lastCol As Long)
    wsTgt.Range(wsTgt.Cells(startRow, firstCol), wsTgt.Cells(endRow, lastCol)).ClearContents
End Sub

' Core: copy a single data point from source (row, col) to target (row, col).
' If source empty, writes NULL_PLACEHOLDER.
Public Sub CopyOnePoint( _
    ByVal wsSrc As Worksheet, ByVal wsTgt As Worksheet, _
    ByVal srcRow As Long, ByVal srcCol As Long, _
    ByVal tgtRow As Long, ByVal tgtCol As Long, _
    Optional ByVal nullText As String = NULL_PLACEHOLDER _
)
    Dim v As Variant
    v = wsSrc.Cells(srcRow, srcCol).Value

    If LenB(Trim$(CStr(v))) = 0 Then
        wsTgt.Cells(tgtRow, tgtCol).Value = nullText
    Else
        wsTgt.Cells(tgtRow, tgtCol).Value = v
    End If
End Sub



'==========================================================
' Top-level button macro (EngageSequenceCommand)
'==========================================================
Public Sub EngageSequenceCommand_PLA()
    Dim scrn As Boolean: scrn = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo CleanFail

    Dim wsSrc As Worksheet, wsTgt As Worksheet, wsCaller As Worksheet
    Set wsSrc = Worksheets(SRC_SHEET)
    Set wsTgt = Worksheets(TGT_SHEET)
    Set wsCaller = ActiveSheet  ' button lives on source in this scenario

    ' Resolve caller row (row of the button). If you wire ActiveX, pass explicitly instead.
    Dim callerRow As Long
    callerRow = ResolveCallerRow()
    If callerRow <= 0 Then
        Err.Raise vbObjectError + 513, , "Couldn't resolve the caller row."
    End If

    ' Optional capacity check & auto-archive/clear if needed
    
    
    
    
    
    CheckForSpacesRemaining wsTgt

    
    
    
    
    
    ' Determine quantity
    Dim qty As Double
    qty = ResolveQuantity(1)

    '==========================
    ' LEAD datapoint (PLA F -> OS B)
    '==========================
    Dim leadSrcCol As Long: leadSrcCol = ColNumber("F") ' PLA column F
    Dim leadTgtCol As Long: leadTgtCol = TGT_LEAD_COL   ' OS column B

    Dim leadValue As String
    leadValue = CStr(wsSrc.Cells(callerRow, leadSrcCol).Value)
    If LenB(Trim$(leadValue)) = 0 Then leadValue = NULL_PLACEHOLDER

    ' Decide target row:
    '  - If merging duplicates (by lead datapoint), use existing row; else use next empty row.
    Dim tgtRow As Long
    If ENABLE_MERGE_DUPLICATES And leadValue <> NULL_PLACEHOLDER Then
        tgtRow = FindDuplicateRow(wsTgt, leadValue, leadTgtCol, TGT_STARTROW, TGT_ENDROW)
    End If
    If tgtRow = 0 Then
        tgtRow = NextEmptyRow(wsTgt, TGT_LEAD_COL, TGT_STARTROW, TGT_ENDROW)
        If tgtRow = 0 Then
            Err.Raise vbObjectError + 514, , "No empty row available in target range."
        End If
    End If

    ' If duplicate found -> increment quantity; else write quantity fresh
    If ENABLE_MERGE_DUPLICATES And tgtRow <> 0 Then
        Dim prevQty As Variant
        prevQty = wsTgt.Cells(tgtRow, TGT_QTY_COL).Value
        If IsNumeric(prevQty) Then
            wsTgt.Cells(tgtRow, TGT_QTY_COL).Value = CDbl(prevQty) + qty
        ElseIf LenB(CStr(prevQty)) = 0 Then
            wsTgt.Cells(tgtRow, TGT_QTY_COL).Value = qty
        Else
            ' Non-numeric existed; replace with qty
            wsTgt.Cells(tgtRow, TGT_QTY_COL).Value = qty
        End If
    Else
        wsTgt.Cells(tgtRow, TGT_QTY_COL).Value = qty
    End If

    ' Always write the LEAD datapoint (ensures row anchoring)
    CopyOnePoint wsSrc, wsTgt, callerRow, leadSrcCol, tgtRow, leadTgtCol, NULL_PLACEHOLDER

    '==========================
    ' OTHER datapoints (map PLA -> OrderSheet)
    '==========================
    ' PLA (J -> E) OS
    CopyOnePoint wsSrc, wsTgt, callerRow, ColNumber("J"), tgtRow, ColNumber("E"), NULL_PLACEHOLDER
    
    ' PLA ( E-> D) OS
    CopyOnePoint wsSrc, wsTgt, callerRow, ColNumber("E"), tgtRow, ColNumber("D"), NULL_PLACEHOLDER

    ' PLA (L -> G) OS
    CopyOnePoint wsSrc, wsTgt, callerRow, ColNumber("L"), tgtRow, ColNumber("G"), NULL_PLACEHOLDER

    ' PLA (N -> H) OS
    CopyOnePoint wsSrc, wsTgt, callerRow, ColNumber("N"), tgtRow, ColNumber("H"), NULL_PLACEHOLDER

    ' PLA (O -> J) OS
    CopyOnePoint wsSrc, wsTgt, callerRow, ColNumber("O"), tgtRow, ColNumber("J"), NULL_PLACEHOLDER

    ' Autofit (optional light touch)
    wsTgt.Rows(tgtRow).EntireRow.AutoFit

CleanExit:
    Application.ScreenUpdating = scrn
    Exit Sub

CleanFail:
    Application.ScreenUpdating = scrn
    MsgBox "Sequence error: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

