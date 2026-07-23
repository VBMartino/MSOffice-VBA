Attribute VB_Name = "Module11"
Option Explicit

'==========================
' Global configuration
'==========================
Public Button_TargetSheet As String          ' e.g., "PLA Pipe & Ftgs"
Public Button_TargetColumn As Long           ' e.g., 7 for column G
Public Button_StartRow As Long               ' e.g., 3
Public Button_EndRow As Long                 ' e.g., 123
Public Button_EndRowInclusive As Long            ' e.g., 123+1 inclusive of last line

' Column used for the interim selects in Step 2 (original macro used column H)
Public Button_HelperColumn As Long           ' e.g., 8 for column H

Public Button_ShapeType As MsoAutoShapeType
Public Button_ShapeText As String
Public Button_FontName As String
Public Button_FontSize As Long
Public Button_FontBold As Boolean
Public Button_FontColor As Long
Public Button_ShapeStyle As MsoShapeStyleIndex
Public Button_ShapeWidthPts As Double       ' width in points
Public Button_ShapeHeightPts As Double      ' height in points




Public Button_ClickMacro_Step1 As String     ' Assigned in Step1 (optional)
Public Button_ClickMacro_Step4 As String     ' Assigned in Step4 (authoritative)

' Initialize defaults once per session (call this before the sequence, or from Workbook_Open)
Public Sub Button_InitGlobals()

    Button_TargetSheet = "PLA Pipe & Ftgs"
    Button_TargetColumn = 7                     ' Column A=1,B=2,C=3,D=4,ETC......
    Button_HelperColumn = 8                     ' The helper column exists to serve as a parking spot for the curser
    Button_StartRow = 3                         ' Set's Start row for the Target
    Button_EndRow = 123                         ' Set's End row for the Target
    Button_EndRowInclusive = 1 + Button_EndRow  ' ******DO NOT EDIT*******  Ensures that EndRow is inclusive of the last line when neccesary

    Button_ShapeType = msoShapeFlowchartTerminator
    Button_ShapeText = "+ OrderSheet"                 ' Set's Button text
    Button_FontName = "Arial"
    Button_FontSize = 8
    Button_FontBold = True
    Button_FontColor = RGB(255, 255, 255)
    Button_ShapeStyle = msoShapeStylePreset37
    Button_ShapeWidthPts = 72                   ' 1.0"  (72 points per inch)
    Button_ShapeHeightPts = 15                  ' 0.2"  (14.4 points = 0.2 inches)




    Button_ClickMacro_Step1 = "AddToOS"       ' If you want Step1 to set OnAction
    Button_ClickMacro_Step4 = "EngageSequenceCommand_PLA"   ' Final OnAction applied in Step4
End Sub



Public Sub Step0ClearShapesInTargetColumn()
    Dim ws As Worksheet
    Dim rng As Range
    Dim shp As Shape

    ' Ensure globals are initialized
    If Button_TargetSheet = vbNullString Then Button_InitGlobals

    ' Work on the configured target sheet
    Set ws = ThisWorkbook.Worksheets(Button_TargetSheet)

    ' Build the target range: target column, from Button_StartRow to Button_EndRowInclusive
    Set rng = ws.Range(ws.Cells(Button_StartRow, Button_TargetColumn), _
                       ws.Cells(Button_EndRowInclusive, Button_TargetColumn))

    ' Delete shapes whose top-left cell falls within that range
    For Each shp In ws.Shapes
        If Not shp.TopLeftCell Is Nothing Then
            If Not Intersect(shp.TopLeftCell, rng) Is Nothing Then
                shp.Delete
            End If
        End If
    Next shp
End Sub

' Step 1: create exactly one button in the first target cell (no band loop)
Public Sub Step1CreateStyledShapesMBoMSpec()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim cell As Range
    Dim shapeLeft As Double, shapeTop As Double
    Dim shapeWidth As Double, shapeHeight As Double
    Dim anchorCol As Long, anchorRow As Long

    ' Ensure globals are initialized
    If Button_TargetSheet = vbNullString Then Button_InitGlobals

    Set ws = ThisWorkbook.Worksheets(Button_TargetSheet)
    anchorCol = Button_TargetColumn
    anchorRow = Button_StartRow
    Set cell = ws.Cells(anchorRow, anchorCol)

    ' Remove only shapes anchored to the first target cell
    For Each shp In ws.Shapes
        If Not shp.TopLeftCell Is Nothing Then
            If shp.TopLeftCell.Address = cell.Address Then
                shp.Delete
            End If
        End If
    Next shp

    ' Dimensions (points)
    shapeWidth = Button_ShapeWidthPts
    shapeHeight = Button_ShapeHeightPts

    ' Center in the first target cell
    shapeLeft = cell.Left + (cell.Width - shapeWidth) / 2
    shapeTop = cell.Top + (cell.Height - shapeHeight) / 2

    ' Create the single shape
    Set shp = ws.Shapes.AddShape(Button_ShapeType, shapeLeft, shapeTop, shapeWidth, shapeHeight)

    With shp
        .TextFrame.Characters.Text = Button_ShapeText
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter

        With .TextFrame.Characters.Font
            .Name = Button_FontName
            .Size = Button_FontSize
            .Bold = Button_FontBold
            .Color = Button_FontColor
        End With

        .ShapeStyle = Button_ShapeStyle
        .Name = Button_ShapeText & "_" & anchorRow

        If Len(Button_ClickMacro_Step1) > 0 Then
            .OnAction = Button_ClickMacro_Step1
        End If
    End With
End Sub

' Helper: convert a column number to its letter (e.g., 7 -> "G")
Private Function Button_ColLetter(ByVal colNum As Long) As String
    Button_ColLetter = Split(Cells(1, colNum).Address(False, False), "1")(0)
End Function

' Step 2: keep original behavior (select/zoom/autofill) but parameterize targets
Public Sub Step2ButtonCreation()
    Dim ws As Worksheet
    Dim colLetter As String, helperColLetter As String
    Dim firstCellAddr As String
    Dim destRangeExpr As String
    Dim shapeName As String

    ' Ensure globals are initialized
    If Button_TargetSheet = vbNullString Then Button_InitGlobals

    Set ws = ThisWorkbook.Worksheets(Button_TargetSheet)
    colLetter = Button_ColLetter(Button_TargetColumn)
    helperColLetter = Button_ColLetter(Button_HelperColumn)

    ' Build addresses/names from globals
    firstCellAddr = colLetter & Button_StartRow                         ' e.g., "G3"
    destRangeExpr = colLetter & Button_StartRow & ":" & colLetter & Button_EndRow  ' e.g., "G3:G123"
    shapeName = Button_ShapeText & "_" & Button_StartRow                    ' e.g., "+ O.S._3"

    ' --- Keep the original behavior exactly; only parameterized ---
    ws.Shapes.Range(Array(shapeName)).Select
    ws.Range(helperColLetter & Button_StartRow).Select
    ActiveWindow.Zoom = 250
    ws.Range(firstCellAddr).Select

    ' The critical line you asked to parameterize:
    Selection.AutoFill Destination:=ws.Range(destRangeExpr), Type:=xlFillDefault

    ws.Range(destRangeExpr).Select
    ws.Range(helperColLetter & (Button_EndRow + 1)).Select
    ActiveWindow.Zoom = 100
    ActiveWindow.SmallScroll Down:=-116
    ActiveWindow.ScrollColumn = Button_TargetColumn - 1
    ws.Range(colLetter & "1:" & colLetter & "2").Select
End Sub

Public Sub Step3RenameShapes_InColumnRange()
    Dim ws As Worksheet
    Dim rng As Range
    Dim shp As Shape
    Dim i As Long
    Dim basePrefix As String
    Dim targetName As String

    If Button_TargetSheet = vbNullString Then Button_InitGlobals

    Set ws = ActiveSheet                         ' or ThisWorkbook.Worksheets(Button_TargetSheet)
    Set rng = ws.Range(ws.Cells(Button_StartRow, Button_TargetColumn), ws.Cells(Button_EndRow, Button_TargetColumn))
    basePrefix = Button_ShapeText
    i = 1

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo CleanFail

    For Each shp In ws.Shapes
        If shp.Parent Is ws Then
            If Not shp.TopLeftCell Is Nothing Then
                If Not Intersect(shp.TopLeftCell, rng) Is Nothing Then
                    targetName = GetUniqueShapeName(ws, basePrefix & CStr(i))
                    shp.Name = targetName
                    i = i + 1
                End If
            End If
        End If
    Next shp

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbExclamation, "Rename Shapes"
    Resume CleanExit
End Sub

Private Function GetUniqueShapeName(ByVal ws As Worksheet, ByVal desiredName As String) As String
    Dim candidate As String
    Dim n As Long
    candidate = desiredName
    n = 1
    Do While ShapeNameExists(ws, candidate)
        n = n + 1
        candidate = desiredName & "_" & n
    Loop
    GetUniqueShapeName = candidate
End Function

Private Function ShapeNameExists(ByVal ws As Worksheet, ByVal shapeName As String) As Boolean
    Dim s As Shape
    On Error Resume Next
    Set s = ws.Shapes(shapeName)
    ShapeNameExists = Not s Is Nothing
    Set s = Nothing
    On Error GoTo 0
End Function

Public Sub Step4AssignMacroToShapesInTargetRange()
    Dim ws As Worksheet
    Dim leftX As Double, rightX As Double
    Dim topY As Double, bottomY As Double
    Dim shp As Shape
    Dim cx As Double, cy As Double
    Dim updated As Long, skipped As Long

    If Button_TargetSheet = vbNullString Then Button_InitGlobals

    Set ws = ThisWorkbook.Worksheets(Button_TargetSheet)

    ' Spatial band for the target column & rows
    leftX = ws.Columns(Button_TargetColumn).Left
    rightX = leftX + ws.Columns(Button_TargetColumn).Width
    topY = ws.Cells(Button_StartRow, Button_TargetColumn).Top
    bottomY = ws.Cells(Button_EndRowInclusive, Button_TargetColumn).Top + ws.Cells(Button_EndRowInclusive, Button_TargetColumn).Height

    Application.ScreenUpdating = False
    On Error GoTo CleanFail

    For Each shp In ws.Shapes
        cx = shp.Left + shp.Width / 2
        cy = shp.Top + shp.Height / 2

        If cx >= leftX And cx <= rightX And cy >= topY And cy <= bottomY Then
            On Error Resume Next
            shp.OnAction = Button_ClickMacro_Step4
            If Err.Number = 0 Then
                updated = updated + 1
            Else
                skipped = skipped + 1
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next shp

    Application.ScreenUpdating = True
    MsgBox updated & " shapes assigned to '" & Button_ClickMacro_Step4 & _
           "'.  Skipped: " & skipped & ".", vbInformation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error assigning macros: " & Err.Description, vbExclamation
End Sub

' Orchestrator: runs steps with pauses
Public Sub RunButtonCreationSequence()
    On Error GoTo CleanFail

    If Button_TargetSheet = vbNullString Then Button_InitGlobals

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Running sequence..."

    Button_InitGlobals
    PauseSeconds 0.1
    
    Step0ClearShapesInTargetColumn
    PauseSeconds 0.1

    Step1CreateStyledShapesMBoMSpec     ' only the first target cell
    PauseSeconds 0.1

    Step2ButtonCreation               ' parameterized, preserves original behavior
    PauseSeconds 0.1

    Step3RenameShapes_InColumnRange
    PauseSeconds 0.1

    Step4AssignMacroToShapesInTargetRange

CleanExit:
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    MsgBox "Sequence stopped due to an error: " & Err.Description, vbExclamation, "RunSequence"
    Resume CleanExit
End Sub

Public Sub PauseSeconds(ByVal seconds As Single)
    Dim t As Date
    t = Now + (seconds / 86400#) ' seconds to days
    Application.Wait t
    DoEvents
End Sub

