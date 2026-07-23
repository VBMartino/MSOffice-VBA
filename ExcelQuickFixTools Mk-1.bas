Attribute VB_Name = "ExcelQuickFixToolsMk1"
Sub RemoveSkippedBlankRows()
    Dim ws As Worksheet
    Dim inputStr As String
    Dim parts() As String
    Dim startRow As Long, endRow As Long
    Dim r As Long
    Dim lastCol As Long
    
    Set ws = ActiveSheet
    
    ' Prompt for row range (e.g., 17-67)
    inputStr = InputBox("Enter row range (e.g. 17-67):")
    If inputStr = "" Then Exit Sub
    
    parts = Split(inputStr, "-")
    If UBound(parts) <> 1 Then Exit Sub
    
    startRow = CLng(parts(0))
    endRow = CLng(parts(1))
    
    ' Determine last used column for the sheet
    lastCol = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    ' Loop bottom-up and delete blank rows
    For r = endRow To startRow Step -1
        If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol))) = 0 Then
            ws.Rows(r).Delete
        End If
    Next r
End Sub

