Sub Macro1()
'
' Macro1 Macro
'
' Keyboard Shortcut: Ctrl+l
Dim value
value = ActiveCell.value
Dim selectedRow
Dim myCell As Range

     ActiveCell.Select
    Selection.Copy
    Sheets("Typical Bath").Select
    Range("A1").Select
    
    Set myCell = Cells.Find(What:=value, After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False)
        
    If myCell Is Nothing Then
        Sheets("Dimension &QTY").Select
        ActiveCell.Select
        Exit Sub
    End If
    
    myCell.Activate
    
selectedRow = ActiveCell.row

    Range("B" & selectedRow & ":AA" & selectedRow).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Dimension &QTY").Select
    ActiveCell.Select
    ActiveSheet.Paste
'
End Sub


Sub Macro2()
'
' Macro2 Macro
'
' Keyboard Shortcut: Ctrl+k
'
Dim value, row, column
    row = ActiveCell.row
    value = ActiveCell.value
    column = ActiveCell.column
    
    If IsEmpty(value) Then
    
        Range(column & (row + 1)).Select
        Exit Sub
    
    End If
    
    
    Application.Run "'CASCADE Pymble.xlsx'!Macro1"
    Cells(column, (row + 1)).Select
    
End Sub
