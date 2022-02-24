Sub TuDongxoathua()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call XoaThua
    Next
    Application.ScreenUpdating = True
End Sub
Sub XoaThua()
'
' Macro5 Macro
'
' Keyboard Shortcut: Ctrl+u
'
If Range("A29").Value = 5 And Range("A34").Value < 1 Then

Rows("44:49").Delete Shift:=xlUp
Rows("44:49").RowHeight = 28.5
ElseIf Range("A24").Value = 4 And Range("A29").Value < 1 Then

Rows("44").Delete Shift:=xlUp
Rows("44:50").RowHeight = 28.5

ElseIf Range("A34").Value = 6 And Range("A39").Value < 1 Then

Rows("44:54").Delete Shift:=xlUp
Rows("44:59").RowHeight = 28.5

ElseIf Range("A39").Value = 7 And Range("A44").Value < 1 Then

Rows("44:59").Delete Shift:=xlUp
Rows("44:55").RowHeight = 28.5
ElseIf Range("A44").Value = 8 And Range("A49").Value < 1 Then

Rows("49:64").Delete Shift:=xlUp
Rows("9:48").RowHeight = 16
Rows("49:54").RowHeight = 28
Else

MsgBox "Sheet: " + ActiveSheet.Name + " can sua bang tay "
ActiveSheet.Tab.color = 255
End If

   
End Sub
