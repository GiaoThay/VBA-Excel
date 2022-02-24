Sub dancachdong()
'
' Macro5 Macro
'
' Keyboard Shortcut: Ctrl+u
'
    
    Selection.RowHeight = 31
   
End Sub
Sub dancachdong99()
'
' Macro5 Macro
'
' Keyboard Shortcut: Ctrl+u
'
    
    Selection.RowHeight = 16
   
End Sub
Sub XoaThua1dong()
'
' Macro5 Macro
'
' Keyboard Shortcut: Ctrl+u
'

    Selection.Delete Shift:=xlUp
   
End Sub

Sub DoiChuKy()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call DoiChuKy1
    Next
    Application.ScreenUpdating = True
End Sub
Sub DoiChuKy1()
'
' Macro5 Macro
'

'
    Selection.Replace What:="u) Ng", Replacement:="u)" & vbLf & "Ng", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=True
        
    Selection.Replace What:="NG (K", Replacement:="NG" & vbLf & "(K", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=True
    Selection.Replace What:="y... th", Replacement:="y 17 th", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=True
    Selection.Replace What:="ng... n", Replacement:="ng 9 n", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=True
    Selection.Replace What:="m 20...", Replacement:="m 2021", LookAt:=xlPart _
        , SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, _
        ReplaceFormat:=True
End Sub


