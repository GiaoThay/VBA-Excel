Sub bokhong1()
'
' bokhong Macro
'

'
    
    ActiveWindow.DisplayZeros = False
End Sub
Sub bokhong()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call bokhong1
    Next
    Application.ScreenUpdating = True
End Sub
