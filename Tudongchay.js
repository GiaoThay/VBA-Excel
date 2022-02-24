Sub TuDongChay()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call RunCode
    Next
    Application.ScreenUpdating = True
End Sub
Sub RunCode()
    ' Macro3 Macro
'
' Keyboard Shortcut: Ctrl+j
'
    Columns("A:A").Select
    Selection.ColumnWidth = 3.75
    Columns("B:B").ColumnWidth = 4.29
    Columns("C:C").ColumnWidth = 4.29
    Columns("D:D").ColumnWidth = 4.29
    Columns("E:E").ColumnWidth = 10.71
    Columns("F:F").ColumnWidth = 3.29
    Columns("G:G").ColumnWidth = 6.29
    Columns("H:K").ColumnWidth = 2.57
    Columns("L:M").ColumnWidth = 3.29
    Columns("N:N").ColumnWidth = 8
    Columns("O:O").ColumnWidth = 1.71
    Columns("P:P").ColumnWidth = 6
    Columns("Q:Q").ColumnWidth = 3.29
    Columns("R:R").ColumnWidth = 9.29
    Columns("S:S").ColumnWidth = 3.29
    Columns("T:T").ColumnWidth = 5
    Columns("U:U").ColumnWidth = 7.86
    Columns("V:V").ColumnWidth = 4.43
    Columns("W:W").ColumnWidth = 7.14
    Columns("X:X").ColumnWidth = 3.29
    Columns("Y:Y").ColumnWidth = 8.86
    Columns("Z:Z").ColumnWidth = 3.29
    Columns("AA:AA").ColumnWidth = 3.29
    Columns("AB:AB").ColumnWidth = 3.29
    Columns("AC:AC").ColumnWidth = 5.71
    Columns("AD:AD").ColumnWidth = 9
    Columns("AE:AE").ColumnWidth = 22.57
    Columns("AF:AF").ColumnWidth = 1
    ActiveWindow.SmallScroll Down:=-48
    Rows("9:54").RowHeight = 17.25
    Range("A1:AF53").Select
    With Selection.Font
        .Name = "Times New Roman"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Font
        .Name = "Times New Roman"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
 Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = ""
    Application.PrintCommunication = False
    
End Sub
