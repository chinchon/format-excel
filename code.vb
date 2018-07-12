Sub prettify()
'
' prettify Macro
' make excel tables more readable quickly
'
' Keyboard Shortcut: Ctrl+q
'
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("F6").Select
    ActiveWindow.SmallScroll ToRight:=0
    Rows("1:1").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub
