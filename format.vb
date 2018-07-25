Sub format()
'
' prettify Macro
' make excel tables readable quickly
'
' Keyboard Shortcut: Ctrl+q
'
    Application.ScreenUpdating = False

    ' turn on filter
    Rows("1:1").AutoFilter
    
    ' auto fit columns using external macro
    Call mAutoWidth
    
    ' wrap header row
    Rows("1:1").WrapText = True
    
    ' freeze header row
    Rows("2:2").Select
    ActiveWindow.FreezePanes = True
    
    ' reset cursor to A1
    Range("A1").Select

    Application.ScreenUpdating = True
End Sub


' https://www.mrexcel.com/forum/excel-questions/655535-how-autofit-column-width-up-maximum-size.html#post3249648
Sub mAutoWidth()
    Dim mCell As Range
    Application.ScreenUpdating = False

    For Each mCell In ActiveSheet.UsedRange.Rows(1).Cells
    mCell.EntireColumn.AutoFit
    If mCell.EntireColumn.ColumnWidth > 40 Then _
    mCell.EntireColumn.ColumnWidth = 40
    Next mCell

    Application.ScreenUpdating = True
End Sub

