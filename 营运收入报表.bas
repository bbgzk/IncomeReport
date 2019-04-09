Attribute VB_Name = "模块1"

Sub 营运收入报表()
'
' 营运收入报表 宏
'

'
    Windows("Kt_Xb_ProductDayData_rep.xls").Activate
    Cells.Select
    Selection.RowHeight = 24.75
    Selection.ColumnWidth = 8.86

    Range("c1") = Range("b2") + Range("i3")
    
    Rows("8:10").Delete
    Rows("2:5").Delete
    Columns("P:R").Delete
    Columns("M:N").Delete
    Columns("K:K").Delete
    Columns("I:I").Delete
    Columns("F:G").Delete
    Columns("A:B").Delete
       
    Range("A2").Select
    Selection.Copy
    Range("A1:G1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlLTR
        .MergeCells = False
    End With
    Selection.Merge
    Selection.Font.Bold = True
    Selection.Font.Size = 12
    Range("A1:G22").Select
End Sub



