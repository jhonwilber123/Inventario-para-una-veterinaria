# Inventario-para-una-veterinaria

hola mundo C#

Sub Macro1()
    '
' Macro1 Macro
'

'
    Sheets("Salidas").Select
    Rows("2:8").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Sheets("FACTURA").Select
    Range("B2").Select
    Selection.Copy
    Sheets("Salidas").Select
    Range("A2:A8").Select
    ActiveSheet.Paste
    Sheets("FACTURA").Select
    Range("A9:A15").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Salidas").Select
    Range("B2").Select
    ActiveSheet.Paste
    Sheets("FACTURA").Select
    Range("E9:E15").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Salidas").Select
    Range("D2").Select
    Application.CutCopyMode = False
    Range("D2:D8").Select
    Sheets("FACTURA").Select
    Selection.Copy
    Sheets("Salidas").Select
    ActiveSheet.Paste
    Sheets("FACTURA").Select
    Range("G7").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Salidas").Select
    Range("F2:F8").Select
    ActiveWindow.SmallScroll Down:=-3
    Range("I2:I8").Select
    ActiveSheet.Paste
    Range("B2:B8").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Selection.EntireRow.Delete
    Sheets("FACTURA").Select
    Range("A1:G20").Select
    Sheets("HISTORIAL DE FACTURACION").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Sheets("FACTURA").Select
    Selection.Copy
    Sheets("HISTORIAL DE FACTURACION").Select
    Range("A1").Select
    ActiveSheet.Paste
    Sheets("FACTURA").Select
    Application.CutCopyMode = False
    Sheets("HISTORIAL DE FACTURACION").Select
    Range("A1:G20").Select
End Sub
