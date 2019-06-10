Attribute VB_Name = "Module1"
Sub Auto_excel()
    Dim excelApp As Excel.Application
    Set excelApp = New Excel.Application
    excelApp.Workbooks.Add
    excelApp.Visible = True
    excelApp.Range("A1").Select
    excelApp.ActiveCell.FormulaR1C1 = "objeto"
    excelApp.Range("B1").Select
    excelApp.ActiveCell.FormulaR1C1 = "precio"
    excelApp.Range("C1").Select
    excelApp.ActiveCell.FormulaR1C1 = "cantidad"
    excelApp.Range("D1").Select
    excelApp.ActiveCell.FormulaR1C1 = "total"
    excelApp.Range("A2").Select
    excelApp.ActiveCell.FormulaR1C1 = "Mesa"
    excelApp.Range("A3").Select
    excelApp.ActiveCell.FormulaR1C1 = "Silla"
    excelApp.Range("A4").Select
    excelApp.ActiveCell.FormulaR1C1 = "Tv"
    excelApp.Range("A5").Select
    excelApp.ActiveCell.FormulaR1C1 = "Pc"
    excelApp.Range("B2").Select
    excelApp.ActiveCell.FormulaR1C1 = "50000"
    excelApp.Range("B3").Select
    excelApp.ActiveCell.FormulaR1C1 = "100000"
    excelApp.Range("B4").Select
    excelApp.ActiveCell.FormulaR1C1 = "1000000"
    excelApp.Range("B5").Select
    excelApp.ActiveCell.FormulaR1C1 = "2000000"
    excelApp.Range("C2").Select
    excelApp.ActiveCell.FormulaR1C1 = "1"
    excelApp.Range("C3").Select
    excelApp.ActiveCell.FormulaR1C1 = "2"
    excelApp.Range("C4").Select
    excelApp.ActiveCell.FormulaR1C1 = "2"
    excelApp.Range("C5").Select
    excelApp.ActiveCell.FormulaR1C1 = "3"
    excelApp.Range("D2").Select
    excelApp.Application.CutCopyMode = False
    excelApp.ActiveCell.FormulaR1C1 = "=RC[-2]*RC[-1]"
    excelApp.Range("D2").Select
    excelApp.Selection.AutoFill Destination:=Range("D2:D5"), Type:=xlFillDefault
    excelApp.Range("D2:D5").Select
End Sub

