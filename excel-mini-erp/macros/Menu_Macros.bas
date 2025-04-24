Attribute VB_Name = "Menu_Macros"
Option Explicit
Sub Menu_HideAll()
With Sheet1
    .Shapes("ReportGrp").Visible = msoFalse
    .Shapes("PurchGrp").Visible = msoFalse
    .Shapes("IncExpGrp").Visible = msoFalse
    .Shapes("InvoiceGrp").Visible = msoFalse
    .Shapes("HomeERP").Visible = msoFalse
    .Range("C:AH").EntireColumn.Hidden = True
    ActiveWindow.ScrollColumn = 1
End With
End Sub

Sub Menu_Home()
Menu_HideAll
Sheet1.Shapes("HomeERP").Visible = msoCTrue
Sheet1.Range("C:D").EntireColumn.Hidden = False
End Sub

Sub Menu_IncomeExpense()
Menu_HideAll
Sheet1.Shapes("IncExpGrp").Visible = msoCTrue
Sheet1.Range("E:K").EntireColumn.Hidden = False
End Sub

Sub Menu_Invoice()
Menu_HideAll
Sheet1.Shapes("InvoiceGrp").Visible = msoCTrue
Sheet1.Range("L:R").EntireColumn.Hidden = False
End Sub

Sub Menu_Purchase()
Menu_HideAll
Sheet1.Shapes("PurchGrp").Visible = msoCTrue
Sheet1.Range("S:Y").EntireColumn.Hidden = False
End Sub

Sub Menu_Report()
Menu_HideAll
Sheet1.Shapes("ReportGrp").Visible = msoCTrue
Sheet1.Range("Z:AH").EntireColumn.Hidden = False
End Sub
