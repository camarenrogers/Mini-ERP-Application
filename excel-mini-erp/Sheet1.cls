VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
Dim FoundItem As Range
If Target.CountLarge > 1 Then Exit Sub
If Not Intersect(Target, Range("M7:M51")) Is Nothing Then
    Set FoundItem = Sheet8.Range("Items").Find(Target.Value)
    If Not FoundItem Is Nothing Then 'Found
        Range("N" & Target.Row).Value = Sheet8.Range("B" & FoundItem.Row).Value 'Description
        Range("O" & Target.Row).Value = 1 'Default Qty
        Range("P" & Target.Row).Value = Sheet8.Range("E" & FoundItem.Row).Value 'Sale Price
        Range("Q" & Target.Row).Value = "=O" & Target.Row & "*P" & Target.Row 'Total Price
     End If
End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
