Attribute VB_Name = "ERP_Macros"
Option Explicit

Sub Transaction_Save()
Dim TransRow As Long, TransCol As Long
TransRow = Sheet9.Range("A99999").End(xlUp).Row + 1 'First Avail. Trans Row
For TransCol = 1 To 6
    Sheet9.Cells(TransRow, TransCol).Value = Sheet1.Range(Sheet9.Cells(1, TransCol).Value).Value
Next TransCol
Sheet1.Range("G3,I3,G5,I5,G7:I7").ClearContents
MsgBox "Transaction Saved"
End Sub

Sub Invoice_Save()

End Sub

Sub Purchase_Save()

End Sub
