Sub arrangeSalesHistory()

Dim r As Integer
Dim day As Date

r = salesHistorySheet.UsedRange.rows.Count

If paymentMethodRange.Value = "" Then
MsgBox ("결제방법을 입력해주세요.")
Exit Sub
End If

salesNumberRange.Offset(r).Value = r

For Each iterateKey In productsDict
salesHistorySheet.UsedRange.Find(iterateKey, lookat:=xlWhole).Offset(r).Value = paymentSheet.UsedRange.Find(iterateKey, lookat:=xlWhole).Offset(1).Value
Next iterateKey

salesDateRange.Offset(r).Value = Date
salesTimeRange.Offset(r).Value = time
salesMethodRange.Offset(r).Value = paymentMethodRange.Value
salesTotalAmountRange.Offset(r).Value = paymentTotalAmountRange.Value

If paymentMethodRange.Value = "기타" Then
salesNoteRange.Offset(r).Value = dataNoteRange.Value
dataNoteRange.Value = ""
salesTotalAmountRange.Offset(r).Value = 0
End If

Application.Run ("initQuantity")
paymentLastNumberRange.Value = r

End Sub

Sub calculateAmount()

Dim totalAmount As Double

For Each iterateKey In productsDict
totalAmount = totalAmount + paymentSheet.UsedRange.Find(iterateKey, lookat:=xlWhole).Offset(1).Value * productsDict(iterateKey)
Next iterateKey
paymentTotalAmountRange.Value = totalAmount

End Sub

