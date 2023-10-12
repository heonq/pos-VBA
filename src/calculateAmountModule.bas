Sub increaseAmount(ByRef productName As String)

Dim targetRange As Range
Set targetRange = paymentSheet.UsedRange.Find(productName, lookat:=xlWhole).Offset(1)

targetRange.Value = targetRange.Value + 1

Application.Run ("calculateAmount")

End Sub

Sub decreaseAmount(ByRef productName As String)

Dim targetRange As Range
Set targetRange = paymentSheet.UsedRange.Find(productName, lookat:=xlWhole).Offset(1)

targetRange.Value = targetRange.Value - 1

Application.Run ("calculateAmount")

End Sub
