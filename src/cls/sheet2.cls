VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub ballCapSpinButton_SpinDown()

Application.Run "decreaseAmount", productBarcodeDict("NABBAAC01")
Application.Run ("calculateAmount")

End Sub

Private Sub ballCapSpinButton_SpinUp()

Application.Run "increaseAmount", productBarcodeDict("NABBAAC01")
Application.Run ("calculateAmount")

End Sub

Private Sub halfSleeveLargeSpinButton_SpinDown()

Application.Run "decreaseAmount", productBarcodeDict("NABBATP01")
Application.Run ("calculateAmount")

End Sub

Private Sub halfSleeveLargeSpinButton_SpinUp()

Application.Run "increaseAmount", productBarcodeDict("NABBATP01")
Application.Run ("calculateAmount")

End Sub

Private Sub halfSleeveSmallSpinButton_SpinDown()

Application.Run "decreaseAmount", productBarcodeDict("NABBATP02")
Application.Run ("calculateAmount")

End Sub

Private Sub halfSleeveSmallSpinButton_SpinUp()

Application.Run "increaseAmount", productBarcodeDict("NABBATP02")
Application.Run ("calculateAmount")

End Sub

Private Sub slvlessSpinButton_SpinDown()

Application.Run "decreaseAmount", productBarcodeDict("NABBATP03")
Application.Run ("calculateAmount")

End Sub

Private Sub slvlessSpinButton_SpinUp()

Application.Run "increaseAmount", productBarcodeDict("NABBATP03")
Application.Run ("calculateAmount")

End Sub

Private Sub socksSpinButton_SpinDown()

Application.Run "decreaseAmount", productBarcodeDict("NABBAAC02")
Application.Run ("calculateAmount")

End Sub

Private Sub socksSpinButton_SpinUp()

Application.Run "increaseAmount", productBarcodeDict("NABBAAC02")
Application.Run ("calculateAmount")

End Sub

Private Sub stringBagSpinButton_SpinDown()

Application.Run "decreaseAmount", productBarcodeDict("NABBAAC03")
Application.Run ("calculateAmount")

End Sub

Private Sub stringBagSpinButton_SpinUp()

Application.Run "increaseAmount", productBarcodeDict("NABBAAC03")
Application.Run ("calculateAmount")

End Sub

Private Sub windBreakerBlackSpinButton_SpinDown()

Application.Run "decreaseAmount", productBarcodeDict("NABBAOT01")
Application.Run ("calculateAmount")

End Sub

Private Sub windBreakerBlackSpinButton_SpinUp()

Application.Run "increaseAmount", productBarcodeDict("NABBAOT01")
Application.Run ("calculateAmount")

End Sub

Private Sub windBreakerSilverSpinButton_SpinDown()

Application.Run "decreaseAmount", productBarcodeDict("NABBAOT02")
Application.Run ("calculateAmount")

End Sub

Private Sub windBreakerSilverSpinButton_SpinUp()

Application.Run "increaseAmount", productBarcodeDict("NABBAOT02")
Application.Run ("calculateAmount")

End Sub

Private Sub Worksheet_Activate()

barcodeInputRange.Select

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
Application.enableEvents = False

Dim rng As Range

Set rng = ActiveSheet.Range(Target.Address)

If Not Application.Intersect(rng, barcodeInputRange) Is Nothing Then
On Error GoTo safeExit

If productBarcodeDict.exists(Target.Value) Then
Application.Run "increaseAmount", productBarcodeDict(Target.Value)
Else: Application.Run Target.Value
End If
On Error GoTo safeExit

rng.Value = ""
End If

Application.Run ("calculateAmount")

safeExit:
Application.enableEvents = True

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

barcodeInputRange.Select

End Sub
