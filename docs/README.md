# 포스 프로그램 사용법

## 상품 추가 방법

- initModule에서 productsDict에 key로 상품 이름을, 값에 상품 가격을 입력합니다.
- productBarcodeDict에 key로 상품의 바코드를, 값에 상품 이름을 입력합니다.
- 결제창 시트에 열을 추가하고 상품 이름을 입력합니다.
  - 여기서 입력하는 상품 이름은 모두 일치해야 합니다.
- 디자인 모드를 클릭하고 기존의 스핀버튼을 복사해서 상품 이름이 입력된 셀을 선택하고 붙여넣습니다.

- 스핀버튼을 우클릭하고 속성을 눌러 스핀버튼의 이름을 편집합니다.
- 스핀버튼을 더블 클릭하면 코드 편집창이 표시됩니다.
- 아래 코드를 스핀버튼의 이름과 상품 바코드를 입력해서 코드 편집기의 sheet2(결제창) 추가해줍니다.

```vba
Private Sub 스핀버튼이름_SpinDown()
Application.Run "decreaseAmount", productBarcodeDict(바코드)
Application.Run ("calculateAmount")
End Sub

Private Sub 스핀버튼이름_SpinUp()
Application.Run "increaseAmount", productBarcodeDict(바코드)
Application.Run ("calculateAmount")
End Sub
```

- 대시보드 시트에 열을 추가하고 상품명을 입력합니다.
- 다른 상품 열에 있는 판매 수량을 구하는 수식을 복사 붙여넣기 합니다.

## 결제 방법

- 스핀버튼을 누르거나 바코드를 입력해서 상품 수량을 입력합니다.
- 결제방법 버튼을 누르거나 바코드를 입력해서 결제 방법을 입력합니다.
- 결제완료 버튼을 누르거나 바코드를 입력해서 결제를 완료합니다.
- 바코드를 입력하기 위해선 A3 셀이 선택된 상태이고 셀 고정이 켜져있어야 합니다.
- 칸을 이동했을 때 다시 A3 셀이 선택된다면 셀이 고정되어 있는 상태입니다.
- 기타 결제를 선택할 경우 기타 사유를 입력합니다.
- 기타 결제의 경우 판매액 합계에 포함되지 않습니다.

## 바코드 생성

- 상품 바코드는 원하는 바코드를 입력하면 됩니다.
- 결제 방법,결제 완료 바코드는 해당 Sub 이름을 입력해야만 정상적으로 작동됩니다.
- 만약 원하는 바코드가 있다면 Sub 이름을 해당 바코드에 맞게 수정합니다.

## init

- init은 public으로 선언한 변수들에 값을 할당합니다.
- 엑셀파일을 실행하면 기본적으로 init이 실행되지만, 코드를 수정하면 다시 init을 실행해야 합니다.
- 결제창에 있는 초기화 버튼을 누르면 init이 실행됩니다.