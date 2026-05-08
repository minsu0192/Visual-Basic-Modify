' ============================================================
' 샘플 1: 매출 데이터 정리 매크로
' AI에게 "매출 시트의 50000원 초과 행을 표시해줘"로 받은 코드 가정
' ============================================================

Sub 매출데이터정리()
    ' 마지막 행 번호
    Const LAST_ROW = 100

    ' 시트 이름
    Const SHEET_NAME = "매출"

    ' 금액 임계값 (원)
    Const THRESHOLD = 50000

    Dim i As Long
    For i = 2 To LAST_ROW
        If Cells(i, 3).Value > THRESHOLD Then
            Sheets(SHEET_NAME).Range("A1").Value = "확인"
        End If
    Next i

    MsgBox "정리 완료"
End Sub
