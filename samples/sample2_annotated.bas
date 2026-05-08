' ============================================================
' 샘플 2: 어노테이션 사용 — 위젯에 min/max/options 강제 지정
' '@param 라벨 [min=.. max=.. step=.. options=A|B|C]
' ============================================================

Sub 보고서생성()
    ' @param 시작행 min=1 max=10000 step=1
    Const START_ROW = 2

    ' @param 끝행 min=1 max=10000
    Const END_ROW = 500

    ' @param 대상시트 options=매출|매입|재고|인사
    Const SHEET_NAME = "매출"

    ' @param 통과기준
    Const PASS_SCORE = 80

    Dim i As Long
    For i = START_ROW To END_ROW
        If Cells(i, 5).Value >= PASS_SCORE Then
            Sheets(SHEET_NAME).Range("F" & i).Value = "합격"
        Else
            Sheets(SHEET_NAME).Range("F" & i).Value = "미달"
        End If
    Next i

    Dim showAlert As Boolean
    ' @param 알림표시
    showAlert = True

    If showAlert Then
        MsgBox("처리 완료")
    End If
End Sub


Sub 데이터복사()
    ' 원본 시트
    Const FROM_SHEET = "원본"
    ' 대상 시트
    Const TO_SHEET = "결과"

    Sheets(FROM_SHEET).Range("A1:D100").Copy _
        Destination:=Sheets(TO_SHEET).Range("A1")
End Sub
