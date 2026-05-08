' ============================================================
' 샘플 3: 실제 사용자가 가져온 매크로 (회귀 테스트용)
' 시티은행 KRW 명세서를 같은 폴더에서 찾아 현재 시트로 가져오기
' ============================================================

Sub ImportRimowaCitiKRW()

    ' ===== 1. 변수 선언 =====
    Dim wbSource As Workbook        ' 데이터를 가져올 외부 파일
    Dim wsSource As Worksheet       ' 그 파일의 첫 번째 시트
    Dim wsTarget As Worksheet       ' 지금 매크로 실행 중인 시트
    Dim folderPath As String        ' 현재 파일이 있는 폴더 경로
    Dim fileName As String          ' Dir 함수로 찾을 파일명
    Dim foundFile As String         ' 실제로 찾은 파일의 전체 경로
    Dim lastRow As Long             ' "선택기준"이 나오는 행 번호
    Dim i As Long

    ' ===== 2. 지금 실행 중인 시트(붙여넣을 곳) 지정 =====
    Set wsTarget = ActiveSheet

    ' ===== 3. 같은 폴더에서 "Rimowa Bank statement from Citi KRW"로 시작하는 파일 찾기 =====
    folderPath = ThisWorkbook.Path & "\"
    fileName = Dir(folderPath & "Rimowa Bank statement from Citi KRW*.xls*")

    ' 파일이 없으면 메시지 띄우고 종료
    If fileName = "" Then
        MsgBox "같은 폴더에 'Rimowa Bank statement from Citi KRW'로 시작하는 파일이 없습니다.", vbExclamation
        Exit Sub
    End If

    foundFile = folderPath & fileName

    ' ===== 4. 화면 깜빡임 방지 (속도 향상) =====
    Application.ScreenUpdating = False

    ' ===== 5. 외부 파일 열기 (읽기 전용) =====
    Set wbSource = Workbooks.Open(fileName:=foundFile, ReadOnly:=True)
    Set wsSource = wbSource.Sheets(1)   ' 첫 번째 시트

    ' ===== 6. "선택기준"이라는 글자가 나오는 행 찾기 =====
    ' 14행부터 아래로 내려가면서 A~H열에서 "선택기준" 텍스트를 검색
    lastRow = 0
    For i = 14 To wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row + 100
        Dim j As Long
        For j = 1 To 8   ' A=1, H=8
            If InStr(1, CStr(wsSource.Cells(i, j).Value), "선택기준") > 0 Then
                lastRow = i - 1   ' "선택기준" 바로 위 행까지
                Exit For
            End If
        Next j
        If lastRow > 0 Then Exit For
    Next i

    ' "선택기준"을 못 찾으면 경고
    If lastRow < 14 Then
        MsgBox "'선택기준' 텍스트를 찾지 못했거나 14행 이후에 데이터가 없습니다.", vbExclamation
        wbSource.Close SaveChanges:=False
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' ===== 7. 데이터를 값으로만 복사 =====
    ' 배열로 한 번에 옮기는 게 가장 빠르고 깔끔합니다
    Dim dataArr As Variant
    dataArr = wsSource.Range("A14:H" & lastRow).Value

    ' ===== 8. 현재 시트의 A1부터 값 붙여넣기 =====
    wsTarget.Range("A1").Resize(UBound(dataArr, 1), UBound(dataArr, 2)).Value = dataArr

    ' ===== 9. 외부 파일 닫기 (저장 안 함) =====
    wbSource.Close SaveChanges:=False

    Application.ScreenUpdating = True

    MsgBox "완료! " & (lastRow - 14 + 1) & "행이 복사되었습니다.", vbInformation

End Sub
