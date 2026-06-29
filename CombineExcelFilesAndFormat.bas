Sub CombineExcelFilesAndFormat()
    ' =========================================================================
    ' 프로그램명: 엑셀 파일 자동 취합 및 마스킹/강조 프로그램
    ' 작성 목적: 특정 키워드가 포함된 파일의 데이터를 하나로 모으고 보안 및 강조 처리
    ' 깃허브 저장소 업로드용 (상대 경로 적용으로 어디서나 실행 가능)
    ' =========================================================================
    
    Dim wbTarget As Workbook, wbSource As Workbook
    Dim wsTarget As Worksheet, wsSource As Worksheet
    Dim folderPath As String, fileName As String
    Dim nextRow As Long, lastRow As Long, lastCol As Long, i As Long
    Dim isFirstFile As Boolean
    Dim nameCol As Long, nameLen As Integer
    Dim fullName As String, maskedName As String
    Dim sh As Worksheet, rCell As Range
    
    ' 속도 향상 및 팝업 창 차단
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set wbTarget = ThisWorkbook
    
    ' -------------------------------------------------------------------------
    ' 1. '취합' 시트 존재 여부 확인 및 초기화
    ' -------------------------------------------------------------------------
    For Each sh In wbTarget.Worksheets
        If sh.Name = "취합" Then
            Set wsTarget = sh
            Exit For
        End If
    Next sh
    
    ' '취합' 시트가 없으면 맨 앞에 새로 생성
    If wsTarget Is Nothing Then
        Set wsTarget = wbTarget.Worksheets.Add(Before:=wbTarget.Worksheets(1))
        wsTarget.Name = "취합"
    Else
        ' 이미 존재하면 데이터 영역(제목행 제외) 초기화
        lastRow = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row
        If lastRow > 1 Then
            wsTarget.Range("A2:Z" & lastRow).ClearContents
        End If
    End If
    
    ' -------------------------------------------------------------------------
    ' 2. 같은 폴더 내 파일 검색 및 데이터 복사 (원본 수정 방지)
    ' -------------------------------------------------------------------------
    ' 깃허브 공유를 위해 매크로 파일 기준 상대 경로(ThisWorkbook.Path) 사용
    folderPath = wbTarget.Path & "\"
    fileName = Dir(folderPath & "*.xl*")
    isFirstFile = True
    
    Do While fileName <> ""
        ' 자신을 제외하고, 파일명에 '_파일취합_'이 포함된 파일만 필터링
        If fileName <> wbTarget.Name And InStr(fileName, "_파일취합_") > 0 Then
            
            ' 원본 보호를 위해 읽기 전용(ReadOnly:=True)으로 열기
            Set wbSource = Workbooks.Open(folderPath & fileName, ReadOnly:=True)
            
            ' '민원데이터' 시트가 있는지 확인
            Set wsSource = Nothing
            For Each sh In wbSource.Worksheets
                If sh.Name = "민원데이터" Then
                    Set wsSource = sh
                    Exit For
                End If
            Next sh
            
            ' 시트가 존재할 경우 데이터 복사 수행
            If Not wsSource Is Nothing Then
                lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
                
                If lastRow >= 2 Then
                    nextRow = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row + 1
                    
                    ' 첫 파일은 제목행 포함 복사, 이후 파일은 데이터만 복사
                    If isFirstFile Then
                        wsSource.Range("A1", wsSource.Cells(lastRow, wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column)).Copy _
                            wsTarget.Range("A1")
                        isFirstFile = False
                    Else
                        wsSource.Range("A2", wsSource.Cells(lastRow, wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column)).Copy _
                            wsTarget.Range("A" & nextRow)
                    End If
                End If
            End If
            
            ' 원본 파일은 변경사항 없이 닫기
            wbSource.Close SaveChanges:=False
        End If
        fileName = Dir
    Loop
    
    ' -------------------------------------------------------------------------
    ' 3. 개인정보 보호를 위한 '민원인명' 마스킹 처리
    ' -------------------------------------------------------------------------
    On Error Resume Next
    nameCol = Application.WorksheetFunction.Match("민원인명", wsTarget.Rows(1), 0)
    On Error GoTo 0
    
    If nameCol > 0 Then
        lastRow = wsTarget.Cells(wsTarget.Rows.Count, nameCol).End(xlUp).Row
        For i = 2 To lastRow
            fullName = Trim(wsTarget.Cells(i, nameCol).Value)
            nameLen = Len(fullName)
            
            If nameLen = 2 Then
                maskedName = Left(fullName, 1) & "*"
            ElseIf nameLen = 3 Then
                maskedName = Left(fullName, 1) & "*" & Right(fullName, 1)
            ElseIf nameLen = 4 Then
                maskedName = Left(fullName, 1) & "**" & Right(fullName, 1)
            ElseIf nameLen > 4 Then
                maskedName = Left(fullName, 1) & String(nameLen - 2, "*") & Right(fullName, 1)
            Else
                maskedName = fullName
            End If
            wsTarget.Cells(i, nameCol).Value = maskedName
        Next i
    End If
    
    ' -------------------------------------------------------------------------
    ' 4. 중요 키워드(긴급, 민원, 주의) 빨간색 굵은 글씨 강조
    ' -------------------------------------------------------------------------
    lastRow = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row
    lastCol = wsTarget.Cells(1, wsTarget.Columns.Count).End(xlToLeft).Column
    
    If lastRow >= 2 Then
        For Each rCell In wsTarget.Range(wsTarget.Cells(2, 1), wsTarget.Cells(lastRow, lastCol))
            If Not IsError(rCell.Value) Then
                If InStr(rCell.Value, "긴급") > 0 Or _
                   InStr(rCell.Value, "민원") > 0 Or _
                   InStr(rCell.Value, "주의") > 0 Then
                    
                    rCell.Font.Color = RGB(255, 0, 0)
                    rCell.Font.Bold = True
                End If
            End If
        Next rCell
    End If
    
    ' 엑셀 환경 원상복구
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "모든 파일 취합, 마스킹 및 키워드 강조 작업이 완료되었습니다!", vbInformation
End Sub
