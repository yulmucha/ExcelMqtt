# ExcelMqtt
## 목적
엑셀 시트 내 버튼에 매크로 지정할 때 Visual Basic for Applications(VBA)가 활용되는데,  
VBA에서 C# library(.dll)를 참조하여 사용할 수 있도록 하는 것.

## 사소한(?) 문제
사용자가 직접 관리자 권한으로 Powershell을 열어, 작성해 둔 스크립트(Release 폴더와 함께 있는 register_for_COM_interop.ps1)를 찾아 실행해야 함.
그렇게 해야 VBA에서 참조 가능한 .tlb 파일이 만들어짐.

## 설치 및 실행 매뉴얼
1. 라이브러리 파일 준비
2. Powershell(관리자) 실행 및 라이브러리 파일과 함께 있는 스크립트 실행(.dll과 같은 경로에 .tlb 파일 생성됨)
3. 엑셀에서 [파일]-[옵션]-[리본 사용자 지정] 화면으로 이동 후 오른쪽에서 "개발 도구"에 체크 후 [확인]
4. [개발 도구] 탭 클릭-[삽입] 버튼 클릭-"단추(양식 컨트롤)" 선택하여 버튼 추가
5. "매크로 지정" 창에서 매크로 이름 정하고 [새로 만들기] 버튼 클릭(VBA 창이 열림)
6. VBA 창이 뜨면 [도구]-[참조]-[찾아보기]에서 .tlb 파일 선택 후 [열기]
7. VBA의 모듈 코드 편집 창에서 Sub과 End Sub 사이에 라이브러리 사용하는 Visual Basic 코드(글 맨 아래에 첨부) 추가 후 VBA 창 닫기
8. 엑셀로 가서 [파일]-[다른 이름으로 저장], "파일 이름"과 "파일 형식"을 지정하는 창에서 "파일 형식"을 "Excel 매크로 사용 통합 문서 (*.xlsm)"으로 지정 후 저장
(9. 파일을 다시 열었을 때 상단에 보안 경고가 뜨는데 [콘텐츠 사용] 버튼 클릭)

## 라이브러리 사용 Visual Basic 코드
```vb
Dim sheet As Worksheet
Dim tools As ExcelMqtt
Set sheet = ActiveSheet
Set tools = New ExcelMqtt

tools.SetExcelFileName (Application.ActiveWorkbook.FullName)
tools.SetExcelSheetName ("sheet1")
tools.SetBrokerHostName ("esb.myivision.com")
tools.SetBrokerPort (1884)
tools.SetClientId ("ExcelVBA")
tools.SetUsername ("mqttuser9")
tools.SetPassword ("qhdkscjfwj123")
tools.SetTopic ("/excel")
tools.SetPropertyRow (2)
tools.SetStartRecordRow (4)
tools.SetChunkSize (3)
tools.Publish
```

* 파일에 내용이 있어야 함
* 파일이 저장된 파일이어야 함
