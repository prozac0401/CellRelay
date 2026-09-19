# CellRelay

CellRelay는 Excel 파일의 한 열을 위에서 아래로 읽어 웹 작업을 순차 수행하는
Windows 데스크톱 프로그램입니다. 현재 기본 동작은 AWS Skill Builder 교육 상세
페이지에서 Excel의 사용자 식별값을 한 명씩 검색하고 교육에 자동 할당하는
것입니다. 기존의 범용 Text Clear 감지 방식도 선택적으로 유지합니다.

## v0.3.2 간헐 저장 실패와 중지 처리 개선

- Excel 오류 기록의 최종 파일 교체가 Windows 공유/잠금 위반(32/33)으로 실패하면
  검증된 사본의 교체만 최대 5회 시도합니다. 재시도 대기 합계는 최대 1.5초이며,
  매번 원본 변경 여부를 다시 확인합니다. 권한 부족이나 외부 수정은 중지합니다.
- Excel COM이 호출을 명시적으로 거절한 경우(`RPC_E_CALL_REJECTED`,
  `RPC_E_SERVERCALL_RETRYLATER`)만 제한적으로 재시도합니다. 알 수 없는 저장 오류로
  웹 할당이나 Excel 저장 전체를 반복하지 않습니다.
- 임시파일 삭제 실패가 최초 저장 오류를 덮지 않도록 했습니다. 오류 메시지에는
  Excel 저장 실패와 원래 웹 작업 결과를 함께 표시합니다.
- 저장 응답이 늦으면 15초마다 대기 시간을 표시합니다. 로그에는 버전, 프로세스,
  작업 셀, 저장 단계, 소요 시간과 재시도 정보를 남깁니다.
- `완료 / Done` 클릭 후 확인 창이 닫힌 사실을 확인했다면, 화면 안정화 대기 중
  중지해도 완료 버튼 처리 기록을 보존합니다. 실제 서버 등록 성공 검증을 뜻하지는
  않으며, 최종 명단 대조는 계속 필요합니다.

COM 호출 자체가 반환되지 않는 경우의 강제 복구는 아직 지원하지 않습니다. 이 경우
진행 중인 저장과 새 작업이 겹치지 않도록 현재 셀에서 응답을 기다립니다. 별도 프로세스
격리와 시간 제한은 후속 개선 대상으로 남아 있습니다.

## v0.3.1 Excel COM 호환성

Windows에서는 `pywin32`의 COM으로 설치된 Microsoft Excel을 실행하여 파일을
엽니다. `.xlsx`, `.xlsm`, `.xls`, `.xlsb`를 선택할 수 있습니다. 확장자가 `.xlsx`여도
암호화·사내 보안 적용 등으로 실제 파일이 ZIP이 아닌 경우, 같은 PC의 Excel에서
열 수 있는 파일은 COM을 통해 읽도록 처리합니다. 파일을 ZIP으로 직접 해석할 때의
`File is not a zip file` 오류에 의존하지 않습니다.

읽을 열은 시작 셀부터 첫 빈 셀까지 메모리에 보관하고 Excel 파일을 닫습니다.
오류 기록도 읽기에 사용한 방식으로 저장합니다. COM은 작업마다 별도 Excel 인스턴스를
사용하며 사용자가 이미 열어 둔 Excel 창에 연결하거나 그 창을 종료하지 않습니다.
매크로 실행과 외부 링크 갱신은 비활성화합니다. Excel에서 요구하는 암호 입력이나
보안 승인은 자동 처리하지 않으므로 해당 조건이 있으면 파일 열기에 실패할 수 있습니다.

Excel COM이 설치되지 않았거나 시작되지 않으면 일반 `.xlsx`/`.xlsm`에 한해 기존
OOXML 방식으로 읽습니다. ZIP이 아닌 파일에는 데스크톱 Excel이 필요하다는 안내를
표시합니다. COM이 실행된 뒤 파일 열기에 실패한 경우에는 해당 Excel 오류를 표시합니다.

## v0.3.0 화면 구성

창 크기에 맞춰 설정과 실행 현황을 재배치하는 밝은 데스크톱 UI를 사용합니다.
넓은 창에서는 설정과 실행 현황을 좌우로, 좁은 창에서는 위아래로 표시합니다.
높이가 부족하면 본문을 스크롤할 수 있으며 작업 버튼은 하단에 유지됩니다.
창 크기 변경은 입력값, 확인된 과정, 진행 중인 작업을 초기화하지 않습니다.
최소 크기는 560×320이며 최초 실행 크기는 화면의 사용 가능 영역에 맞춰 조정합니다.
1040px 미만에서는 세로 배치로 전환합니다. 크기는 Qt의 논리 픽셀 기준입니다.

- Excel 데이터, 브라우저 연결, 실행 현황을 구분하고 시작 버튼을 강조합니다.
- AWS 모드에서는 확인된 과정을, Text 모드에서는 Text Selector를 표시합니다.
- 상태를 색상과 텍스트로 함께 표시하고, 현재 값과 메시지는 선택해 복사할 수 있습니다.
- 긴 파일 경로와 URL은 입력 영역 안에서 확인하며 창 전체의 너비를 강제로 늘리지 않습니다.
- 스크롤 중 닫힌 Sheet/동작 방식 목록에 마우스가 올라가도 선택값을 바꾸지 않습니다.
- 기존 검색 검사, 완료 버튼 처리, 일시정지/중지, 오류 열 저장 로직은 유지합니다.

디자인은 [Linear의 제품 UI](https://linear.app/plan)와
[Microsoft Fluent의 여백과 레이아웃 지침](https://fluent2.microsoft.design/layout)을
참고했습니다. 브랜드 이미지나 전용 UI 라이브러리는 포함하지 않았습니다.
PySide6 기본 위젯과 공통 스타일을 사용하므로 추가 UI 패키지는 필요하지 않습니다.

## AWS Skill Builder 자동 할당

```text
현재 Excel 셀 읽기
  -> "교육 할당" 클릭
  -> "사용자에 할당" 클릭
  -> 사용자 선택 창의 "사용자 찾기"에 값 입력
  -> Enter로 검색 실행
  -> 검색 요청 완료 또는 로딩 완료 확인 + 결과 안정화 대기
  -> 결과가 정확히 1개이고 검색값과 일치하는지 확인
     -> 불일치: "취소" 클릭, 별도 오류 열에 원인을 저장하고 다음 행
  -> 해당 행 체크
  -> "할당" 클릭
  -> AWS 사용자 등록 확인 창에서 등록 체크박스 선택
  -> "완료" 클릭 후 상세 페이지가 안정될 때까지 잠시 대기
  -> 현재 셀의 완료 버튼 처리 상태 기록 (메인 사용자 표 검사 안 함)
  -> 같은 열의 다음 행으로 이동하고 반복
  -> 다음 행에서 "교육 할당" 버튼을 다시 찾고 활성화 대기
  -> 전체 작업 후 관리자가 최종 등록 명단을 별도로 확보하여 대조
```

검색 완료가 확인된 결과가 0개 또는 2개 이상이거나 이메일/사용자 ID가 정확하게
일치하지 않으면, 아직 `할당`을 누르기 전이므로 사용자 선택 창의 `취소`를 누릅니다.
검색값이 전체 이메일이면 전체 주소를, `사용자ID` 또는 `사용자ID@` 형식이면
이메일의 `@` 앞부분 전체를 비교합니다. `사용자ID@일부도메인` 형식은 허용하지 않습니다.
앞뒤 공백과 대소문자는 무시하지만 부분 문자열이나 이름/팀의 일치는 허용하지 않습니다.
다음 페이지가 활성화된 검색 결과도 한 명으로 판단하지 않습니다.

원본 시트의 사용 영역 오른쪽에 **`CellRelay 검색 오류`** 열을 추가해 같은 행에
시간, 원본 셀 주소, 오류 코드와 사유를 기록합니다. 기존 열을 밀거나 글꼴을 바꾸지
않습니다. 같은 열이 있으면 재사용하고 같은 행 재오류는 최신 사유로 갱신합니다.
성공한 재시도는 기존 오류 이력을 지우지 않으므로 기록된 시각과 진행 로그를 함께
확인하세요. 1행부터 데이터인 경우 오류 열의 1행에 제목과 해당 오류를 함께 적습니다.
UI의
`추가 안 됨` 개수와 `config/progress.json`의 `last_skipped_cell`, `skipped_count`,
`phase`에도 이 상태를 기록합니다. Excel 파일을 저장할 수 없으면 기록 없이 넘어가지
않고 `ERROR`로 중지합니다.

반면 `할당`을 누른 뒤 최종 확인 창 조작이나 팝업 종료 확인에서 실패하면 실제 반영
여부가 불명확하므로 기존처럼 즉시 `ERROR`로 중지하고 Excel 행을 이동하지 않습니다.
페이지에서 실제 할당 여부를 확인한 후 재시작 셀을 결정해야 합니다.

v0.2.2부터 교육 상세 페이지 아래쪽의 `사용자 / Users` 표는 검사하지 않습니다.
이메일이 현재 페이지에 있는지, 다른 페이지에 있는지, `Proxy-enrolled` 등 어떤 상태가
표시되는지는 다음 행 진행과 Excel 오류 기록에 영향을 주지 않습니다.
사용자 선택 **검색 팝업**의 정확한 이메일/ID 검사와 체크박스 검사는 계속 수행합니다.

이 버전의 처리 완료는 등록 체크박스를 선택하고 `Done / 완료`를 누른 뒤 팝업이
닫힌 사실을 확인했다는 뜻입니다. 이후 화면 안정화 대기를 수행하며, 이 대기 중
중지 요청이 오면 확인된 완료 기록을 보존하고 대기를 끝냅니다. AWS 서버의 실제 등록 성공이나 좌석
확보를 검증한 결과가 아닙니다. 전체 작업 후 **관리자가 최종 등록 명단을 별도로
확보해 Excel과 대조**해야 합니다. 완료된 행에는 Excel 성공 표시를 추가하지 않으며,
진행 JSON과 로그에만 완료 버튼 처리 사실을 기록합니다. 기존 오류 이력도 자동 삭제하지
않습니다. 이전 버전이 남긴 `ENROLLMENT_UNCONFIRMED`는 그대로 남을 수 있으나,
현재 버전에서는 메인 표 미표시/등록 상태를 이유로 이 오류를 새로 생성하지 않습니다.

검색 요청 실패(`SEARCH_FAILED`), 완료 미확인/시간 초과(`SEARCH_TIMEOUT`),
최종 확인 창 조작/종료 실패 역시 오류 열에 기록하지만 건너뛰지
않고 현재 행에서 중지합니다. 검색 결과 없음(`NOT_FOUND`), 여러 결과
(`MULTIPLE_MATCHES`), 정확한 ID 불일치(`IDENTITY_MISMATCH`)만 취소 후 진행합니다.

오류 저장은 별도 Excel QThread에서 실행합니다. 첫 수정 직전에 원본 옆에
`원본이름.cellrelay-backup-고유값.원래확장자` 백업을 만들고, 임시 파일 검증 후
원자적으로 교체합니다. COM에서는 Excel의 `SaveCopyAs`로 원래 파일 형식의 사본을
저장하고 Excel로 다시 열어 오류 셀을 검증합니다. 따라서 ZIP이 아닌 파일에도 같은
저장 경로를 사용합니다. Excel이 통합 문서를 다시 저장하므로 파일 내부의 바이트가
동일하게 유지되는 것은 아닙니다. OOXML 대체 경로는 해당 시트 XML의 오류 셀만 수정하여
수식/계산 캐시와 다른 ZIP 구성요소를 보존합니다. 두 방식 모두 Excel에서 파일을
열어 잠갔거나 작업 도중 외부 수정이 감지되면 덮어쓰지 않고 중지합니다.

COM 동작은 Microsoft의 [Workbooks.Open](https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.open)과
[Workbook.SaveCopyAs](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.savecopyas)를 사용합니다.

검색값으로는 결과 행에 그대로 표시되는 이메일 주소 사용을 권장합니다. 이미 할당된
사용자가 검색에서 제외된 경우에도 `NOT_FOUND`로 기록될 수 있으므로 관리자의 최종
명단 대조에 포함하세요. 잔여 좌석 부족 등 실제 등록 결과는 이 프로그램이 검사하지
않습니다.

첫 빈 Excel 셀을 만나면 전체 작업을 완료합니다. COM에서는 Excel이 반환하는 셀 값을
사용하며, 정수 식별값에 `.0`이 붙지 않도록 처리합니다. OOXML 대체 경로는
`data_only=True`로 마지막 수식 계산 캐시를 읽으므로 캐시가 없는 수식은 빈 셀로 보일 수 있습니다.

## 디렉터리 구조

```text
CellRelay/
├─ main.py                         # 프로그램 진입점과 로깅 초기화
├─ app/
│  ├─ ui/main_window.py            # 화면 구성과 Signal 기반 표시 갱신
│  ├─ ui/theme.py                  # 공통 색상, 글꼴, 포커스/상태별 스타일
│  ├─ excel/excel_reader.py        # COM 우선 단일 열 읽기 / OOXML 대체 경로
│  ├─ excel/com_excel.py           # COM 초기화, 값 읽기, 저장 사본 재열기 검증
│  ├─ excel/result_writer.py       # 별도 QThread의 오류 열 저장/백업/원본 보존
│  ├─ browser/browser_worker.py    # QThread에서 실행되는 Playwright 작업
│  ├─ browser/aws_skill_builder.py # AWS 전용 검색/일치 검사/체크/완료 버튼 처리
│  ├─ browser/aws_labels.py        # 한국어/영어 접근성 이름 사전
│  ├─ browser/search_observer.py   # 검색값 관련 요청 완료/실패 관찰
│  ├─ core/controller.py           # 작업 순서, pause/resume/stop, 안전한 행 이동
│  ├─ core/state.py                # 명시적 상태 머신과 화면용 진행 모델
│  ├─ core/logging_config.py       # 회전 로그 파일 설정
│  └─ config/settings.py           # 설정/진행 상태 JSON 원자적 저장
├─ config/
│  ├─ settings.json                # 사용자 설정
│  └─ progress.json                # 실행 중 생성되는 복구용 상태(커밋 제외)
├─ logs/                           # cellrelay.log 생성 위치
└─ tests/                          # Excel, JSON, 상태 머신 단위 테스트
```

UI는 브라우저를 직접 제어하지 않습니다. Controller가 명령 Signal을 보내고,
`BrowserWorker`가 전용 `QThread`에서 Playwright Async API와 전용 asyncio event
loop를 실행한 뒤 결과 Signal만 UI 스레드로 돌려줍니다. pause와 stop은 긴 clear
대기 중에도 처리되도록 thread-safe event로 전달됩니다. AWS 페이지 고유의 접근성 이름과 locator는
`aws_skill_builder.py`와 `aws_labels.py`에 분리되어 일반 브라우저 계층과 결합되지 않습니다.

## 상태 머신

다음 상태가 UI의 실행 상태 영역에 표시됩니다.

- `IDLE`
- `LOADING_EXCEL`
- `READY`
- `INPUTTING`
- `WAITING_FOR_CLEAR`
- `PAUSED`
- `STOPPING`
- `COMPLETED`
- `ERROR`

AWS 자동 할당 반복은 `READY -> INPUTTING -> INPUTTING ... -> COMPLETED`이며,
세부 단계는 마지막 메시지와 `config/progress.json`의 `phase`에 기록됩니다. 범용
Text 모드는 `READY -> INPUTTING -> WAITING_FOR_CLEAR -> INPUTTING`입니다. Pause
시 현재 브라우저 단계를 유지하고 Resume 시 같은 위치부터 계속합니다.
일시정지 시간은 검색/준비 타임아웃에서 제외합니다. 이미 실행 중인 클릭은 되돌릴 수
없지만 다음 단계는 중단됩니다. 완료 버튼 처리 결과가 Pause/Stop과 교차해서 도착해도 먼저
저장하며 다음 작업과 독립적으로 보존합니다. `STOPPING`에서는 이전 브라우저 명령의
종료 응답과 진행 중인 Excel 저장을 기다린 뒤에만 재시작을 허용합니다. 창 닫기도
같은 정리 절차를 거칩니다. 중지 후 시작 셀은 확인/기록 완료된 행만 반영하여 갱신됩니다.

## 설치

Python 3.12 이상이 필요합니다. COM 읽기/저장에는 Windows용 데스크톱 Microsoft
Excel도 설치되어 있어야 합니다. PowerShell에서 다음을 실행합니다.

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m playwright install chromium
```

기본 브라우저는 설치된 Microsoft Edge(`channel="msedge"`)입니다. Edge를 실행할
수 없으면 Playwright Chromium으로 한 번 fallback합니다. fallback을 사용하려면
위의 `playwright install chromium` 단계가 필요합니다.

## 실행

```powershell
python main.py
```

소스 폴더에서는 `run_cellrelay.cmd`를 더블클릭해도 실행할 수 있습니다.

1. `찾기`에서 Excel 파일(`.xlsx`, `.xlsm`, `.xls`, `.xlsb`)을 선택합니다.
2. Sheet와 시작 셀(예: `B5`)을 지정합니다.
3. 동작 방식에서 `AWS Skill Builder 사용자 자동 할당`을 선택합니다.
4. 기본 [AWS 로그인 페이지](https://skillbuilder.aws/login?redirect=https%3A%2F%2Fskillbuilder.aws%2F) URL을 두고 `브라우저 열기`를 누릅니다.
5. 열린 브라우저에서 사용자가 직접 로그인하고, 작업할 과정 상세 페이지로
   이동합니다. CellRelay는 로그인 정보나 인증 화면을 자동 조작하지 않습니다.
6. `AWS 페이지 확인`을 누릅니다. 과정 상세 탭은 최대 15초,
   React의 `교육 할당` / `Assign training` 버튼은 최대 20초 동안 기다립니다.
   성공하면 실제 과정 URL이 `확인된 과정`에 표시됩니다. 로그인 URL을 바꿀 필요가 없습니다.
7. 확인이 성공해 `시작` 버튼이 활성화되면 `시작`을 누릅니다. 이후의 사용자 검색,
   할당, 사용자 등록 확인은 CellRelay가 수행합니다.

`AWS 페이지 확인`이 성공하기 전에는 Excel 값을 브라우저로 보내지 않습니다.
확인 시 고정한 교육 ID와 `orgId`를 작업 중에도 검사하므로 다른 교육/조직으로
이동하면 중지합니다. 이전 버전의 AWS 설정은 최초 실행 시 로그인 URL로 전환됩니다.

로그인 과정에서 새 탭이 열려도 CellRelay가 연 Edge 창의 모든 탭에서 과정 상세
탭 하나를 찾아 제어 대상으로 선택합니다. 과정 상세 탭이 여러 개면 임의 선택하지
않으므로 작업할 탭 하나만 남기세요. 다른 교육을 작업할 때도 해당 페이지를 같은 Edge 창에서
연 뒤 `AWS 페이지 확인`을 다시 실행하면 됩니다. 기존에 따로 열려 있던 개인 Edge
창은 CellRelay가 제어할 수 없습니다.

AWS 모드에서는 Text Selector를 사용하지 않습니다. 프로그램은 한국어/영어 접근성
이름으로 `교육 할당`, `사용자에 할당`, `사용자 선택`, `사용자 찾기`, `사용자`,
`할당` 요소를 찾습니다. 동적으로 생성되는 element ID나 CSS class에는 의존하지
않습니다. `사용자 찾기`는 팝업 내부와 AWS 포털이 페이지 바깥에 렌더링한 영역의
placeholder 및 textbox/searchbox/combobox 역할을 확인하고, 입력 가능해질 때까지
최대 20초 기다립니다. 값을 `fill()`로 입력해 일치 여부를 검증한 뒤 `Enter`를 한 번
눌러 AWS 검색 필터를 적용합니다. 사용자 선택 창에서 `할당`을 누른 뒤에는 AWS의
최종 사용자 등록 확인 창에서는 `선택한 모든 사용자를 등록하고 싶습니다.`
체크박스를 선택하고 활성화된 `완료` 버튼을 누릅니다. 창이 닫힌 뒤에는 1.5초 동안
페이지가 안정되기를 기다립니다. 이후 메인 사용자 표를 조회하지 않고 완료 버튼 처리
기록을 먼저 저장한 다음 Excel 행으로 이동합니다. 다음 행에서 `교육 할당`
버튼을 DOM에서 새로 찾아 표시되고 활성화될 때까지 최대 20초 기다립니다. 이 준비
확인이 실패해도 이전 사람의 완료 버튼 처리 기록은 취소되지 않습니다. 새로 찾은 버튼을
다시 눌러 새 사용자 선택 창에 다음 검색어를 입력합니다. 따라서 남은 모달이나 저장
중인 화면이 다음 행의 클릭을 가로막지 않습니다.

`Text Clear 감지 (수동 Action)`을 선택하면 기존 범용 모드로 동작합니다. 이 모드의
경우 URL과 하나의 요소에만 일치하는 Text Selector를 입력하고 `Text 영역 확인`을
실행합니다. `textarea`와 텍스트 계열 `input`을 `fill()`한 뒤 사용자가 웹 Action을
직접 수행하면, 500ms 동안 안정적으로 empty가 유지될 때 다음 셀로 이동합니다.

Selector가 없거나 여러 요소와 일치하거나, 요소가 숨겨졌거나, disabled/readonly
상태이면 범용 Text 모드의 테스트가 실패합니다. `contenteditable`, iframe 자동
탐색, clipboard/keyboard 입력은 아직 포함하지 않습니다.

## clear 안정성 및 진행 상태

BrowserWorker는 `input`, `change`, DOM mutation을 우선 관찰합니다. React 등의
코드가 이벤트 없이 `value` property를 바꾸는 경우를 위해 50~200ms 간격의
브라우저 내부 보조 확인도 사용합니다. 빈 값이 나타난 뒤 500ms 타이머가 끝나는
시점에 다시 빈 값인지 확인하므로 순간적인 rerender를 완료로 오인하지 않습니다.

`config/progress.json`에는 `workflow_mode`, `current_cell`,
`last_completed_cell`, `last_skipped_cell`, `skipped_count`, `phase`, 처리 개수 등이
원자적으로 저장됩니다. clear 직후에는
`phase: "CLEARED"`, AWS 완료 버튼 처리와 팝업 종료·안정화 직후에는
`phase: "AWS_ASSIGNMENT_SUBMITTED"`, 검색 불일치를 별도 오류 열에 저장한 직후에는
`phase: "AWS_ASSIGNMENT_SKIPPED"`를 행 이동보다 먼저 기록합니다. MVP에는 자동
복구 UI가 아직 없지만, 중복 처리를 판단할 수 있는 정보는 보존됩니다.
추가 필드 `confirmed_training_url`, `terminal_outcome`, `last_error`,
`error_report_backup`은 확인된 과정, 마지막 결과, 오류 및 백업 위치를 보존합니다.
`AWS_ASSIGNMENT_SUBMITTED`는 서버 등록 결과가 아닌 UI 처리 기록입니다.
이전 버전의 `AWS_ASSIGNMENT_CONFIRMED` 진행 기록은 자동으로 바꾸지 않습니다.

## 로그와 테스트

로그는 `logs/cellrelay.log`에 기록되며 5MB 단위로 최대 3개 백업을 유지합니다.
예외에는 stack trace가 포함됩니다.

```powershell
python -m pip install -r requirements-dev.txt
python -m pytest
```

테스트에는 실제 Edge를 headless로 사용하는 한국어/영어 AWS 모의 페이지가 포함됩니다.
메뉴, Enter 검색, 정확한 ID 검사, 취소/다음 사람, 등록 체크박스와 완료 버튼 처리,
메인 사용자 표가 없거나 다른 페이지/대기 상태여도 연속 진행, 다음 버튼 재탐색,
느린 검색/HTTP 오류/타임아웃, Pause/Stop 경합, 오류 저장 실패, 수식 캐시/백업 보존을
검증합니다. 실제 AWS 계정에는 어떤 변경도 하지 않습니다. 실제 로그인된 AWS 양쪽
언어의 운영 화면 검증을 대체하지는 않으며, AWS가 문구/구조를 변경하면 안전하게
중지할 수 있습니다. 처음에는 소량으로 확인한 후 전체 파일에 적용하세요.

배포 EXE에서는 다음 진단 옵션으로 운영과 동일한 QThread에서 브라우저를 연 뒤
후속 Playwright 명령까지 실행하는 회귀 스모크 테스트를 수행할 수 있습니다.

```powershell
.\CellRelay.exe --browser-worker-smoke-test
```

Excel COM과 배포된 pywin32 구성요소는 다음 옵션으로 점검할 수 있습니다.
임시 구형 `.xls` 파일로 읽기, 오류 저장, 백업, 재열기를
확인한 뒤 삭제합니다. 사용자 파일이나 AWS 계정은 수정하지 않습니다.
결과는 로그의 `Packaged Excel COM smoke test passed/failed`에 기록됩니다.

```powershell
.\CellRelay.exe --excel-smoke-test
```

## Windows 실행 파일 빌드

Python 3.12 가상환경 `.venv312`가 준비된 상태에서 다음을 실행합니다.

```powershell
.\build_exe.ps1
```

빌드 스크립트는 `build_windows.py`를 호출하여 자식 빌드 프로세스의 DLL 검색
PATH를 Python/Windows 경로로 제한합니다. 개발 도구(Poppler 등)의 동명 ICU DLL이
배포본에 섞여 Qt 시작을 방해하는 것을 방지하며 시스템 PATH 자체는 변경하지 않습니다.

빌드 결과는 다음 위치에 생성됩니다.

```text
dist/CellRelay.exe
```

PySide6, Playwright, pywin32 실행 구성요소를 포함한 단일 파일이므로 `CellRelay.exe` 하나만
복사해 별도로 배포할 수 있습니다. 실행할 때 구성요소를 Windows 임시 폴더에
자동으로 압축 해제하므로 첫 실행은 폴더형 배포보다 조금 느릴 수 있습니다. 기본
브라우저는 Windows에 설치된 Microsoft Edge입니다. Microsoft Excel 자체는 배포본에
포함되지 않으므로 COM이 필요한 PC에는 별도로 설치되어 있어야 합니다.

EXE에서 생성되는 설정, 진행 상태, 로그는 다음 사용자 폴더에 저장됩니다.

```text
%LOCALAPPDATA%\CellRelay\config
%LOCALAPPDATA%\CellRelay\logs
```

로컬 실사용 점검용 `test.xlsx`/`test.xlsm` 파일은 대소문자와 관계없이 Git에서
제외됩니다. 자동 생성되는 Excel 백업/임시 파일 역시 Git과 배포에서 제외됩니다.
