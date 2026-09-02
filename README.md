# CellRelay

CellRelay는 `.xlsx` 파일의 한 열을 위에서 아래로 읽어 웹 작업을 순차 수행하는
Windows 데스크톱 프로그램입니다. 현재 기본 동작은 AWS Skill Builder 교육 상세
페이지에서 Excel의 사용자 식별값을 한 명씩 검색하고 교육에 자동 할당하는
것입니다. 기존의 범용 Text Clear 감지 방식도 선택적으로 유지합니다.

## AWS Skill Builder 자동 할당

```text
현재 Excel 셀 읽기
  -> "교육 할당" 클릭
  -> "사용자에 할당" 클릭
  -> 사용자 선택 창의 "사용자 찾기"에 값 입력
  -> Enter로 검색 실행
  -> 결과 목록이 안정될 때까지 대기
  -> 결과가 정확히 1개이고 검색값과 일치하는지 확인
  -> 해당 행 체크
  -> "할당" 클릭
  -> 메인 사용자 표에서 할당 결과 확인
  -> 현재 셀 완료 상태 기록
  -> 같은 열의 다음 행으로 이동하고 반복
```

검색 결과가 0개 또는 2개 이상이거나, 결과 행에 검색값이 포함되지 않거나, 할당 후
메인 사용자 표에서 결과를 확인할 수 없으면 즉시 `ERROR`로 중지합니다. 이 경우
Excel 행은 이동하지 않습니다. 이미 실제 할당이 수행됐지만 결과 확인만 실패했을
가능성이 있으므로, 해당 사용자의 할당 여부를 페이지에서 확인한 후 재시작 셀을
결정해야 합니다.

검색값으로는 결과 행에 그대로 표시되는 이메일 주소 사용을 권장합니다. 교육의
잔여 좌석이 부족하거나 이미 할당된 사용자가 검색 대상에서 제외되는 경우에도
작업은 현재 셀에서 안전하게 중지됩니다.

첫 빈 Excel 셀을 만나면 전체 작업을 완료합니다. Excel 수식 셀은
`data_only=True`로 열어 파일에 저장된 마지막 계산 결과를 사용합니다. 계산 결과
캐시가 없는 수식은 빈 셀로 보일 수 있습니다.

## 디렉터리 구조

```text
CellRelay/
├─ main.py                         # 프로그램 진입점과 로깅 초기화
├─ app/
│  ├─ ui/main_window.py            # 화면 구성과 Signal 기반 표시 갱신
│  ├─ excel/excel_reader.py        # openpyxl 단일 열 순차 읽기
│  ├─ browser/browser_worker.py    # QThread에서 실행되는 Playwright 작업
│  ├─ browser/aws_skill_builder.py # AWS 전용 메뉴/검색/체크/할당/검증
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
`aws_skill_builder.py`에만 있어 일반 브라우저 계층과 결합되지 않습니다.

## 상태 머신

다음 상태가 UI 하단에 표시됩니다.

- `IDLE`
- `LOADING_EXCEL`
- `READY`
- `INPUTTING`
- `WAITING_FOR_CLEAR`
- `PAUSED`
- `COMPLETED`
- `ERROR`

AWS 자동 할당 반복은 `READY -> INPUTTING -> INPUTTING ... -> COMPLETED`이며,
세부 단계는 마지막 메시지와 `config/progress.json`의 `phase`에 기록됩니다. 범용
Text 모드는 `READY -> INPUTTING -> WAITING_FOR_CLEAR -> INPUTTING`입니다. Pause
시 현재 브라우저 단계를 유지하고 Resume 시 같은 위치부터 계속합니다.

## 설치

Python 3.12 이상이 필요합니다. PowerShell에서 다음을 실행합니다.

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

1. `찾기`에서 `.xlsx` 파일을 선택합니다.
2. Sheet와 시작 셀(예: `B5`)을 지정합니다.
3. 동작 방식에서 `AWS Skill Builder 사용자 자동 할당`을 선택합니다.
4. AWS Skill Builder 교육 상세 URL을 입력하고 `브라우저 열기`를 누릅니다.
5. 열린 브라우저에서 사용자가 직접 로그인하고, 입력한 교육 상세 페이지로
   이동합니다. CellRelay는 로그인 정보나 인증 화면을 자동 조작하지 않습니다.
6. `AWS 페이지 확인`을 눌러 실제 브라우저 URL과 `교육 할당` 버튼을 확인합니다.
7. 확인이 성공해 `시작` 버튼이 활성화되면, `시작`을 누르고 자동 할당 확인 창에서
   승인합니다.

`AWS 페이지 확인`이 성공하기 전에는 Excel 시작 셀을 읽거나 브라우저로 보내지
않습니다. 현재 브라우저의 교육 ID 또는 `orgId`가 입력한 URL과 다르면 확인이
실패하므로, 다른 교육이나 조직에 잘못 할당하는 것을 방지합니다.

로그인 과정에서 새 탭이 열려도 CellRelay가 연 Edge 창의 모든 탭에서 현재 UI에
입력된 교육 ID와 `orgId`가 정확히 같은 탭을 찾아 자동으로 제어 대상으로
선택합니다. 교육 URL을 바꾼 경우에도 해당 페이지를 같은 CellRelay Edge 창에서
연 뒤 `AWS 페이지 확인`을 다시 실행하면 됩니다. 기존에 따로 열려 있던 개인 Edge
창은 CellRelay가 제어할 수 없습니다.

AWS 모드에서는 Text Selector를 사용하지 않습니다. 프로그램은 한국어/영어 접근성
이름으로 `교육 할당`, `사용자에 할당`, `사용자 선택`, `사용자 찾기`, `사용자`,
`할당` 요소를 찾습니다. 동적으로 생성되는 element ID나 CSS class에는 의존하지
않습니다. `사용자 찾기`는 placeholder와 textbox/searchbox/combobox 역할을
순서대로 확인하며, 값을 `fill()`로 입력해 일치 여부를 검증한 뒤 `Enter`를 한 번
눌러 AWS 검색 필터를 적용합니다.

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
`last_completed_cell`, `phase`, 처리 개수 등이 원자적으로 저장됩니다. clear 직후에는
`phase: "CLEARED"`, AWS 할당 결과 확인 직후에는
`phase: "AWS_ASSIGNMENT_CONFIRMED"`를 행 이동보다 먼저 기록합니다. MVP에는 자동
복구 UI가 아직 없지만, 중복 처리를 판단할 수 있는 정보는 보존됩니다.

## 로그와 테스트

로그는 `logs/cellrelay.log`에 기록되며 5MB 단위로 최대 3개 백업을 유지합니다.
예외에는 stack trace가 포함됩니다.

```powershell
python -m pip install -r requirements-dev.txt
python -m pytest
```

테스트에는 실제 Edge를 headless로 사용한 AWS 모의 페이지 통합 테스트가 포함됩니다.
이 테스트는 메뉴, 사용자 검색, 단일 결과 검사, checkbox, 할당, 메인 표 검증 전체
흐름과 다중 검색 결과 거부 동작을 확인하며 실제 AWS 계정에는 어떤 변경도 하지
않습니다.

배포 EXE에서는 다음 진단 옵션으로 운영과 동일한 QThread에서 브라우저를 연 뒤
후속 Playwright 명령까지 실행하는 회귀 스모크 테스트를 수행할 수 있습니다.

```powershell
.\CellRelay.exe --browser-worker-smoke-test
```

## Windows 실행 파일 빌드

Python 3.12 가상환경 `.venv312`가 준비된 상태에서 다음을 실행합니다.

```powershell
.\build_exe.ps1
```

빌드 결과는 다음 위치에 생성됩니다.

```text
dist/CellRelay.exe
```

PySide6와 Playwright 실행 구성요소를 포함한 단일 파일이므로 `CellRelay.exe` 하나만
복사해 별도로 배포할 수 있습니다. 실행할 때 구성요소를 Windows 임시 폴더에
자동으로 압축 해제하므로 첫 실행은 폴더형 배포보다 조금 느릴 수 있습니다. 기본
브라우저는 Windows에 설치된 Microsoft Edge입니다.

EXE에서 생성되는 설정, 진행 상태, 로그는 다음 사용자 폴더에 저장됩니다.

```text
%LOCALAPPDATA%\CellRelay\config
%LOCALAPPDATA%\CellRelay\logs
```

로컬 실사용 점검용 `test.xlsx`/`test.xlsm` 파일은 대소문자와 관계없이 Git에서
제외됩니다. 배포 파일에는 포함되지 않습니다.
