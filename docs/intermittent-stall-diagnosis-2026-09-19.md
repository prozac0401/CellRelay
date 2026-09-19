# 간헐 정지 / Excel 저장 점유 진단

조사일: 2026-09-19  
분석 기준: 커밋 `0c3ea21` (v0.3.1 코드), 조사 시작 시 작업 트리 변경 없음.

## 후속 개선: v0.3.2

아래 본문은 수정 전 진단을 보존한 것이다. 사용자 후속 설명은 “당시 셀 위치를
그대로 두고 다시 실행하니 작업이 이어졌다”이며, 중지/재개 버튼을 조작했는지는
확인되지 않았다. 완료 직후 중지 경합은 이번 사건의 확정 원인이 아니다.

후속 요청에 따라 일시적인 파일 교체 잠금과 명시적 COM 호출 거절에만 제한된
재시도를 추가하고, 정리 오류에 의한 최초 오류 가림을 방지했다. 저장 경과 시간,
원래 작업 결과, 세부 단계 로그를 보완하고 완료 클릭·팝업 종료 관찰 후 중지와
완료 기록이 교차하는 문제도 수정했다. 원본 변경 감지와 백업은 유지한다.

COM 호출 자체의 무응답 강제 복구, 웹 작업 전체 자동 재실행, Excel 반영 없이
다음 셀로 이동하는 정책 변경, 중복 실행 방지는 이번 변경에 포함하지 않았다.

구현 후 확인한 결과:

- Controller·파일 저장·COM 회귀 테스트: `51 passed, 6 deselected`.
  실제 Excel 실행을 반복하는 native 테스트 6개는 이 실행에서 제외했다.
- 별도 임시 `.xlsx`에 실제 Excel COM으로 오류를 기록한 뒤 openpyxl로 읽어
  오류 셀, 기존 입력값, 수식, 원본 백업과 임시파일 정리를 확인했다. 통과했으며
  약 71초가 소요됐다. 사용자 원본이나 실제 AWS 할당은 변경하지 않았다.
- 변경한 Python 파일의 Ruff 검사와 `git diff --check`를 통과했다.
- 단일 Edge의 로컬 합성 페이지에서 브라우저 기능 6개를 확인했다. 완료 클릭 전
  중지, 확인 창 종료 관찰 전 중지는 완료 처리하지 않았고, 종료 관찰 후
  중지·일시정지 후 중지·일시정지는 완료 신호를 정확히 한 번 보존했다.
  정상 실행의 안정화 대기도 유지됐다. 기능 검증 6/6 통과, 전체 프로세스 종료 코드 0.
  다만 검증 정리 단계의 `browser.close()`는 30초 제한에 도달했고 이후
  `driver.stop()`이 성공했다. 브라우저 종료 지연까지 해결됐다는 뜻은 아니다.
  증거: `build/browser_settle_harness.log`.
- 최종 `dist-v0.3.2/CellRelay.exe`의 수정 모듈 6개와 `main`을 추출해 현재 소스와
  코드 객체를 대조했다(경로만 정규화). 모두 일치했고 내장 버전은 `0.3.2`였다.
- 별도 임시 설정 경로로 최종 EXE의 `--smoke-test`를 실행했다. 시작·종료 로그와
  종료 코드 0을 확인했다. 사용자 설정 파일은 변경하지 않았다.

최종 EXE: 87,363,969 bytes. SHA256:
`88c9b5796c1da00d66cfe23cc2fdfa336b3b1360db2e80e4d5b6d093effc419c`.

## 결론과 확인 범위

“중간 셀에서 멈추지만 같은 셀부터 다시 시작하면 정상 진행”이라는 증상은 일시적인 Excel COM 응답 거절이나 파일 잠금과 부합한다. 코드에는 이 상황에서 자동 복구하지 못하는 경로가 있다. 특히 파일 교체가 한 번만 실패해도 중지하고, COM 호출이 반환되지 않으면 저장뿐 아니라 중지/종료도 계속 기다릴 수 있다.

다만 이번 발생 시점의 로그를 확보하지 못했으므로 실제 원인이 Excel, 보안 프로그램, 외부 파일 열기, 웹 검색 지연 중 무엇인지는 확정하지 않았다. 사용자가 기억하는 Writing/프로세스 점유 문구의 정확한 원문도 확인하지 못했다. 현재 AppState에는 `WRITING`이라는 상태가 없으며, Excel 저장 중에는 별도 메시지를 표시한다.

확인한 자료:

- 저장소 `logs/cellrelay.log`: 마지막 기록 2026-09-05, 주로 실행/스모크 테스트.
- EXE 저장 경로 `%LOCALAPPDATA%/CellRelay/logs/cellrelay.log`: 마지막 기록 2026-09-03. 과거 AWS 검색 입력 영역을 찾지 못한 오류는 있지만 이번 Excel 저장 장애의 증거는 없다.
- EXE 경로의 `config/progress.json`: 2026-09-02의 과거 오류 기록. 이번 작업 진행 상태로 해석하지 않았다.
- 조사 시 CellRelay 프로세스는 발견되지 않았고, 창 제목이 없는 Excel 프로세스 하나가 있었다. 소유 관계와 점유 파일을 확인하지 못했으므로 이를 잔류 프로세스나 원인으로 단정하지 않았다.

실행본이 이 소스와 같은 버전인지도 확인이 필요하다. v0.3.1 이전 실행본에는 COM 관련 판단을 그대로 적용할 수 없다.

## 1. 일시 파일 잠금을 즉시 실패로 처리

관련 코드: `app/excel/result_writer.py:310`, `app/core/controller.py:553`.

오류 저장은 임시 사본 생성·검증 후 `os.replace(temp_path, path)`로 원본을 교체한다. 이 교체가 `PermissionError`를 내면 재시도하지 않고 저장 실패를 전달한다. Controller는 해당 셀에서 중지한다. 파일이 외부에서 변경된 경우에도 원본 보호를 위해 중지하며, 이 보호는 유지해야 한다.

잠금 주체는 사용자가 연 Excel일 수도 있고, 다른 프로세스일 수도 있다. 구체적 주체는 현재 자료로 알 수 없다. 원본 교체 잠금과 Excel COM의 busy는 서로 다른 오류다.

재현: 임시 OOXML 통합 문서에 첫 번째 파일 교체만 실패하고 이후 성공하도록 모의했다.

```text
TRANSIENT_FIRST: PermissionError replace_calls=1 original_preserved=True temps=0
MANUAL_REPEAT: success=True replace_calls=2 backup_preserved=True backup_count=1
```

일시 점유가 해제되어도 첫 실행은 실패하고, 수동 재실행은 성공하는 경로를 확인했다. 이는 실제 사용자 사건의 재현이 아니라 코드의 장애 처리 재현이다.

## 2. COM 호출 무응답에 대한 시간 제한 없음

관련 코드: `app/excel/com_excel.py:57`, `app/excel/com_excel.py:91`, `app/excel/com_excel.py:187`, `app/excel/result_writer.py:344`.

Windows의 자동 선택은 일반 `.xlsx`도 COM을 우선 사용한다. 오류를 기록할 때 전용 Excel 인스턴스를 시작하고 원본 열기 → 셀 수정 → SaveCopyAs → 사본 재열기/검증 → 닫기/종료를 수행한다. 호출 및 정리 단계에 제한 시간과 busy 재시도가 없다.

ExcelWorker는 쓰기 함수가 반환되거나 예외를 던져야 완료 신호를 보낸다. 호출이 돌아오지 않으면 `_excel_pending`이 유지되고, Controller의 `_finish_stop_if_ready()`도 끝나지 않는다. QThread 분리는 UI 작업을 분리하지만, COM 호출의 강제 중단을 보장하지 않는다.

모의 검증에서는 저장 완료 신호만 보류했다. 브라우저의 중지 응답이 도착해도 현재 셀에서 STOPPING을 유지했고, 종료 요청도 기다렸다. 저장 실패 응답을 주입하자 ERROR 처리 및 종료가 진행됐다.

```text
withheld_writer_reply: STOPPING browser_ack=True excel_pending=True cell=B5
shutdown_waits: True
after_injected_failure: ERROR shutdown_started=True
```

이는 실제 Excel 무응답을 유발한 실험이 아니라 완료 신호 미도착의 영향을 확인한 실험이다. 코드에는 이 대기를 끝내는 watchdog이 없다.

Microsoft는 Office COM 호출이 다른 처리나 모달 대화상자 때문에 대기/거절될 수 있음을 설명한다. `DisplayAlerts=False`도 보안 경고에는 적용되지 않는다. 현재 코드는 경고·이벤트·매크로 억제와 빈 암호 지정 등 완화 조치를 이미 사용하지만, 무응답 가능성을 없애는 것은 아니다.

## 3. 정리 오류가 최초 원인을 가릴 수 있음

관련 코드: `app/excel/result_writer.py:306`, `app/excel/com_excel.py:80`, `app/excel/com_excel.py:118`.

COM 저장 도중 실패한 뒤 `finally`에서 임시파일을 삭제한다. 이 삭제도 점유로 실패하면 UI에는 처음의 COM 실패 대신 마지막 PermissionError가 전달될 수 있다. Excel 닫기/종료 예외는 로그만 남기므로 열린 핸들이 남는 경우도 함께 조사해야 한다.

모의 검증에서 COM 저장에 `RuntimeError('original COM busy')`, 임시파일 삭제에 PermissionError를 주입했다. 호출자에게 보이는 예외는 `PermissionError: temporary file is occupied`였고, 최초 RuntimeError는 예외의 context에만 남았다. 실제 로그 traceback에는 연결된 예외가 남을 수 있지만, UI는 `str(exc)`만 사용한다.

따라서 “다른 프로세스가 사용 중” 문구만으로 원본 파일 잠금인지 임시파일 정리 실패인지 판단하면 안 된다.

## 4. Excel 쓰기는 웹 오류의 후속 단계일 수 있음

관련 코드: `app/core/controller.py:515`, `app/core/controller.py:521`, `app/core/controller.py:635`, `app/browser/aws_skill_builder.py:69`.

정상 완료한 모든 행을 Excel에 쓰는 구조가 아니다. 검색 결과 없음·불일치 등으로 건너뛸 때 또는 브라우저 작업 실패 시 오류 열을 저장한다. 웹 검색 제한 시간은 기본 15초, 준비 확인은 기본 20초이며, 완료가 확인되지 않으면 현재 행에서 중지한다.

가능한 연쇄는 다음과 같다.

```text
웹 검색 지연/오류 또는 검색 결과 불일치
→ Excel 오류 기록 시작
→ COM 대기 또는 파일 교체 실패
→ 같은 셀에서 중지
→ 잠시 후 같은 셀 재시작 시 정상
```

웹 지연이 최초 원인이었는지, Excel 저장이 실제로 멎었는지는 작업 단계와 오류 원문이 필요하다. 할당/완료 버튼을 누른 뒤 결과가 불명확한 행은 웹 작업 전체를 자동 재실행하면 안 된다.

## 5. 완료 직후 중지하면 완료 기록을 놓치는 별도 경합

관련 코드: `app/browser/aws_skill_builder.py:388`, `app/browser/aws_skill_builder.py:539`, `app/browser/browser_worker.py:423`.

Done 클릭과 팝업 종료 후 1.5초 안정화 대기 중 중지 요청이 들어오면 `AwsAssignmentCancelled`가 발생하여 완료 신호가 전달되지 않는다. 실제 제출은 끝났지만 현재 셀에 남을 수 있다.

로컬 headless 합성 페이지에서 Done 처리와 팝업 종료가 완료된 후 대기 진입 시 중지를 주입했다. 제출 목록에 대상이 들어 있고 확인 창은 닫힌 상태였지만 결과는 AwsAssignmentCancelled였다. 실제 AWS 서버의 등록 성공을 검증한 실험은 아니다.

이는 보고된 자동 정지의 직접 원인으로 확인된 사항은 아니다. 다만 멈춰 보이는 상황에서 중지 후 같은 셀을 다시 시작하는 복구 동작과 관련되므로, 최종 클릭·팝업 종료 사실을 안정화 대기보다 먼저 기록하고 중지와 교차해도 보존하도록 개선할 필요가 있다. 실제 서버 등록 여부는 기존 정책대로 별도 확인 대상이다.

## 개선 우선순위

### 우선 1: 오류 분류·재시도·진단

- 원래 업무 오류와 Excel 저장 오류, 임시파일 정리 오류를 별도로 보존한다. 정리 실패가 최초 예외를 덮지 않도록 한다.
- 파일 공유 위반과 검증된 일시적 COM busy 오류에만 지연 및 총시간 제한을 둔 재시도를 적용한다. 권한 부족, 외부 파일 수정, 잘못된 입력 등을 무조건 반복하지 않는다.
- 파일 교체를 재시도할 때마다 원본 signature를 확인하고, 원본이 변경되었으면 중지한다. 기존 백업/검증/원본 보호 동작은 유지한다.
- 로그에 버전, 실행 ID, PID, 셀, backend, 세부 단계 시작/끝, 소요 시간, 재시도 횟수, HRESULT/WinError를 남긴다. COM 시작·원본 열기·사본 저장·검증·닫기·파일 교체·정리를 구분한다.
- UI에 저장 단계, 경과 시간, 재시도 상황을 표시한다. 현재의 포괄적인 “저장 중”만으로는 장시간 작업과 무응답을 구분하기 어렵다.
- 최종 Done 처리·팝업 종료 직후의 확인된 사실은 안정화 대기 중 중지되더라도 보존한다. 제출 이전 재시도와 제출 이후 결과 확인을 구분한다.

### 우선 2: 반환하지 않는 COM 호출의 격리

- COM 작업을 별도 작업 프로세스로 옮기고 부모에서 제한 시간을 감시한다. 재시도만으로는 반환하지 않는 호출을 해결할 수 없다.
- 자식 프로세스는 사본 생성·검증을 맡고 원본 최종 교체는 부모가 맡는 구조가 적절하다. timeout 이후 이전 작업이 늦게 원본을 덮는 경합을 방지한다.
- 앱이 생성한 전용 Excel의 PID와 소유 관계를 추적하여 정리한다. 모든 EXCEL.EXE를 일괄 종료하는 방식은 사용하지 않는다.
- Qt 스레드를 강제 종료하거나 `_excel_pending`만 해제해 다음 행을 진행하는 것은 해결책이 아니다. 살아 있는 이전 저장 작업과 새 작업이 겹칠 수 있다.
- 시작/파일 읽기의 COM 작업도 현재 UI 스레드에서 호출되므로 같은 격리 구조에 포함한다.

### 우선 3: 업무 결과 기록과 원본 Excel 기록 분리 검토

- 행별 결과를 먼저 앱 전용 내구성 기록에 저장하고, 원본 Excel 반영을 별도 단계로 관리하면 원본 파일 점유가 웹 작업 복구에 미치는 영향을 줄일 수 있다.
- 다만 현재 정책은 Excel 오류 기록 완료 전에는 행을 넘기지 않는 것이다. “나중에 Excel에 반영하고 계속 진행”은 이 정책을 바꾸므로 미반영 내역 표시·재처리·원본 변경 충돌 처리를 함께 설계해야 한다.
- 공유 설정/진행 JSON도 고정 `.tmp` 이름을 사용한다. 앱 중복 실행 시 충돌 가능성을 줄이기 위해 단일 실행 보호와 고유 임시파일을 검토한다. 이번 사건에서 중복 실행이 있었다는 증거는 없다.

## 검증 결과와 구현 후 확인할 항목

기존 테스트:

```text
.venv312/Scripts/python.exe -m pytest -q tests/test_controller.py tests/test_result_writer.py tests/test_settings.py tests/test_excel_com.py -k "not native"
26 passed, 6 deselected
```

실제 Excel을 자동으로 실행하는 native 테스트는 제외했다. 실제 AWS 사용자 할당이나 사용자 원본 파일 수정 없이 위 모의 실험을 수행했다. 앱 코드는 변경하지 않았다.

개선 시에는 일시 잠금 후 해제, COM busy 후 회복, 영구 잠금, COM 무응답, cleanup 실패의 원인 보존, 원본 외부 변경 중 재시도, timeout 후 늦은 응답, 중지/종료 중 저장, 실제 할당 완료 후 재실행 방지까지 확인해야 한다.

## 다음 발생 시 원인 확정에 필요한 정보

발생 시각, 실제 실행 버전, 마지막 셀, 마지막 메시지 원문, 중지 버튼 반응 여부와 함께 해당 실행 환경의 `logs/cellrelay.log` 및 `config/progress.json`을 확인한다. EXE는 `%LOCALAPPDATA%/CellRelay/`, 소스 실행은 저장소 아래에 저장된다. 이번 조사에서 확인한 과거 progress 파일을 현재 복구에 사용해서는 안 된다.

## 공식 참고 자료

- [Office 스레딩 및 busy 호출](https://learn.microsoft.com/en-us/visualstudio/vsto/threading-support-in-office?view=visualstudio)
- [Excel AutomationSecurity: DisplayAlerts와 보안 경고](https://learn.microsoft.com/en-us/office/vba/api/excel.application.automationsecurity)
- [COM 오류 코드: RPC_E_CALL_REJECTED / RPC_E_SERVERCALL_RETRYLATER](https://learn.microsoft.com/en-us/windows/win32/com/com-error-codes-3)
- [Windows 파일 공유 모드 및 이름 변경 접근](https://learn.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-createfilea)

별도 프로세스 감시와 부모의 최종 파일 교체는 위 문서가 특정 Python 구현을 지시한다는 뜻이 아니라, 현재 코드의 실행·복구 구조를 바탕으로 한 설계 제안이다.
