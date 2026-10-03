# CellRelay v0.3.3 · 재개 안내와 릴리스 검증

날짜: 2026-10-03 KST · 범위: 기존 v0.3.2의 작은 문구 수정

## 변경 범위

기준은 v0.3.2 `17e83c0abe9a4f5c133226968b18e9c23efdb173`이며,
재개 문구 패치는 [PR #1](https://github.com/prozac0401/CellRelay/pull/1)의
`f8a7d772156a6d468b51464d557c245c5629d80e`다.
일반 재개 안내를 `Text 입력을 재개했습니다.`에서 `작업을 재개했습니다.`로 바꿨다.
Text Clear 대기 안내, 상태 전환, worker 명령, 저장, 완료 기록과 다음 행 이동은 유지한다.
런타임 버전은 `app/__init__.py`의 `0.3.3`이며 최종 소스 SHA와 산출물 해시는
릴리스 본문 및 `SHA256SUMS.txt`로 연결한다.

## 이번에 실행한 확인

| 확인 | 결과와 범위 |
|---|---|
| 새 재개 사례 | Windows Python 3.12.14 / PySide6 6.11.2에서 2 PASS, 13 deselected. SEARCHING_USER와 CONFIRMING_ASSIGNMENT, 메시지 1회, 반복 재개 무동작, 단계·셀·처리 개수 보존. 실제 Qt/QThread와 합성 OOXML fixture이며 실제 AWS 요청은 보내지 않음 |
| Controller 회귀 | 15 PASS. 실제 브라우저 명령을 분리한 기존 Controller 시험 범위. 전체 suite 결과가 아님 |
| 실제 Qt 화면 | 합성 SEARCHING_USER 상태에서 Computer use로 일시정지·재개. 새 문구가 화면에 표시되고 B5·처리 0·단계 유지. 사용자 Excel과 AWS에 연결하지 않은 소스 화면 확인 |
| 최종 EXE 빌드 | Windows 11 22631 / Python 3.12.14 / PyInstaller 6.22.2. `build_exe.ps1 -SkipInstall -DistDirectory dist-v0.3.3` 성공. 단일 Windows x64 EXE, 87,240,197 bytes, Authenticode NotSigned |
| 최종 EXE 코드 대조 | 내장 버전 0.3.3, 앱 모듈 20개와 main의 code object 285개가 최종 소스와 일치. 아카이브 이름 365개에서 업무 Office/CSV/TSV 파일, 프로젝트 설정·로그·백업 항목 0개. dependency 파일 내부의 임의 개인정보까지 판별하는 검사는 아님 |
| 최종 EXE 스모크 | `--smoke-test`, `--playwright-smoke-test`, `--browser-worker-smoke-test`, `--excel-smoke-test` 모두 PASS / 종료 코드 0. 각 실행의 별도 LOCALAPPDATA와 합성 자료를 사용. 내장 버전과 정상 종료 로그 확인, 스레드 종료 지연·ERROR·Traceback 없음. Edge 합성 textarea 입력과 별도 Excel COM 생성·읽기·저장·백업·재열기 확인 |

최종 `CellRelay.exe` SHA-256:
`6e1e75a6cda367147a168c2462cc4357405a7f125584e95a4d667b1e463fdf1c`.

기존 개발 환경의 Python base 경로가 없어 실행되지 않는 사실을 확인했다.
원래 환경은 바꾸지 않고 별도 작업 트리에 동 버전 Python 3.12.14와
기존 개발 의존성을 연결하여 확인했다. 실제 사용자는 Python 개발 환경이 필요 없는
단일 Windows x64 EXE를 사용한다.

## 보존하는 제한과 미실행

이번에는 실제 AWS 계정·한국어/영어 운영 화면의 사용자 할당, 전체 GUI·전체 suite,
다른 PC·새 프로필·재부팅을 반복하지 않았다. 기존 릴리스의 수동 시험을 이번 후보의
시험 수치로 합산하지 않는다. 저장소에 GitHub Actions는 구성되지 않았으며 CI PASS가 아니다.

COM 호출 자체가 반환되지 않을 때의 강제 복구는 지원하지 않는다. 과거 브라우저 종료
지연 기록은 보존한다. 프로그램의 완료 처리는 Done 클릭과 팝업 종료를 관찰한 결과이며
AWS 서버 등록 성공을 입증하지 않는다. 최종 등록 명단은 관리자가 별도로 대조해야 한다.
이번 문구 수정이 이러한 제한을 해결했다고 판정하지 않는다.

[v0.3.2 실제 시험과 실패 이력](intermittent-stall-diagnosis-2026-09-19.md)은 당시 근거로 보존한다.
실사용 파일·계정 정보·개인 설정·원시 로그는 릴리스 자산에 포함하지 않는다.
