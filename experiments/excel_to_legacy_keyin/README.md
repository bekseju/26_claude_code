# 엑셀 → 레거시(VB/Delphi) 프로그램 자동 키인 학습 키트

엑셀 데이터를 읽어서 오래된 Windows 프로그램에 자동으로 입력하는 방법을 **Python(pywinauto)** 과 **AutoHotkey v2** 두 가지로 연습합니다.
실제 업무 프로그램 대신 쓸 수 있는 **연습용 레거시 프로그램(mock_app)** 도 함께 들어 있습니다.

```
excel_to_legacy_keyin/
├─ mock_app/   연습용 "옛날 프로그램" (Win32 네이티브 컨트롤, 로그인, 메시지박스, 메뉴, 처리 지연)
├─ data/       샘플 엑셀/CSV (+ 실무에서 자주 보는 데이터 함정을 일부러 넣음)
├─ python/     pywinauto 단계별 예제 01~05
└─ ahk/        AutoHotkey v2 단계별 예제 01~04
```

## 1. 준비

```powershell
cd experiments\excel_to_legacy_keyin
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt

python data\make_sample_excel.py        # (선택) 샘플 데이터 다시 만들기
```

- AutoHotkey는 **v2**를 설치하세요 (https://www.autohotkey.com). `.ahk` 파일을 더블클릭하면 실행됩니다.
- AHK의 `01_read_data.ahk` Excel COM 예제는 Excel이 설치되어 있어야 합니다(CSV 예제는 Excel 없이 동작).

## 2. 연습용 레거시 프로그램 (mock_app)

`mock_app\run_mock_app.bat` 을 실행하세요. 로그인 정보: **admin / 1234**  (`-NoLogin`, `-NoDelay` 옵션 지원)

| 요소 | 연습 포인트 |
|---|---|
| 로그인 창 | 창 전환, 창 대기 |
| Edit 8개 (금액은 읽기 전용) | 값 입력, 입력 후 다시 읽어서 검증, 프로그램이 계산한 값 읽기 |
| 지역 ComboBox (드롭다운 목록) | 항목 선택, 목록에 없는 값 처리 |
| 개인/법인 Radio, VIP CheckBox | 상태 확인 후 클릭 |
| 저장 → 0.3~1.5초 "처리중" 멈춤 | 고정 sleep 대신 **조건 대기** |
| 성공/오류 메시지박스(#32770) | 내용 읽어 성공/실패 판정, 닫기 |
| 고객코드 중복·형식 검증 | 프로그램이 거부하는 데이터 처리 |
| Win32 메뉴, Ctrl+N, Alt+S 단축키 | 키보드 전용 입력 |

저장된 데이터는 `mock_app\saved_records.csv` 에 쌓이니 자동화 결과를 직접 확인할 수 있습니다.
처음부터 다시 하려면 메뉴 **파일 → 저장 데이터 모두 삭제** 를 누르세요.

> 이 프로그램은 PowerShell WinForms로 만들었지만, 일부러 `FlatStyle=System` 과 Win32 `MainMenu` 를 써서
> VB6/Delphi처럼 **표준 Win32 컨트롤**로 보이게 했습니다. 그래서 같은 코드를 실제 레거시 프로그램에 그대로 적용할 수 있습니다.

## 3. 학습 순서

### Python (pywinauto) — `python\` 폴더에서 실행

`config.py` = 설정, `common.py` = 엑셀 읽기·정규화 / 실행·연결 / 컨트롤 찾기 / 입력(`fill_order`) / 저장 판정(`save_and_get_result`) 공통 함수

| 파일 | 내용 |
|---|---|
| `01_read_excel.py` | openpyxl/pandas로 읽기, **데이터 정규화**(전화번호 앞자리 0, 3.0→3, 날짜 3종, 공백/대소문자), 사전 검증 |
| `02_inspect_app.py` | 창 목록, `print_control_identifiers()`, 컨트롤 위치 순서, 필드 매핑 → `controls_dump.txt` |
| `03_input_one.py` | 컨트롤 종류별 입력 방법 (Edit/Combo/Radio/Check/Button/MsgBox) + 입력값 검증 |
| `04_bulk_input.py` | **실무형 일괄 입력**: 결과를 `orders_result.xlsx` 에 행마다 기록, 이어서 하기, ESC 중지, 로그 |
| `05_keyboard_only.py` | 컨트롤 인식이 안 될 때: Tab 순서 + 클립보드 붙여넣기 + 단축키 |

```powershell
python 04_bulk_input.py --dry-run     # 입력 없이 데이터 확인
python 04_bulk_input.py               # 실행 (재실행하면 실패한 행만 다시 시도)
python 04_bulk_input.py --restart     # 처음부터
```

### AutoHotkey v2 — `ahk\` 폴더

| 파일 | 내용 |
|---|---|
| `common.ahk` | 실행·연결·로그인, 컨트롤 위치순 정렬, 메시지박스 대기, CSV 파서 |
| `01_read_data.ahk` | CSV 읽기 / Excel COM 읽기 (`.Text` 와 `.Value` 비교) |
| `02_input_one.ahk` | `ControlSetText`, `ControlChooseString`, `ControlSetChecked`, `ControlClick` |
| `03_bulk_input.ahk` | 일괄 입력, 결과 CSV(`data\orders_result_ahk.csv`), 이어서 하기, **Esc 중지 / F8 일시정지** |
| `04_keyboard_only.ahk` | `WinActivate` + `SendText` + 클립보드 + Tab 순서 |

### 기대 결과 (샘플 12건)
성공 8건 / 실패·건너뜀 4건 — 고객코드 없음, 코드 중복(프로그램이 거부), 수량 "다섯", 지역 "제주".

## 4. 실제 업무 프로그램에 적용하기

1. **구조부터 파악하기** — `02_inspect_app.py` 와 함께 *AutoHotkey Window Spy*, *Inspect.exe*(Windows SDK), *Spy++* 로 컨트롤 클래스를 확인합니다.
   - VB6: `ThunderRT6TextBox`, `ThunderRT6ComboBox`, `ThunderRT6CommandButton` …
   - Delphi: `TEdit`, `TComboBox`, `TButton`, `TStringGrid` …
   - 둘 다 **Label에는 핸들이 없어서** 라벨 이름으로 컨트롤을 찾을 수 없습니다 → 이 키트처럼 **위치 순서**로 찾거나, AHK에서는 ClassNN(`TEdit3`)으로 찾습니다.
2. **설정 바꾸기** — Python은 `python\config.py`(창 제목, 실행 명령, `EDIT_ORDER`, 타임아웃 등), AHK는 `common.ahk` 맨 위 설정 블록만 바꾸면 됩니다.
3. **입력 방식 고르기** (위에서부터 시도)
   1. 컨트롤 직접 제어 (`set_edit_text` / `ControlSetText`) — 빠르고, 창이 뒤에 있어도 동작
   2. 실제 키 입력 (`type_keys` / `ControlSend`) — 프로그램이 KeyPress/LostFocus 이벤트에서만 값을 처리할 때
   3. 키보드 전용 (Tab + 클립보드) — 그리드, 커스텀 컨트롤, Citrix/원격 화면
4. **pywinauto backend** — 옛날 프로그램은 대부분 `win32`. 컨트롤이 안 보이면 `uia` 로 바꿔 보세요.

## 5. 자주 겪는 문제

| 증상 | 원인 / 해결 |
|---|---|
| 한글이 깨지거나 `ㅊ0001` 처럼 입력됨 | 한/영 IME 상태 문제 → `set_edit_text`/`ControlSetText` 또는 **클립보드 붙여넣기** 사용 |
| 화면엔 값이 보이는데 저장하면 빈 값 | 프로그램이 키 이벤트에서만 값을 반영 → `type_keys` 로 입력하거나 입력 후 `{TAB}` 로 포커스 이동(LostFocus 유발) |
| 체크박스/라디오가 바뀌었는데 프로그램이 모름 | `check()`(BM_SETCHECK)는 Click 이벤트를 안 일으킴 → **`click()`** 사용 |
| 가끔 실패, 천천히 하면 성공 | 고정 `sleep` 대신 **창/값이 나타날 때까지 대기** (`wait`, `wait_until`, `WinWait`) |
| 아무 일도 안 일어남 | 관리자 권한으로 실행된 프로그램은 **자동화 스크립트도 관리자 권한**으로 실행해야 함 (UIPI) |
| 화면 잠금/원격 최소화 시 실패 | 키보드·마우스 방식은 화면이 켜져 있어야 함 → 컨트롤 직접 제어 방식 사용 |
| 결과 엑셀 저장 실패 | `orders_result.xlsx` 가 Excel에 열려 있음 → 닫고 재실행 (이어서 하기 지원) |
| 엑셀 숫자가 `3.0`, 전화번호 0 사라짐 | `python\common.py` 의 `to_text`, `to_phone`, `to_date_text` 정규화 참고 |

> ⚠️ 실제 업무 시스템에는 먼저 **테스트 DB/소량 데이터**로 검증한 뒤 적용하세요. 결과 파일과 로그를 남기는 구조(04/03 예제)를 유지하는 것이 좋습니다.
