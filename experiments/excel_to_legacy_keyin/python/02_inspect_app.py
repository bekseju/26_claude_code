"""
[2단계] 대상 프로그램 분석 - "어떤 컨트롤이 있고, 어떻게 찾을 수 있나?"

자동화의 80%는 이 단계입니다. 실제 업무 프로그램에서도 먼저 이 스크립트로 구조를 확인하세요.
  1) 떠 있는 최상위 창 목록 (제목 / 클래스 / PID)
  2) 대상 창의 컨트롤 트리 (print_control_identifiers) -> controls_dump.txt 저장
  3) Edit 컨트롤을 위치 순서로 정렬해서 출력 (라벨에 핸들이 없는 VB6/Delphi 대응)

backend 선택:
  - "win32" : VB6, Delphi, MFC, 일반 Win32 프로그램  (대부분의 옛날 프로그램은 이것)
  - "uia"   : WPF, 최신 .NET, 일부 Delphi(VCL 스타일 적용) 등 win32 로 안 보일 때 시도
  * 도구: Inspect.exe(Windows SDK), Spy++(Visual Studio), AutoHotkey Window Spy 도 함께 쓰면 좋습니다.

실행: python 02_inspect_app.py
"""
from pathlib import Path

from pywinauto import Desktop

from common import get_fields, sorted_children, start_or_connect
from config import BACKEND, MAIN_TITLE_RE

OUT = Path(__file__).with_name("controls_dump.txt")

# 1) 최상위 창 목록 ----------------------------------------------------------
print("=== 현재 보이는 최상위 창 (일부) ===")
for w in Desktop(backend=BACKEND).windows(visible_only=True):
    title = w.window_text()
    if title:
        print(f"  {title[:40]:40s} | class={w.class_name()[:40]:40s} | pid={w.process_id()}")

# 2) 대상 창 컨트롤 트리 -------------------------------------------------------
app, main = start_or_connect(do_login=True)
print(f"\n=== 대상 창: {main.window_text()} ===")
main.print_control_identifiers(filename=str(OUT))
print(f"컨트롤 트리를 {OUT.name} 에 저장했습니다. (child_window(...) 예시 코드가 함께 들어 있음)")

# 3) 컨트롤 하나씩 살펴보기 -----------------------------------------------------
print("\n=== 자식 컨트롤 목록 ===")
for c in main.children():
    r = c.rectangle()
    print(f"  {c.friendly_class_name():12s} | text={c.window_text()[:15]!r:20s} "
          f"| class={c.class_name()[:35]:35s} | pos=({r.left},{r.top})")

print("\n=== Edit 컨트롤 (위치 순서) ===")
for i, e in enumerate(sorted_children(main, "Edit")):
    r = e.rectangle()
    print(f"  [{i}] pos=({r.left},{r.top})  현재값={e.window_text()!r}")

print("\n=== 필드 매핑 결과 (common.get_fields) ===")
fields = get_fields(main)
for name, ctrl in fields.items():
    print(f"  {name:6s} -> {ctrl.friendly_class_name():12s} handle={ctrl.handle}")

print("\n콤보 항목:", fields["지역"].item_texts())
print("\n참고: 창 제목 정규식 =", MAIN_TITLE_RE)
