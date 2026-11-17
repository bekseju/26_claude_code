"""
[5단계] 키보드만으로 입력 - 컨트롤 인식이 안 되는 프로그램용 '최후의 수단'

다음과 같은 경우에 씁니다.
  - 그리드/커스텀 컨트롤이라 print_control_identifiers 에 안 보임 (Delphi 서드파티, VB6 FlexGrid 등)
  - set_edit_text 로 넣으면 프로그램이 값을 인식 못함 (KeyPress 이벤트에서만 처리)
  - 원격 데스크톱/Citrix 화면 등 컨트롤 정보 자체가 없음

핵심 기법
  - 창을 맨 앞으로(set_focus) -> 첫 필드로 포커스 이동 (여기선 Ctrl+N)
  - 값 입력 후 {TAB} 으로 다음 필드 이동   (Tab 순서를 미리 손으로 확인해 둘 것!)
  - 한글/특수문자는 '클립보드에 넣고 Ctrl+V' 가 가장 안전 (한/영 IME 상태 영향 없음)
  - 콤보는 {HOME} 후 {DOWN n} 으로 n 번째 항목 선택
  - 라디오는 방향키, 체크박스는 {SPACE}, 버튼은 Alt+단축키(%s = Alt+S)
  - 실제 키 입력이므로 실행 중 마우스/키보드를 건드리면 안 됨

실행: python 05_keyboard_only.py
"""
import time

from pywinauto.keyboard import send_keys

from common import Order, close_message_box, set_clipboard, start_or_connect, wait_message_box
from config import REGIONS, SAVE_TIMEOUT

STEP = 0.15          # 키 입력 사이 대기(초). 느린 프로그램이면 늘리세요.

order = Order(code="C9002", name="김키보드", phone="010-5555-0000", region="대전", cust_type="법인",
              vip=True, order_date="2026-11-07", qty="2", price="30000", memo="키보드 전용 입력 (Tab + 클립보드)")


def keys(k: str) -> None:
    send_keys(k)
    time.sleep(STEP)


def type_text(text: str) -> None:
    """직접 타이핑 (영문/숫자 위주. 한글 IME 가 켜져 있으면 깨질 수 있음). 특수문자 + ^ % ~ ( ) { } [ ] 는 이스케이프"""
    send_keys("".join("{%s}" % c if c in "+^%~(){}[]" else c for c in text), with_spaces=True, pause=0.02)
    time.sleep(STEP)


def paste(text: str) -> None:
    """클립보드 붙여넣기 (Ctrl+A 로 기존 값 선택 후 덮어씀)"""
    set_clipboard(text)
    keys("^a^v")


app, main = start_or_connect()
main.set_focus()                         # 반드시 창을 맨 앞으로!
keys("^n")                               # Ctrl+N : 새로 입력 + 고객코드에 포커스
time.sleep(0.5)

type_text(order.code); keys("{TAB}")     # 영문/숫자는 직접 타이핑
paste(order.name); keys("{TAB}")         # 한글은 클립보드
type_text(order.phone); keys("{TAB}")

keys("{HOME}")                           # 콤보: 첫 항목 → ↓ n 번
if (n := REGIONS.index(order.region)) > 0:
    keys("{DOWN %d}" % n)
keys("{TAB}")

if order.cust_type == "법인":            # 라디오: 포커스는 '개인' → → 키로 '법인'
    keys("{RIGHT}")
keys("{TAB}")
if order.vip:                            # 체크박스: 스페이스로 토글
    keys("{SPACE}")
keys("{TAB}")

for value in (order.order_date, order.qty, order.price):
    type_text(value); keys("{TAB}")
paste(order.memo)

print("입력 완료 - 화면을 확인하세요. 2초 후 저장합니다.")
time.sleep(2)
send_keys("%s")                          # Alt+S = 저장(&S)

box = wait_message_box(app, timeout=SAVE_TIMEOUT)
if box:
    print(f"메시지박스: [{box.title}] {box.text}")
    close_message_box(box)
else:
    print("메시지박스 없음")
