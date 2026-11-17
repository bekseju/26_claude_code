"""
[3단계] 한 건 입력 - 컨트롤 종류별 입력 방법 익히기

  Edit     : set_edit_text()  (WM_SETTEXT, 빠르고 한글 안전)  /  type_keys() (실제 키 입력)
  ComboBox : select("부산")   (목록에 없으면 ValueError)
  Radio    : click()          (check() 는 BM_SETCHECK 만 보내서 Click 이벤트가 안 뜰 수 있음)
  CheckBox : 상태가 다를 때만 click()
  Button   : click() (메시지 전송, 마우스 안 움직임) / click_input() (실제 마우스 클릭)
  MsgBox   : #32770 창을 기다렸다가 내용 읽고 확인 클릭
  검증     : window_text() 로 다시 읽기 (프로그램이 계산한 금액도 읽기)

여기서 익힌 단계는 common.fill_order() / save_and_get_result() 로 묶여 04 에서 재사용됩니다.

실행: python 03_input_one.py
"""
from common import (Order, get_fields, save_and_get_result, select_combo, set_checked,
                    set_text, start_or_connect)

order = Order(code="C9001", name="홍길동", phone="010-1234-0000", region="부산", cust_type="법인",
              vip=True, order_date="2026-11-07", qty="4", price="12500", memo="pywinauto 한 건 입력 테스트")

app, main = start_or_connect()
f = get_fields(main)
f["초기화"].click()

# --- Edit: set_edit_text + 다시 읽어서 검증 (common.set_text)
for name, value in {"고객코드": order.code, "고객명": order.name, "전화번호": order.phone,
                    "주문일자": order.order_date, "수량": order.qty, "단가": order.price,
                    "비고": order.memo}.items():
    set_text(f[name], value)

# (참고) type_keys : KeyDown/KeyPress 이벤트로만 값을 받는 필드용 - 포커스를 옮기고 실제로 타이핑
#   f["고객명"].set_edit_text("")
#   f["고객명"].type_keys(order.name, with_spaces=True)

select_combo(f["지역"], order.region)    # ComboBox
f[order.cust_type].click()               # Radio
set_checked(f["VIP"], order.vip)         # CheckBox (토글)

print("콤보 선택값:", f["지역"].selected_text())
print("법인:", f["법인"].get_check_state() == 1, "/ VIP:", f["VIP"].get_check_state() == 1)
print("프로그램이 계산한 금액:", f["금액"].window_text())

ok, message = save_and_get_result(app, f)    # 저장 → 메시지박스 판정 → 닫기
print(f"{'성공' if ok else '실패'}: {message}")
if not ok:
    print("  (이미 등록된 코드라면 정상입니다. 메뉴 '파일 → 저장 데이터 모두 삭제' 후 재실행)")
