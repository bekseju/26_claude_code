"""
공통 모듈
  1) 엑셀 읽기/정규화   2) 프로그램 실행·연결   3) 컨트롤 찾기
  4) 컨트롤 입력        5) 메시지박스           6) 기타 유틸
설정값은 config.py 에 있습니다.
"""
from __future__ import annotations

import ctypes
import re
import time
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from pywinauto import Application, timings
from pywinauto.findwindows import ElementNotFoundError

import config as cfg

BUTTON_CLASS_RE = r"(?i).*button.*"     # Win32 Button, ThunderRT6CommandButton, TButton, WindowsForms10.BUTTON


# =============================================================== 1) 엑셀 읽기
@dataclass
class Order:
    code: str
    name: str
    phone: str
    region: str
    cust_type: str = "개인"   # 개인 / 법인
    vip: bool = False
    order_date: str = ""      # YYYY-MM-DD
    qty: str = ""
    price: str = ""
    memo: str = ""
    row: int = 0              # 엑셀 행 번호 (결과 기록용)
    status: str = ""          # 이전 실행의 처리결과 (이어서 하기용)


def to_text(v: Any) -> str:
    """셀 값 -> 입력용 문자열  (3.0 -> '3', None -> '', 앞뒤 공백 제거)"""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def to_date_text(v: Any) -> str:
    """datetime / '2026/11/05' / 엑셀 일련번호(46335) -> 'YYYY-MM-DD'"""
    if v in (None, ""):
        return ""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, date):
        return v.isoformat()
    if isinstance(v, (int, float)):                        # 엑셀 날짜 일련번호
        return (date(1899, 12, 30) + timedelta(days=int(v))).isoformat()
    s = str(v).strip().replace("/", "-").replace(".", "-")
    try:
        return datetime.strptime(s, "%Y-%m-%d").strftime("%Y-%m-%d")
    except ValueError:
        return s                                           # 그대로 넘기고 프로그램이 판단하게 함


def to_phone(v: Any) -> str:
    """숫자로 저장돼 0 이 사라진 전화번호 복구: 1012345678 -> 010-1234-5678"""
    s = to_text(v)
    if re.fullmatch(r"\d{9,10}", s) and not s.startswith("0"):
        s = "0" + s
    if re.fullmatch(r"01\d{8,9}", s):
        return f"{s[:3]}-{s[3:-4]}-{s[-4:]}"
    return s


def to_bool(v: Any) -> bool:
    return to_text(v).upper() in ("Y", "O", "TRUE", "1", "예", "V")


def read_orders(path: Path = cfg.EXCEL_PATH, sheet: str = cfg.SHEET_NAME) -> list[Order]:
    """헤더 이름으로 열을 찾으므로 열 순서가 바뀌어도 동작합니다."""
    wb = load_workbook(path, data_only=True)               # data_only : 수식 대신 계산된 값
    try:
        ws = wb[sheet]
        col = {to_text(c.value): i for i, c in enumerate(ws[1])}
        orders = []
        for r_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if all(v in (None, "") for v in row):
                continue                                   # 빈 줄 건너뜀

            def g(h: str) -> Any:
                return row[col[h]] if h in col else None

            orders.append(Order(
                code=to_text(g("고객코드")).upper(),
                name=to_text(g("고객명")),
                phone=to_phone(g("전화번호")),
                region=to_text(g("지역")),
                cust_type=to_text(g("고객구분")) or "개인",
                vip=to_bool(g("VIP")),
                order_date=to_date_text(g("주문일자")),
                qty=to_text(g("수량")),
                price=to_text(g("단가")),
                memo=to_text(g("비고")),
                row=r_idx,
                status=to_text(g("처리결과")),
            ))
        return orders
    finally:
        wb.close()


# ====================================================== 2) 프로그램 실행/연결
def start_or_connect(do_login: bool = True):
    """떠 있으면 연결(connect), 없으면 실행(start) 후 (app, 메인창) 반환"""
    app = Application(backend=cfg.BACKEND)
    try:
        app.connect(title_re=f"{cfg.MAIN_TITLE_RE}|{cfg.LOGIN_TITLE_RE}", timeout=2)
        print("[연결] 실행 중인 프로그램에 연결했습니다.")
    except (ElementNotFoundError, timings.TimeoutError):
        print("[실행]", cfg.APP_CMD)
        app.start(cfg.APP_CMD, wait_for_idle=False)

    login = app.window(title_re=cfg.LOGIN_TITLE_RE)
    main = app.window(title_re=cfg.MAIN_TITLE_RE)
    timings.wait_until(cfg.WINDOW_TIMEOUT, 0.3, lambda: login.exists(0) or main.exists(0))
    if do_login and login.exists(0):
        login_to(login)

    main.wait("visible", timeout=cfg.WINDOW_TIMEOUT)
    close_all_message_boxes(app)                # 모달 메시지박스가 떠 있으면 메인 창이 비활성
    main.wait("enabled ready", timeout=cfg.WINDOW_TIMEOUT)
    return app, main


def login_to(login_win) -> None:
    user_id, password = sorted_children(login_win, "Edit")
    user_id.set_edit_text(cfg.LOGIN_ID)
    password.set_edit_text(cfg.LOGIN_PW)
    find_button(login_win, "확인").click()
    print("[로그인] 완료")


# ============================================================ 3) 컨트롤 찾기
def sorted_children(win, friendly_class: str, row_tolerance: int = 8) -> list:
    """
    같은 종류의 자식 컨트롤을 화면 위치 순(위→아래, 왼→오)으로 반환.
    'Edit' 은 Win32 EDIT, VB6 ThunderRT6TextBox, Delphi TEdit/TMemo, WinForms EDIT 를 모두 포함.
    """
    ctrls = [c for c in win.children() if c.friendly_class_name() == friendly_class and c.is_visible()]
    ctrls.sort(key=lambda c: c.rectangle().top)
    rows: list[list] = []
    for c in ctrls:                              # top 차이가 tolerance 이내면 같은 줄
        if rows and abs(c.rectangle().top - rows[-1][0].rectangle().top) <= row_tolerance:
            rows[-1].append(c)
        else:
            rows.append([c])
    return [c for row in rows for c in sorted(row, key=lambda c: c.rectangle().left)]


def find_button(win, title_prefix: str):
    """버튼/라디오/체크박스를 텍스트 앞부분으로 찾기 ('저장' -> '저장(&S)')"""
    return win.child_window(title_re=f"^{re.escape(title_prefix)}", class_name_re=BUTTON_CLASS_RE).wrapper_object()


def get_fields(main) -> dict:
    """화면 컨트롤을 {이름: wrapper} 로 묶어 반환 (필드 매핑의 핵심)"""
    edits = sorted_children(main, "Edit")
    if len(edits) != len(cfg.EDIT_ORDER):
        raise RuntimeError(f"Edit 개수가 예상과 다릅니다: {len(edits)}개 (예상 {len(cfg.EDIT_ORDER)})")
    fields = dict(zip(cfg.EDIT_ORDER, edits))
    fields["지역"] = main.child_window(class_name_re=r"(?i).*combo.*").wrapper_object()
    for name in ("개인", "법인", "VIP", "저장", "초기화"):
        fields[name] = find_button(main, name)
    return fields


# ============================================================ 4) 컨트롤 입력
def set_text(edit, value: str) -> None:
    """WM_SETTEXT 로 값 넣고 다시 읽어서 검증 (빠르고 한글 안전)"""
    edit.set_edit_text(value)
    if edit.window_text() != value:
        raise RuntimeError(f"입력 실패 (기대 {value!r}, 실제 {edit.window_text()!r})")


def select_combo(combo, value: str) -> None:
    """콤보 항목 선택. 목록에 없으면 ValueError"""
    if value not in combo.item_texts():
        raise ValueError(f"콤보 목록에 없음: {value}")
    combo.select(value)


def set_checked(button, checked: bool) -> None:
    """체크박스: 상태가 다를 때만 click() (check() 는 Click 이벤트가 안 뜰 수 있음)"""
    if button.get_check_state() != int(checked):
        button.click()
    timings.wait_until(3, 0.1, lambda: button.get_check_state() == int(checked))


def fill_order(fields: dict, o: Order) -> None:
    """한 건을 화면에 입력 (초기화 → Edit → 콤보 → 라디오 → 체크박스)"""
    fields["초기화"].click()
    timings.wait_until(3, 0.1, lambda: fields["고객코드"].window_text() == "")
    select_combo(fields["지역"], o.region)      # 목록에 없는 값이면 입력 전에 실패
    for name, value in {"고객코드": o.code, "고객명": o.name, "전화번호": o.phone,
                        "주문일자": o.order_date, "수량": o.qty, "단가": o.price,
                        "비고": o.memo.replace("\n", "\r\n")}.items():
        set_text(fields[name], value)
    fields["법인" if o.cust_type == "법인" else "개인"].click()
    set_checked(fields["VIP"], o.vip)


def save_and_get_result(app, fields: dict) -> tuple[bool, str]:
    """저장 클릭 → 메시지박스로 성공/실패 판정 → 닫기.  (성공여부, 메시지) 반환"""
    fields["저장"].click()
    box = wait_message_box(app, timeout=cfg.SAVE_TIMEOUT)
    if box is None:
        raise TimeoutError("저장 응답(메시지박스) 없음")
    close_message_box(box)
    return box.title == cfg.SUCCESS_TITLE, box.text


# ============================================================ 5) 메시지박스
@dataclass
class MsgBox:
    title: str
    text: str
    wrapper: Any    # 창 wrapper (텍스트·상태 읽기)
    spec: Any       # WindowSpecification (child_window 로 버튼 찾기)


def wait_message_box(app, timeout: float = 15) -> MsgBox | None:
    """메시지박스(#32770 = VB MsgBox / Delphi ShowMessage / Win32 MessageBox) 대기"""
    spec = app.window(class_name="#32770", visible_only=True)
    try:
        spec.wait("visible", timeout=timeout)
    except timings.TimeoutError:
        return None
    w = spec.wrapper_object()
    text = " ".join(c.window_text() for c in w.children()
                    if c.friendly_class_name() == "Static" and c.window_text())
    return MsgBox(w.window_text(), text, w, spec)


def close_message_box(box: MsgBox, button_title_re: str = ".*") -> None:
    """버튼을 눌러 닫기. 기본은 첫 번째 버튼(확인/OK), '예' 는 button_title_re='^예'"""
    box.spec.child_window(title_re=button_title_re, class_name="Button", found_index=0).wrapper_object().click()
    timings.wait_until(5, 0.1, lambda: not box.wrapper.is_visible())


def close_all_message_boxes(app) -> None:
    """남아 있는 메시지박스 정리"""
    while (box := wait_message_box(app, timeout=0.5)) is not None:
        print(f"  [정리] 메시지박스 닫음: {box.title} / {box.text}")
        close_message_box(box)


# ============================================================== 6) 기타 유틸
def esc_pressed() -> bool:
    """ESC 키가 눌렸는지 (일괄 입력 비상 정지용)"""
    return bool(ctypes.windll.user32.GetAsyncKeyState(0x1B) & 0x8001)


def set_clipboard(text: str) -> None:
    import win32clipboard  # pywin32 (pywinauto 와 함께 설치됨)
    for _ in range(10):                          # 다른 프로그램이 잡고 있으면 재시도
        try:
            win32clipboard.OpenClipboard()
            break
        except Exception:
            time.sleep(0.05)
    else:
        raise RuntimeError("클립보드를 열 수 없습니다.")
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)
    finally:
        win32clipboard.CloseClipboard()


def now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")
