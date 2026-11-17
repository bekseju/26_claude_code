"""
샘플 엑셀(orders.xlsx) + CSV(orders.csv, AutoHotkey 용) 생성 스크립트

실무에서 자주 만나는 "엑셀 데이터 함정"을 일부러 넣어 두었습니다.
  - 전화번호가 숫자로 저장되어 앞자리 0 이 사라짐      (1012345678)
  - 수량/단가가 float 으로 읽힘                         (3.0)
  - 날짜가 datetime / 문자열("2026/11/05") / 엑셀 일련번호로 섞여 있음
  - 앞뒤 공백, 소문자 고객코드, VIP 표기가 Y / O / TRUE / 빈칸 으로 제각각
  - 프로그램에서 거부될 데이터 (빈 고객코드, 중복, 수량 문자, 없는 지역)

실행:  python make_sample_excel.py
"""
import csv
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

HERE = Path(__file__).resolve().parent

HEADERS = ["No", "고객코드", "고객명", "전화번호", "지역", "고객구분", "VIP",
           "주문일자", "수량", "단가", "비고", "처리결과", "처리시각"]

ROWS = [
    # No, 고객코드, 고객명, 전화번호, 지역, 고객구분, VIP, 주문일자, 수량, 단가, 비고
    [1,  "C0001", "김철수",      "010-1234-5678", "서울", "개인", "Y",  datetime(2026, 11, 1), 3, 15000, "첫 주문"],
    [2,  "C0002", "(주)한빛상사", "02-555-1234",   "부산", "법인", "",   datetime(2026, 11, 1), 10, 8200, ""],
    [3,  "C0003", "이영희",      1012345678,      "대구", "개인", "O",  "2026/11/05",          2.0, 33000.0, "전화번호가 숫자로 저장된 행"],
    [4,  " c0004 ", "박민수 ",    "010-2222-3333", "인천", "개인", "TRUE", 46335,              1, 120000, "코드 소문자+공백, 날짜=엑셀 일련번호"],
    [5,  "C0005", "최지우",      "010-4444-5555", "광주", "개인", "N",  datetime(2026, 11, 3), 5, 4500, "여러 줄\n비고 테스트"],
    [6,  "",      "코드없음",    "010-0000-0000", "서울", "개인", "",   datetime(2026, 11, 3), 1, 1000, "[오류 예상] 고객코드 없음"],
    [7,  "C0003", "중복고객",    "010-9999-8888", "서울", "개인", "",   datetime(2026, 11, 4), 1, 1000, "[오류 예상] 3번과 코드 중복"],
    [8,  "C0008", "정우성",      "010-7777-6666", "대전", "법인", "Y",  datetime(2026, 11, 4), "다섯", 2000, "[오류 예상] 수량이 문자"],
    [9,  "C0009", "강하늘",      "010-1111-2222", "제주", "개인", "",   datetime(2026, 11, 5), 2, 7000, "[오류 예상] 콤보에 없는 지역"],
    [10, "C0010", "(주)미래테크", "031-777-8888",  "울산", "법인", "Y",  datetime(2026, 11, 5), 100, 990, ""],
    [11, "C0011", "윤서연",      "010-3333-4444", "서울", "개인", "",   datetime(2026, 11, 6), 7, 12000, ""],
    [12, "C0012", "한지민",      "010-5555-6666", "부산", "개인", "Y",  datetime(2026, 11, 6), 1, 250000, "고액 주문"],
]


def make_xlsx(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "주문데이터"
    ws.append(HEADERS)
    for r in ROWS:
        ws.append(r + ["", ""])

    head_fill = PatternFill("solid", fgColor="DDEBF7")
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = head_fill
        c.alignment = Alignment(horizontal="center")
    for row in ws.iter_rows(min_row=2, min_col=8, max_col=8):
        row[0].number_format = "yyyy-mm-dd"
    widths = [5, 10, 14, 15, 7, 9, 6, 12, 7, 10, 34, 30, 19]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = w
    ws.freeze_panes = "A2"

    dv = DataValidation(type="list", formula1='"서울,부산,대구,인천,광주,대전,울산"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("E2:E200")

    # 설명 시트
    info = wb.create_sheet("설명")
    for line in [
        ["주문데이터 시트의 각 행을 연습용 프로그램(mock_app)에 입력합니다."],
        ["처리결과/처리시각 열은 자동화 스크립트가 채웁니다 (결과는 orders_result.xlsx 로 저장)."],
        ["[오류 예상] 표시 행은 프로그램이 거부하도록 일부러 넣은 데이터입니다."],
    ]:
        info.append(line)
    info.column_dimensions["A"].width = 90
    wb.save(path)


def make_csv(path: Path) -> None:
    """AutoHotkey 용 CSV. (엑셀 함정 없이 '정리된 문자열'로 저장, UTF-8 BOM)"""
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(HEADERS[:11])
        for r in ROWS:
            r = list(r)
            r[1] = str(r[1]).strip().upper()
            r[2] = str(r[2]).strip()
            if isinstance(r[3], int):
                r[3] = "0" + str(r[3])
            if isinstance(r[7], datetime):
                r[7] = r[7].strftime("%Y-%m-%d")
            elif isinstance(r[7], int):
                r[7] = "2026-11-09"          # 엑셀 일련번호 46335 = 2026-11-09
            else:
                r[7] = str(r[7]).replace("/", "-")
            for i in (8, 9):
                if isinstance(r[i], float):
                    r[i] = int(r[i])
            r[6] = "Y" if str(r[6]).upper() in ("Y", "O", "TRUE") else "N"
            r[10] = str(r[10]).replace("\n", " / ")
            w.writerow(r)


if __name__ == "__main__":
    make_xlsx(HERE / "orders.xlsx")
    make_csv(HERE / "orders.csv")
    print("생성 완료:", HERE / "orders.xlsx", ",", HERE / "orders.csv")
