"""
[1단계] 엑셀 읽기 + 데이터 정규화

엑셀 값을 '그대로' 넣으면 사고가 납니다.
  - 전화번호 1012345678 (0 이 사라짐)    - 수량 3.0 (float)
  - 날짜 datetime / 문자열 / 일련번호 혼재 - 앞뒤 공백, 소문자
'원본 값' 과 '정규화된 값' 을 나란히 확인합니다.

실행: python 01_read_excel.py
"""
from openpyxl import load_workbook

from common import read_orders
from config import EXCEL_PATH, SHEET_NAME

# 1) openpyxl 원본 셀 값 (repr 로 타입까지 확인)
print("=== 1) 원본 셀 값 ===")
wb = load_workbook(EXCEL_PATH, data_only=True)
for row in wb[SHEET_NAME].iter_rows(min_row=1, max_row=5, values_only=True):
    print("  ", [repr(v) for v in row[:10]])
wb.close()

# 2) 정규화된 값
print("\n=== 2) 정규화 후 (common.read_orders) ===")
orders = read_orders()
for o in orders:
    print(f"  행{o.row:>2} | {o.code:6s}| {o.name:8s}| {o.phone:14s}| {o.region:3s}| {o.cust_type}"
          f" | VIP={'Y' if o.vip else 'N'} | {o.order_date:10s} | 수량={o.qty:>4s} | 단가={o.price:>7s}")

# 3) 사전 검증 - 프로그램에 넣기 전에 걸러낼 수 있는 것은 미리 거른다
print("\n=== 3) 사전 검증 ===")
seen: set[str] = set()
for o in orders:
    problems = []
    if not o.code:
        problems.append("고객코드 없음")
    elif o.code in seen:
        problems.append("파일 내 중복 코드")
    if not o.qty.isdigit():
        problems.append(f"수량이 숫자가 아님({o.qty})")
    seen.add(o.code)
    if problems:
        print(f"  행{o.row}: {', '.join(problems)}")

# 4) (참고) pandas - 대량 데이터 확인·필터링에 편리
try:
    import pandas as pd
except ImportError:
    print("\n(pandas 미설치 - 건너뜀)")
else:
    df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME, dtype=str)   # 모든 값을 문자열로
    print("\n=== 4) pandas (dtype=str) ===")
    print(df[["고객코드", "고객명", "전화번호", "주문일자", "수량"]].head(5).to_string())
    print("  ※ dtype=str 이어도 숫자로 저장된 전화번호의 0 은 이미 사라져 있으므로 정규화가 필요합니다.")
