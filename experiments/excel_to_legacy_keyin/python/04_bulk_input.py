"""
[4단계] 엑셀 전체 일괄 입력 - 실무형 구조

  엑셀 → 정규화 → 사전검증 → 입력 → 저장 → 메시지박스 판정 → 결과를 엑셀에 기록
  - 결과는 orders_result.xlsx 에 행마다 즉시 저장 (중간에 죽어도 진행상황 보존)
  - 재실행 시 '성공' 행은 건너뜀 (이어서 하기)
  - ESC : 현재 건을 마치고 중지
  - 오류가 나도 다음 건 진행, 로그(bulk_input.log) 기록

실행:
  python 04_bulk_input.py              # 실행 (이어서 하기)
  python 04_bulk_input.py --dry-run    # 입력 없이 데이터만 확인 (프로그램 연결 안 함)
  python 04_bulk_input.py --limit 3    # 3건만
  python 04_bulk_input.py --restart    # 이전 결과 무시하고 처음부터
"""
import argparse
import logging
import shutil
import time
from pathlib import Path

from openpyxl import load_workbook

from common import (Order, close_all_message_boxes, esc_pressed, fill_order, get_fields, now,
                    read_orders, save_and_get_result, start_or_connect)
from config import EXCEL_PATH, REGIONS, RESULT_PATH, SHEET_NAME

log = logging.getLogger("bulk")


class SkipRow(Exception):
    """입력 전에 걸러진 행 (프로그램에 넣지 않음)"""


class ResultBook:
    """결과 엑셀의 처리결과/처리시각 열에 기록하고 즉시 저장"""

    def __init__(self, path: Path):
        self.path = path
        self.wb = load_workbook(path)
        self.ws = self.wb[SHEET_NAME]
        headers = [c.value for c in self.ws[1]]
        self.col_result = headers.index("처리결과") + 1
        self.col_time = headers.index("처리시각") + 1

    def write(self, row: int, result: str) -> None:
        self.ws.cell(row, self.col_result, result)
        self.ws.cell(row, self.col_time, now())
        try:
            self.wb.save(self.path)
        except PermissionError:
            log.error("%s 가 엑셀에서 열려 있어 저장 실패. 파일을 닫아주세요.", self.path.name)


def precheck(o: Order) -> None:
    if not o.code:
        raise SkipRow("고객코드 없음")
    if not o.qty.isdigit():
        raise SkipRow(f"수량이 숫자가 아님: {o.qty}")
    if o.region not in REGIONS:
        raise SkipRow(f"지역 목록에 없음: {o.region}")


def process(app, main, o: Order) -> str:
    """한 건 처리 → '성공' / '실패: …' / '건너뜀: …' / '오류: …'"""
    try:
        precheck(o)
        fields = get_fields(main)            # 매 건마다 다시 찾음 (창이 다시 그려져도 안전)
        fill_order(fields, o)
        ok, message = save_and_get_result(app, fields)
        return "성공" if ok else f"실패: {message}"
    except SkipRow as e:
        return f"건너뜀: {e}"
    except Exception as e:                   # 예상 못한 오류: 화면 정리 후 다음 건 진행
        log.exception("행%d 처리 중 오류", o.row)
        close_all_message_boxes(app)
        return f"오류: {type(e).__name__}: {e}"


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--restart", action="store_true")
    args = ap.parse_args()

    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s",
        handlers=[logging.FileHandler(Path(__file__).with_name("bulk_input.log"), encoding="utf-8"),
                  logging.StreamHandler()],
    )

    # 원본은 건드리지 않고 결과용 사본에 기록. dry-run 은 사본도 만들지 않음
    fresh = args.restart or not RESULT_PATH.exists()
    if fresh and not args.dry_run:
        shutil.copyfile(EXCEL_PATH, RESULT_PATH)
    orders = read_orders(EXCEL_PATH if fresh and args.dry_run else RESULT_PATH)
    todo = [o for o in orders if o.status != "성공"]
    if args.limit:
        todo = todo[: args.limit]
    log.info("전체 %d건 / 처리 대상 %d건", len(orders), len(todo))

    if args.dry_run:
        for o in todo:
            log.info("[DRY] 행%d %s", o.row, o)
        return

    app, main_win = start_or_connect()
    book = ResultBook(RESULT_PATH)
    log.info("시작합니다. 중지하려면 ESC 를 누르세요. (키보드/마우스는 건드리지 마세요)")
    esc_pressed()                            # 이전 키 상태 초기화

    done = ok = 0
    t0 = time.time()
    for o in todo:
        if esc_pressed():
            log.warning("ESC 감지 - 중지합니다.")
            break
        result = process(app, main_win, o)
        done += 1
        ok += result == "성공"
        log.info("(%d/%d) 행%d %s %s -> %s", done, len(todo), o.row, o.code, o.name, result)
        book.write(o.row, result)

    log.info("완료: 성공 %d / 실패·건너뜀 %d / %.1f초  -> %s", ok, done - ok, time.time() - t0, RESULT_PATH)


if __name__ == "__main__":
    main()
