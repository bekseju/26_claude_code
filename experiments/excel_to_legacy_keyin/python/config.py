"""
[설정] 실제 업무 프로그램에 적용할 때는 이 파일만 바꾸면 됩니다.
"""
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

# ---- 엑셀
EXCEL_PATH = ROOT / "data" / "orders.xlsx"            # 원본 (읽기 전용으로 취급)
RESULT_PATH = ROOT / "data" / "orders_result.xlsx"    # 처리결과를 기록할 사본
SHEET_NAME = "주문데이터"

# ---- 대상 프로그램
APP_CMD = (
    "powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden "
    f'-File "{ROOT / "mock_app" / "LegacyOrderApp.ps1"}"'
)
BACKEND = "win32"              # VB6/Delphi/MFC 는 win32, 안 보이면 "uia" 시도
LOGIN_TITLE_RE = r"^로그인.*"
MAIN_TITLE_RE = r"^주문 등록.*"
LOGIN_ID, LOGIN_PW = "admin", "1234"

# ---- 화면 구성
# Edit 컨트롤을 "위→아래, 왼→오른쪽" 으로 정렬했을 때의 순서.
# VB6 Label / Delphi TLabel 은 핸들이 없어 '라벨 이름으로 찾기'가 안 되므로 위치 순서로 식별합니다.
EDIT_ORDER = ["고객코드", "고객명", "전화번호", "주문일자", "수량", "단가", "금액", "비고"]
REGIONS = ["서울", "부산", "대구", "인천", "광주", "대전", "울산"]   # 지역 콤보 항목 순서
SUCCESS_TITLE = "확인"          # 저장 성공 메시지박스 제목

# ---- 대기 시간(초). 느린 프로그램이면 늘리세요.
WINDOW_TIMEOUT = 20
SAVE_TIMEOUT = 20
