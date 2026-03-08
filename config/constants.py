from decimal import Decimal

ORDER_NO_PADDING: int = 4
FALLBACK_CUSTOMER_CODE: str = "6038106627"
FALLBACK_CUSTOMER_NAME: str = "강남"
FALLBACK_ITEM_CODE: str = "A0001"
VAT_RATE: Decimal = Decimal("0.1")

ERP_PRIMARY_SHEET_NAME: str = "주문서입력"
ERP_HEADER_SCAN_MAX_ROW: int = 20

TRADE_META_LABELS: dict[str, tuple[str, ...]] = {
    "project_no": ("공사번호",),
    "order_no": ("발주번호",),
    "issue_date": ("작성일자", "일자", "날짜"),
    "due_date": ("납기일",),
    "customer_name": ("거래처", "발주처", "수주처", "공급받는자"),
}

TRADE_ITEM_HEADERS: tuple[str, ...] = (
    "순번",
    "POR NO",
    "자재번호",
    "품명 및 규격",
    "단위",
    "발주수량",
    "납품수량",
    "합격수량",
    "단가",
    "금액",
    "저장위치",
)

ERP_DYNAMIC_HEADERS: tuple[str, ...] = (
    "일자",
    "순번",
    "납기일자",
    "프로젝트명",
    "품목명",
    "규격",
    "수량",
    "단가",
    "공급가액",
    "부가세",
    "적요",
)

ERP_REQUIRED_HEADERS: tuple[str, ...] = (
    "일자",
    "순번",
    "거래처코드",
    "거래처명",
    "납기일자",
    "프로젝트명",
    "품목명",
    "규격",
    "수량",
    "단가",
    "공급가액",
    "부가세",
    "적요",
)
