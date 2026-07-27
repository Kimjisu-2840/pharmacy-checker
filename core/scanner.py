# core/scanner.py

def parse_gs1_details(barcode_str):
    """
    바코드를 분석하여 (표준코드, 원본바코드, 일반약여부)를 반환합니다.
    """
    raw_barcode = barcode_str.strip()

    # 1. 표준코드 추출
    # 전문의약품 (GS1-128 / DataMatrix, '01'로 시작)
    if raw_barcode.startswith("01") and len(raw_barcode) >= 16:
        std_code = raw_barcode[2:16]
    # 일반의약품 (EAN-13, 보통 '880'으로 시작하는 13자리)
    else:
        # 사내 마스터 DB가 14자리(앞에 0 포함) 기준이므로 길이를 맞춰줍니다.
        std_code = raw_barcode[:14].zfill(14)

    # 2. 일반약(OTC) 판별 로직
    # 바코드 전체 길이가 14자리 이하이거나, 일련번호 식별자('21')가 없으면 일반약으로 간주
    is_otc = False
    if len(raw_barcode) <= 14 or "21" not in raw_barcode:
        is_otc = True

    # 파싱된 결과가 아닌 '원본 텍스트 자체'를 그대로 반환 (중복 비교용)
    return std_code, raw_barcode, is_otc
