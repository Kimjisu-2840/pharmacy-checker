def parse_gs1_details(barcode_str):
    std_code = barcode_str[2:16] if barcode_str.startswith("01") else barcode_str[3:17]
    exp_date = "미확인"
    if "17" in barcode_str:
        idx = barcode_str.find("17")
        exp_date = f"20{barcode_str[idx+2:idx+4]}년 {barcode_str[idx+4:idx+6]}월"

    serial_num = None
    if "21" in barcode_str:
        idx = barcode_str.rfind("21")
        serial_num = barcode_str[idx+2:]

    return std_code, exp_date, serial_num