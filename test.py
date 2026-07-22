import streamlit as st
import pandas as pd
import re

# ---------------------------------------------------------
# 페이지 기본 설정
# ---------------------------------------------------------
st.set_page_config(
    page_title="의약품 입고 정밀 검수 시스템",
    page_icon="💊",
    layout="wide"
)

st.title("💊 의약품 입고 정밀 검수 시스템 (Cloud Web App)")
st.caption("사내 마스터 DB와 제약사 입고 명세서를 교차 검증하여 바코드 스캔 검수를 진행합니다.")

# 세션 상태(Session State) 초기화 (페이지 새로고침 시 데이터 유지용)
if "master_db" not in st.session_state:
    st.session_state.master_db = {}
if "order_db" not in st.session_state:
    st.session_state.order_db = {}
if "logs" not in st.session_state:
    st.session_state.logs = []

# ---------------------------------------------------------
# 바코드 파싱 함수
# ---------------------------------------------------------
def parse_gs1_details(barcode_str):
    std_code = barcode_str[2:16] if barcode_str.startswith("01") else barcode_str[3:17]
    exp_date = "미확인"
    if "17" in barcode_str:
        idx = barcode_str.find("17")
        exp_date = f"20{barcode_str[idx+2:idx+4]}년 {barcode_str[idx+4:idx+6]}월"

    serial_num = "일련번호 없음"
    if "21" in barcode_str:
        idx = barcode_str.rfind("21")
        serial_num = barcode_str[idx+2:]

    return std_code, exp_date, serial_num

# ---------------------------------------------------------
# 사이드바: 엑셀 파일 업로드 영역
# ---------------------------------------------------------
with st.sidebar:
    st.header("📁 데이터 파일 업로드")
    
    # 1. 사내 마스터 파일
    master_file = st.file_uploader("1. 사내 마스터 엑셀 (.xlsx, .xls)", type=["xlsx", "xls"])
    if master_file and not st.session_state.master_db:
        try:
            df_m = pd.read_excel(master_file, usecols="A,C,H,L,Q,AS", dtype={'A': str, 'Q': str})
            df_m.columns = ["제품코드", "제품전산출력명", "제품규격", "매입처이름", "표준코드", "제약사명"]
            df_m = df_m.dropna(subset=["표준코드"])
            df_m["표준코드"] = df_m["표준코드"].astype(str).str.replace('.0', '', regex=False).str.strip().str.zfill(14)
            df_m["제품코드"] = df_m["제품코드"].astype(str).str.replace('.0', '', regex=False).str.strip().str.zfill(6)
            df_m = df_m.drop_duplicates(subset=["표준코드"], keep="first")
            
            st.session_state.master_db = df_m.set_index("표준코드").to_dict('index')
            st.success(f"마스터 DB 완료 ({len(st.session_state.master_db):,}건)")
        except Exception as e:
            st.error(f"마스터 로드 실패: {e}")

    # 2. 제약사 명세서 파일
    order_file = st.file_uploader("2. 제약사 명세서 엑셀 (.xlsx, .xls)", type=["xlsx", "xls"])
    if order_file and not st.session_state.order_db:
        try:
            df_o = pd.read_excel(order_file)
            header_idx = 4
            df_o.columns = df_o.iloc[header_idx].values
            df_o_data = df_o.iloc[header_idx + 1:].copy()
            
            df_clean = df_o_data[['표준코드', '수량']].dropna(subset=['표준코드']).copy()
            df_clean['표준코드'] = df_clean['표준코드'].astype(str).str.replace('.0', '', regex=False).str.strip().str.zfill(14)
            df_clean['수량'] = pd.to_numeric(df_clean['수량'], errors='coerce').fillna(0).astype(int)
            
            st.session_state.order_db.clear()
            for _, row in df_clean.iterrows():
                code = row['표준코드']
                qty = row['수량']
                if code in st.session_state.order_db:
                    st.session_state.order_db[code]["예정수량"] += qty
                else:
                    st.session_state.order_db[code] = {"예정수량": qty, "스캔수량": 0}
            
            st.success(f"명세서 완료 ({len(st.session_state.order_db):,}품목)")
        except Exception as e:
            st.error(f"명세서 로드 실패: {e}")

# ---------------------------------------------------------
# 메인 영역: 스캔 및 결과 확인
# ---------------------------------------------------------
col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("🔍 실시간 바코드 스캔")
    
    # 두 파일이 모두 로드되었을 때만 스캔 활성화
    if st.session_state.master_db and st.session_state.order_db:
        # Form을 이용하여 엔터 입력 시 자동 전송되도록 처리
        with st.form(key="scan_form", clear_on_submit=True):
            barcode_input = st.text_input("바코드를 스캔하고 Enter를 누르세요:", key="barcode")
            submit_button = st.form_submit_button(label="스캔 입력")
            
        if submit_button and barcode_input:
            std_code, exp_date, s_num = parse_gs1_details(barcode_input.strip())
            
            # 1차 검증: 마스터 DB
            if std_code not in st.session_state.master_db:
                msg = f"❌ [미등록] 사내 DB에 없는 표준코드 ({std_code})"
                st.session_state.logs.insert(0, ("error", msg))
            # 2차 검증: 입고 명세서
            elif std_code not in st.session_state.order_db:
                info = st.session_state.master_db[std_code]
                msg = f"⚠️ [오입고 경고] 금일 명세서에 없는 품목: {info['제품전산출력명']}"
                st.session_state.logs.insert(0, ("warning", msg))
            # 3차 검증: 수량 체크
            else:
                info = st.session_state.master_db[std_code]
                st.session_state.order_db[std_code]["스캔수량"] += 1
                expected = st.session_state.order_db[std_code]["예정수량"]
                scanned = st.session_state.order_db[std_code]["스캔수량"]
                drug_name = info['제품전산출력명']
                
                if scanned > expected:
                    msg = f"🚨 [초과 스캔] {drug_name} ({scanned}/{expected}개)"
                    st.session_state.logs.insert(0, ("error", msg))
                elif scanned == expected:
                    msg = f"✅ [검수 완료] {drug_name} ({scanned}/{expected}개)"
                    st.session_state.logs.insert(0, ("success", msg))
                else:
                    msg = f"🟢 [스캔 완료] {drug_name} ({scanned}/{expected}개 진행 중) | SN: {s_num}"
                    st.session_state.logs.insert(0, ("info", msg))
    else:
        st.warning("👈 좌측 사이드바에서 '사내 마스터'와 '제약사 명세서' 엑셀 파일을 먼저 업로드하세요.")

with col2:
    st.subheader("📋 스캔 기록 로그")
    for log_type, log_msg in st.session_state.logs[:15]: # 최근 15개 출력
        if log_type == "error":
            st.error(log_msg)
        elif log_type == "warning":
            st.warning(log_msg)
        elif log_type == "success":
            st.success(log_msg)
        else:
            st.info(log_msg)