# app.py

import streamlit as st
import pandas as pd
import datetime

# 모듈 임포트
from parsers.geo_young import GeoYoungParser
from core.scanner import parse_gs1_details
from core.report import generate_pdf_from_view_df

st.set_page_config(page_title="의약품 검수 시스템", page_icon="💊", layout="wide")

# ---------------------------------------------------------
# 화면 중앙 하단 보안 문구 및 제작자 서명 고정 (CSS/HTML)
# ---------------------------------------------------------
st.markdown(
    """
    <style>
    .footer-center-signature {
        position: fixed;
        bottom: 12px;
        left: 50%;
        transform: translateX(-50%);
        z-index: 9999;
        font-family: 'Pretendard', 'Arial', sans-serif;
        font-size: 12px;
        font-weight: bold;
        color: #C53030;
        background-color: rgba(255, 255, 255, 0.92);
        padding: 8px 18px;
        border-radius: 20px;
        border: 1px solid #FEB2B2;
        box-shadow: 0px 2px 8px rgba(0, 0, 0, 0.1);
        text-align: center;
        line-height: 1.4;
        white-space: nowrap;
    }
    </style>
    <div class="footer-center-signature">
        🚫 경동팜 전용 프로그램 외부 유출 금지<br>
        ✨ Made By JISU.K
    </div>
    """,
    unsafe_allow_html=True
)

# 세션 상태 초기화
if "master_db" not in st.session_state: st.session_state.master_db = {}
if "order_df" not in st.session_state: st.session_state.order_df = None
if "logs" not in st.session_state: st.session_state.logs = []
if "scanned_serials" not in st.session_state: st.session_state.scanned_serials = set()

with st.sidebar:
    st.header("📁 데이터 설정")
    
    # 1차 사내 마스터 파일 로드
    master_file = st.file_uploader("1. 사내 마스터 엑셀", type=["xlsx", "xls"])
    if master_file and not st.session_state.master_db:
        try:
            df_m = pd.read_excel(master_file, usecols="A,C,H,L,Q,AS", dtype={'A': str, 'Q': str})
            df_m.columns = ["제품코드", "제품전산출력명", "제품규격", "매입처이름", "표준코드", "제약사명"]
            df_m = df_m.dropna(subset=["표준코드"])
            df_m["표준코드"] = df_m["표준코드"].astype(str).str.replace('.0', '', regex=False).str.strip().str.zfill(14)
            st.session_state.master_db = df_m.drop_duplicates(subset=["표준코드"]).set_index("표준코드").to_dict('index')
            st.success("마스터 DB 완료")
        except Exception as e:
            st.error(f"마스터 로드 실패: {e}")

    # 2차 제약사 선택
    provider = st.selectbox("2. 제약사/도매처 선택", ["지오영", "추가예정..."])
    
    # 3차 명세서 파일 로드
    order_file = st.file_uploader("3. 명세서 엑셀", type=["xlsx", "xls"])
    
    if order_file and st.session_state.master_db and st.session_state.order_df is None:
        try:
            if provider == "지오영":
                parser = GeoYoungParser()
                parsed_df = parser.parse(order_file)

            list_data = []
            for idx, row in parsed_df.iterrows():
                std_code = row['표준코드']
                qty = row['수량']
                invoice_name = row['명세서제품명']
                
                master_info = st.session_state.master_db.get(std_code)
                if master_info:
                    prod_code = master_info["제품코드"]
                    prod_name = master_info["제품전산출력명"]
                    prod_spec = master_info["제품규격"]
                    is_registered = True
                else:
                    prod_code = "미등록"
                    prod_name = invoice_name
                    prod_spec = "-"
                    is_registered = False

                list_data.append({
                    "No": idx + 1,
                    "제품코드": prod_code,
                    "제품전산출력명": prod_name,
                    "제품규격": prod_spec,
                    "수량": qty,
                    "스캔수량": 0,
                    "표준코드": std_code,
                    "등록여부": is_registered
                })

            st.session_state.order_df = pd.DataFrame(list_data)
            st.session_state.scanned_serials.clear()
            st.success("명세서 로드 완료")
        except Exception as e:
            st.error(f"명세서 로드 실패: {e}")

    # 보고서 다운로드
    if st.session_state.order_df is not None:
        st.markdown("---")
        st.subheader("📄 보고서 출력")
        today_str = datetime.datetime.now().strftime("%Y-%m-%d")
        pdf_data = generate_pdf_from_view_df(st.session_state.order_df, today_str, provider)
        st.download_button("📥 PDF 다운로드", data=pdf_data, file_name=f"{today_str}_{provider}_입고검수.pdf", mime="application/pdf")

# 메인 화면
if st.session_state.master_db and st.session_state.order_df is not None:
    today_str = datetime.datetime.now().strftime("%Y-%m-%d")
    st.header(f"📊 [{today_str}] {provider} 입고 검수 현황")
    st.markdown("---")

    col_left, col_right = st.columns([1, 1.2])

    with col_left:
        # 1. 실시간 바코드 스캔 영역
        st.subheader("🔍 실시간 바코드 스캔")
        with st.form(key="scan_form", clear_on_submit=True):
            barcode_input = st.text_input("약품 바코드를 스캔하세요:", key="barcode")
            submit_button = st.form_submit_button(label="스캔")

        if submit_button and barcode_input:
            std_code, raw_barcode, is_otc = parse_gs1_details(barcode_input)
            target_idx = st.session_state.order_df.index[st.session_state.order_df["표준코드"] == std_code].tolist()
            
            if not target_idx:
                msg = f"❌ [미등록/오입고] 명세서에 없는 품목 (코드: {std_code})"
                st.session_state.logs.insert(0, ("error", msg))
            else:
                idx = target_idx[0]
                drug_name = st.session_state.order_df.loc[idx, "제품전산출력명"]
                
                if is_otc:
                    st.session_state.order_df.loc[idx, "스캔수량"] += 1
                    msg = f"✅ [일반약 스캔] {drug_name} (+1)"
                    st.session_state.logs.insert(0, ("success", msg))
                else:
                    if raw_barcode in st.session_state.scanned_serials:
                        msg = f"🚨 [중복 차단] 이미 검수된 박스입니다!\\n   - {drug_name}"
                        st.session_state.logs.insert(0, ("error", msg))
                    else:
                        st.session_state.scanned_serials.add(raw_barcode)
                        st.session_state.order_df.loc[idx, "스캔수량"] += 1
                        msg = f"🟢 [전문약 스캔] {drug_name} (S/N 확인완료)"
                        st.session_state.logs.insert(0, ("info", msg))
                st.rerun()

        st.markdown("---")

        # 2. 📦 대용량/수기 수량 조절 패널 (Enter 키 입력 지원)
        st.subheader("📦 대용량 번들 / 수기 수량 조절")
        with st.expander("👉 클릭하여 대량 수량 수기 증감하기", expanded=True):
            # 품목 선택 드롭다운 (No + 제품명)
            item_options = {
                f"[{row['No']}] {row['제품전산출력명']} (현재: {row['스캔수량']}/{row['수량']}개)": row['표준코드']
                for _, row in st.session_state.order_df.iterrows()
            }
            selected_label = st.selectbox("조절할 약품을 선택하세요:", list(item_options.keys()))
            selected_code = item_options[selected_label]
            
            # Form 구조로 변환하여 숫자 입력 후 Enter 키를 누르면 바로 수량이 업데이트되도록 적용
            with st.form(key="manual_adjust_form", clear_on_submit=False):
                m_col1, m_col2 = st.columns([1.5, 1])
                with m_col1:
                    adjust_qty = st.number_input("증감할 수량 입력 후 Enter:", value=10, step=1)
                with m_col2:
                    st.write("") # 간격 맞춤용
                    st.write("")
                    apply_btn = st.form_submit_button("수량 반영", use_container_width=True)

                if apply_btn:
                    target_idx = st.session_state.order_df.index[st.session_state.order_df["표준코드"] == selected_code].tolist()[0]
                    current_qty = st.session_state.order_df.loc[target_idx, "스캔수량"]
                    new_qty = current_qty + adjust_qty
                    
                    if new_qty < 0:
                        st.error("❌ 스캔 수량은 0개 미만이 될 수 없습니다.")
                    else:
                        st.session_state.order_df.loc[target_idx, "스캔수량"] = new_qty
                        drug_name = st.session_state.order_df.loc[target_idx, "제품전산출력명"]
                        sign = f"+{adjust_qty}" if adjust_qty > 0 else str(adjust_qty)
                        msg = f"📝 [수기 조절] {drug_name} ({sign}개 조절 ➔ 현재 {new_qty}개)"
                        st.session_state.logs.insert(0, ("info", msg))
                        st.rerun()

        st.markdown("---")

        # 3. 실시간 로그 표시 (최근 10개)
        st.subheader("📋 실시간 스캔 및 수기 로그")
        for log_type, log_msg in st.session_state.logs[:10]:
            if log_type == "error": st.error(log_msg)
            elif log_type == "success": st.success(log_msg)
            else: st.info(log_msg)

    with col_right:
        st.subheader("📋 [실시간] 입고 검수 대시보드")
        view_df = st.session_state.order_df[["No", "제품코드", "제품전산출력명", "제품규격", "수량", "스캔수량", "등록여부"]].copy()

        def highlight_status(row):
            if not row["등록여부"]:
                return ['background-color: #FFB6C1; color: black;'] * len(row) # 미등록(소프트핑크)
            expected, scanned = row["수량"], row["스캔수량"]
            if scanned == expected and expected > 0:
                return ['background-color: #D4EFDF; color: black;'] * len(row) # 초록 (완료)
            elif scanned > expected:
                return ['background-color: #FADBD8; color: black;'] * len(row) # 빨강 (초과)
            elif scanned > 0:
                return ['background-color: #FCF3CF; color: black;'] * len(row) # 노랑 (진행)
            return [''] * len(row)

        styled_df = view_df.style.apply(highlight_status, axis=1).hide(subset=["등록여부"], axis=1)
        st.dataframe(styled_df, use_container_width=True, height=650)

else:
    st.info("👈 좌측 사이드바에서 사내 마스터 파일과 명세서 파일을 업로드해주세요.")
