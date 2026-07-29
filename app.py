# app.py

import streamlit as st
import pandas as pd
import datetime
import os

# 모듈 임포트
from parsers.geo_young import GeoYoungParser
from core.scanner import parse_gs1_details
from core.report import generate_excel_from_view_df  # 엑셀 엔진 로드

st.set_page_config(page_title="의약품 검수 시스템", page_icon="💊", layout="wide")

# history 폴더 생성 (과거 내역 저장용)
os.makedirs("history", exist_ok=True)

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

# 백그라운드 마스터 데이터 자동 로드 (data/master_data.xlsx)
if not st.session_state.master_db:
    try:
        df_m = pd.read_excel("data/master_data.xlsx", usecols="A,C,H,L,Q,AS", dtype={'A': str, 'Q': str})
        df_m.columns = ["제품코드", "제품전산출력명", "제품규격", "매입처이름", "표준코드", "제약사명"]
        df_m = df_m.dropna(subset=["표준코드"])
        df_m["표준코드"] = df_m["표준코드"].astype(str).str.replace('.0', '', regex=False).str.strip().str.zfill(14)
        st.session_state.master_db = df_m.drop_duplicates(subset=["표준코드"]).set_index("표준코드").to_dict('index')
    except Exception as e:
        st.sidebar.error("⚠️ data/master_data.xlsx 파일을 찾을 수 없습니다.")

with st.sidebar:
    st.header("📁 업무 설정")
    provider = st.selectbox("1. 제약사/도매처 선택", ["지오영", "추가예정..."])
    order_file = st.file_uploader("2. 명세서 엑셀 업로드", type=["xlsx", "xls"])
    
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
                    "No": idx + 1, "제품코드": prod_code, "제품전산출력명": prod_name,
                    "제품규격": prod_spec, "수량": qty, "스캔수량": 0, "표준코드": std_code, "등록여부": is_registered
                })

            st.session_state.order_df = pd.DataFrame(list_data)
            st.session_state.scanned_serials.clear()
            st.success("명세서 로드 완료")
        except Exception as e:
            st.error(f"명세서 로드 실패: {e}")

# ==========================================
# 탭(Tab) UI 구성
# ==========================================
tab_main, tab_history = st.tabs(["🔍 실시간 입고 검수", "📁 과거 검수 엑셀 이력"])

# ----------------- 탭 1: 메인 검수 화면 -----------------
with tab_main:
    if st.session_state.master_db and st.session_state.order_df is not None:
        today_str = datetime.datetime.now().strftime("%Y-%m-%d")
        st.header(f"📊 [{today_str}] {provider} 입고 검수 현황")
        st.markdown("---")

        col_left, col_right = st.columns([1, 1.2])

        with col_left:
            # 1. 스캔 영역
            st.subheader("🔍 바코드 스캔")
            with st.form(key="scan_form", clear_on_submit=True):
                barcode_input = st.text_input("약품 바코드를 스캔하세요:", key="barcode")
                submit_button = st.form_submit_button(label="스캔")

            if submit_button and barcode_input:
                std_code, raw_barcode, is_otc = parse_gs1_details(barcode_input)
                target_idx = st.session_state.order_df.index[st.session_state.order_df["표준코드"] == std_code].tolist()
                
                if not target_idx:
                    st.session_state.logs.insert(0, ("error", f"❌ [미등록/오입고] 명세서에 없는 품목 (코드: {std_code})"))
                else:
                    idx = target_idx[0]
                    drug_name = st.session_state.order_df.loc[idx, "제품전산출력명"]
                    if is_otc:
                        st.session_state.order_df.loc[idx, "스캔수량"] += 1
                        st.session_state.logs.insert(0, ("success", f"✅ [일반약 스캔] {drug_name} (+1)"))
                    else:
                        if raw_barcode in st.session_state.scanned_serials:
                            st.session_state.logs.insert(0, ("error", f"🚨 [중복 차단] 이미 검수된 박스입니다!\n   - {drug_name}"))
                        else:
                            st.session_state.scanned_serials.add(raw_barcode)
                            st.session_state.order_df.loc[idx, "스캔수량"] += 1
                            st.session_state.logs.insert(0, ("info", f"🟢 [전문약 스캔] {drug_name} (S/N 확인완료)"))
                    st.rerun()

            st.markdown("---")

            # 2. 수기 조절 패널
            st.subheader("📦 수기 수량 조절")
            with st.expander("👉 클릭하여 대량 수량 수기 증감하기", expanded=False):
                item_options = {f"[{row['No']}] {row['제품전산출력명']} (현재: {row['스캔수량']}/{row['수량']}개)": row['표준코드'] for _, row in st.session_state.order_df.iterrows()}
                selected_label = st.selectbox("조절할 약품을 선택하세요:", list(item_options.keys()))
                
                with st.form(key="manual_adjust_form", clear_on_submit=False):
                    m_col1, m_col2 = st.columns([1.5, 1])
                    with m_col1:
                        adjust_qty = st.number_input("증감할 수량 입력 후 Enter:", value=10, step=1)
                    with m_col2:
                        st.write("")
                        st.write("")
                        apply_btn = st.form_submit_button("수량 반영", use_container_width=True)

                    if apply_btn:
                        selected_code = item_options[selected_label]
                        target_idx = st.session_state.order_df.index[st.session_state.order_df["표준코드"] == selected_code].tolist()[0]
                        new_qty = st.session_state.order_df.loc[target_idx, "스캔수량"] + adjust_qty
                        
                        if new_qty < 0:
                            st.error("❌ 수량은 0 미만이 될 수 없습니다.")
                        else:
                            st.session_state.order_df.loc[target_idx, "스캔수량"] = new_qty
                            st.session_state.logs.insert(0, ("info", f"📝 [수기 조절] {st.session_state.order_df.loc[target_idx, '제품전산출력명']} ➔ 현재 {new_qty}개"))
                            st.rerun()

            st.markdown("---")

            # 3. 엑셀 즉시 다운로드 및 서버 저장 패널 (이메일 제거됨)
            st.subheader("🏁 검수 완료 및 보고서 저장")
            with st.expander("검수 결과 엑셀 저장 및 다운로드", expanded=True):
                excel_bytes = generate_excel_from_view_df(st.session_state.order_df, today_str, provider)
                # 파일명에 시분초를 추가하여 같은 날 여러번 저장해도 덮어써지지 않게 방지
                now_time = datetime.datetime.now().strftime("%시%분%초")
                filename = f"{today_str}_{provider}_검수결과_{now_time}.xlsx"

                c_btn1, c_btn2 = st.columns(2)
                with c_btn1:
                    st.download_button(
                        "📥 내 PC로 즉시 다운로드",
                        data=excel_bytes,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

                with c_btn2:
                    if st.button("💾 사내 서버(과거 이력)에 저장", use_container_width=True):
                        with st.spinner("엑셀 파일을 서버에 저장 중..."):
                            file_path = os.path.join("history", filename)
                            with open(file_path, "wb") as f:
                                f.write(excel_bytes)
                            st.success("✅ 서버 저장 완료! 상단의 [과거 검수 엑셀 이력] 탭에서 누구나 볼 수 있습니다.")

        with col_right:
            st.subheader("📋 실시간 대시보드")
            view_df = st.session_state.order_df[["No", "제품코드", "제품전산출력명", "수량", "스캔수량", "등록여부"]].copy()

            def highlight_status(row):
                if not row["등록여부"]: return ['background-color: #FFB6C1; color: black;'] * len(row)
                expected, scanned = row["수량"], row["스캔수량"]
                if scanned == expected and expected > 0: return ['background-color: #D4EFDF; color: black;'] * len(row)
                elif scanned > expected: return ['background-color: #FADBD8; color: black;'] * len(row)
                elif scanned > 0: return ['background-color: #FCF3CF; color: black;'] * len(row)
                return [''] * len(row)

            styled_df = view_df.style.apply(highlight_status, axis=1).hide(subset=["등록여부"], axis=1)
            st.dataframe(styled_df, use_container_width=True, height=700)

    else:
        st.info("👈 좌측에서 명세서 파일을 업로드하시면 검수가 시작됩니다.")

# ----------------- 탭 2: 과거 엑셀 이력 조회 -----------------
with tab_history:
    st.header("📁 과거 검수 엑셀 이력 조회")
    st.markdown("이곳에서 이전에 저장된 모든 검수 결과 엑셀 파일(.xlsx)을 다시 조회하고 다운로드할 수 있습니다.")
    
    # history 폴더 내의 파일 목록을 불러옴 (최신순 정렬)
    history_files = sorted(os.listdir("history"), reverse=True)
    excel_files = [f for f in history_files if f.endswith('.xlsx')]
    
    if not excel_files:
        st.info("아직 저장된 엑셀 검수 이력이 없습니다. 검수를 진행하고 저장해 주세요.")
    else:
        for file in excel_files:
            file_path = os.path.join("history", file)
            with open(file_path, "rb") as f:
                excel_bytes_data = f.read()
            
            c1, c2 = st.columns([3, 1])
            with c1:
                st.write(f"📊 **{file}**")
            with c2:
                st.download_button(
                    label="엑셀 다운로드",
                    data=excel_bytes_data,
                    file_name=file,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=file
                )
        st.markdown("---")
