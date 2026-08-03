# app.py

import streamlit as st
import pandas as pd
import datetime
import os
import streamlit.components.v1 as components

# 모듈 임포트
from parsers.geo_young import GeoYoungParser
from core.scanner import parse_gs1_details
from core.report import generate_excel_from_view_df

st.set_page_config(page_title="의약품 검수 시스템", page_icon="💊", layout="wide")

# history 폴더 생성 (과거 이력 저장용)
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
if "last_scanned" not in st.session_state: st.session_state.last_scanned = None

# ==========================================
# ★ 개선: 마스터 파일 내장 자동 로드 (절대 경로 고정)
# ==========================================
if not st.session_state.master_db:
    current_dir = os.path.dirname(os.path.abspath(__file__))
    master_path = os.path.join(current_dir, "data", "master_data.xlsx")
    
    try:
        df_m = pd.read_excel(master_path, usecols="A,C,H,L,Q,AS", dtype={'A': str, 'Q': str})
        df_m.columns = ["제품코드", "제품전산출력명", "제품규격", "매입처이름", "표준코드", "제약사명"]
        df_m = df_m.dropna(subset=["표준코드"])
        df_m["표준코드"] = df_m["표준코드"].astype(str).str.replace('.0', '', regex=False).str.strip().str.zfill(14)
        st.session_state.master_db = df_m.drop_duplicates(subset=["표준코드"]).set_index("표준코드").to_dict('index')
    except Exception as e:
        st.sidebar.error("⚠️ 마스터 엑셀 파일을 찾을 수 없습니다.")
        st.sidebar.warning(f"경로: {master_path}")

with st.sidebar:
    st.header("📁 데이터 및 업무 설정")
    # 사이드바에서 마스터 업로드 기능 제거됨 (자동 로드)
    
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
            st.session_state.last_scanned = None
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
                
            # ★ 개선 1: 스캔창 자동 포커스 유지 JS 삽입
            components.html(
                """
                <script>
                const doc = window.parent.document;
                const input = doc.querySelector('input[type="text"]');
                if (input) {
                    input.focus();
                }
                </script>
                """,
                height=0, width=0
            )

            if submit_button and barcode_input:
                std_code, raw_barcode, is_otc = parse_gs1_details(barcode_input)
                target_idx = st.session_state.order_df.index[st.session_state.order_df["표준코드"] == std_code].tolist()
                
                if not target_idx:
                    st.session_state.logs.insert(0, ("error", f"❌ [미등록/오입고] 명세서에 없는 품목 (코드: {std_code})"))
                    st.session_state.last_scanned = {"name": f"미등록 코드({std_code})", "status": "오입고 경고", "color": "red"}
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
                    
                    # 방금 스캔한 품목 상태 업데이트
                    curr_qty = st.session_state.order_df.loc[idx, "스캔수량"]
                    exp_qty = st.session_state.order_df.loc[idx, "수량"]
                    st.session_state.last_scanned = {
                        "name": drug_name,
                        "scanned": curr_qty,
                        "expected": exp_qty,
                        "status": "완료" if curr_qty == exp_qty else "진행중 (초과)" if curr_qty > exp_qty else "진행중",
                        "color": "green" if curr_qty == exp_qty else "red" if curr_qty > exp_qty else "orange"
                    }
                st.rerun()

            st.markdown("---")

            # 2. 수기 조절 패널
            st.subheader("📦 수기 수량 조절")
            with st.expander("👉 클릭하여 대량 수량 수기 증감하기", expanded=False):
                # ★ 개선 2: 빈 플레이스홀더를 첫 번째 항목으로 추가 (텍스트 지우기 방지)
                item_options = {"--- 약품명 또는 No 검색 ---": None}
                for _, row in st.session_state.order_df.iterrows():
                    label = f"[{row['No']}] {row['제품전산출력명']} (현재: {row['스캔수량']}/{row['수량']}개)"
                    item_options[label] = row['표준코드']
                
                selected_label = st.selectbox("조절할 약품을 검색/선택하세요:", list(item_options.keys()))
                
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
                        if selected_code is None:
                            st.error("❌ 변경할 약품을 선택해주세요.")
                        else:
                            target_idx = st.session_state.order_df.index[st.session_state.order_df["표준코드"] == selected_code].tolist()[0]
                            new_qty = st.session_state.order_df.loc[target_idx, "스캔수량"] + adjust_qty
                            
                            if new_qty < 0:
                                st.error("❌ 수량은 0 미만이 될 수 없습니다.")
                            else:
                                st.session_state.order_df.loc[target_idx, "스캔수량"] = new_qty
                                drug_name = st.session_state.order_df.loc[target_idx, '제품전산출력명']
                                st.session_state.logs.insert(0, ("info", f"📝 [수기 조절] {drug_name} ➔ 현재 {new_qty}개"))
                                
                                # 수기 조절 내용도 하이라이트 카드에 반영
                                st.session_state.last_scanned = {
                                    "name": f"(수기반영) {drug_name}",
                                    "scanned": new_qty,
                                    "expected": st.session_state.order_df.loc[target_idx, "수량"],
                                    "status": "조절완료",
                                    "color": "blue"
                                }
                                st.rerun()

            st.markdown("---")
            
            # ★ 개선 3: 실시간 로그 창 복원
            st.subheader("📋 실시간 스캔 로그")
            log_container = st.container()
            with log_container:
                if not st.session_state.logs:
                    st.caption("아직 스캔된 내역이 없습니다.")
                for log_type, log_msg in st.session_state.logs[:10]:
                    if log_type == "error": st.error(log_msg)
                    elif log_type == "success": st.success(log_msg)
                    else: st.info(log_msg)

        with col_right:
            # ★ 개선 4: 방금 스캔한 품목 하이라이트 카드 배치 (스크롤 찾기 방지)
            if st.session_state.last_scanned:
                ls = st.session_state.last_scanned
                color = ls.get('color', 'blue')
                st.markdown(
                    f"""
                    <div style="padding:15px; border-radius:10px; border:2px solid {color}; background-color:#f9f9f9; margin-bottom:15px;">
                        <h3 style="margin-top:0; color:{color};">📌 방금 스캔한 품목</h3>
                        <p style="font-size:18px; margin:5px 0;"><b>제품명:</b> {ls['name']}</p>
                        <p style="font-size:18px; margin:5px 0;"><b>수량:</b> <span style="font-size:22px; color:#E74C3C;">{ls.get('scanned', '-')}</span> / {ls.get('expected', '-')} 개 (<b>상태:</b> {ls['status']})</p>
                    </div>
                    """,
                    unsafe_allow_html=True
                )
            
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
            st.dataframe(styled_df, use_container_width=True, height=600)
            
            st.markdown("---")
            # 엑셀 다운로드 및 서버 저장 패널 (우측 하단 배치)
            st.subheader("🏁 검수 완료 및 보고서 저장")
            excel_bytes = generate_excel_from_view_df(st.session_state.order_df, today_str, provider)
            now_time = datetime.datetime.now().strftime("%H시%M분%S초")
            filename = f"{today_str}_{provider}_검수결과_{now_time}.xlsx"

            c_btn1, c_btn2 = st.columns(2)
            with c_btn1:
                st.download_button("📥 PC로 엑셀 다운로드", data=excel_bytes, file_name=filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            with c_btn2:
                if st.button("💾 사내 서버(이력 탭)에 저장", use_container_width=True):
                    with st.spinner("저장 중..."):
                        file_path = os.path.join("history", filename)
                        with open(file_path, "wb") as f:
                            f.write(excel_bytes)
                        st.success("✅ 저장 완료!")

    else:
        st.info("👈 명세서 엑셀 파일을 업로드해 주세요. (마스터 파일은 자동 내장되어 있습니다)")

# ----------------- 탭 2: 과거 엑셀 이력 조회 -----------------
with tab_history:
    st.header("📁 과거 검수 엑셀 이력 조회")
    history_files = sorted(os.listdir("history"), reverse=True)
    excel_files = [f for f in history_files if f.endswith('.xlsx')]
    
    if not excel_files:
        st.info("아직 저장된 엑셀 검수 이력이 없습니다.")
    else:
        for file in excel_files:
            file_path = os.path.join("history", file)
            with open(file_path, "rb") as f:
                excel_bytes_data = f.read()
            
            c1, c2 = st.columns([3, 1])
            with c1:
                st.write(f"📊 **{file}**")
            with c2:
                st.download_button(label="엑셀 다운로드", data=excel_bytes_data, file_name=file, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=file)
        st.markdown("---")
