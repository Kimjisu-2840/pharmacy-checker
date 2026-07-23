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
        color: #C53030; /* 보안 경고용 딥 레드 톤 */
        background-color: rgba(255, 255, 255, 0.92);
        padding: 8px 18px;
        border-radius: 20px;
        border: 1px solid #FEB2B2;
        box-shadow: 0px 2px 8px rgba(0, 0, 0, 0.1);
        text-align: center; /* 2줄 문구 중앙 정렬 */
        line-height: 1.4;
        white-space: nowrap; /* 문구 줄바꿈 방지 */
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
if "scanned_serials" not in st.session_state: st.session_state.scanned_serials = set() # [요구사항 1] 중복 필터셋

with st.sidebar:
    st.header("📁 데이터 설정")
    
    # 사내 마스터 파일 로드
    master_file = st.file_uploader("1. 사내 마스터 엑셀", type=["xlsx", "xls"])
    if master_file and not st.session_state.master_db:
        df_m = pd.read_excel(master_file, usecols="A,C,H,L,Q,AS", dtype={'A': str, 'Q': str})
        df_m.columns = ["제품코드", "제품전산출력명", "제품규격", "매입처이름", "표준코드", "제약사명"]
        df_m = df_m.dropna(subset=["표준코드"])
        df_m["표준코드"] = df_m["표준코드"].astype(str).str.replace('.0', '', regex=False).str.strip().str.zfill(14)
        st.session_state.master_db = df_m.drop_duplicates(subset=["표준코드"]).set_index("표준코드").to_dict('index')
        st.success("마스터 DB 완료")

    # 제약사 선택 (전략 패턴 적용)
    provider = st.selectbox("2. 제약사/도매처 선택", ["지오영", "추가예정..."])
    
    # 명세서 파일 로드
    order_file = st.file_uploader("3. 명세서 엑셀", type=["xlsx", "xls"])
    
    if order_file and st.session_state.master_db and st.session_state.order_df is None:
        try:
            # 선택된 제약사에 맞는 파서 객체 생성 및 실행
            if provider == "지오영":
                parser = GeoYoungParser()
                parsed_df = parser.parse(order_file)

            list_data = []
            for idx, row in parsed_df.iterrows():
                std_code = row['표준코드']
                qty = row['수량']
                invoice_name = row['명세서제품명']
                
                # [요구사항 2] 마스터 DB에 없으면 명세서 이름 사용 및 미등록 처리
                master_info = st.session_state.master_db.get(std_code)
                if master_info:
                    prod_code = master_info["제품코드"]
                    prod_name = master_info["제품전산출력명"]
                    prod_spec = master_info["제품규격"]
                    is_registered = True
                else:
                    prod_code = "미등록"
                    prod_name = invoice_name # 명세서 원본 이름 사용
                    prod_spec = "-"
                    is_registered = False

                list_data.append({
                    "No": idx + 1, "제품코드": prod_code, "제품전산출력명": prod_name,
                    "제품규격": prod_spec, "수량": qty, "스캔수량": 0,
                    "표준코드": std_code, "등록여부": is_registered
                })

            st.session_state.order_df = pd.DataFrame(list_data)
            st.session_state.scanned_serials.clear()
            st.success("명세서 로드 완료")
        except Exception as e:
            st.error(f"명세서 로드 실패: {e}")

# 메인 화면
if st.session_state.master_db and st.session_state.order_df is not None:
    today_str = datetime.datetime.now().strftime("%Y-%m-%d")
    st.header(f"📊 [{today_str}] {provider} 입고 검수 현황")
    st.markdown("---")

    col_left, col_right = st.columns([1, 1.2])

    with col_left:
        st.subheader("🔍 실시간 바코드 스캔")
        with st.form(key="scan_form", clear_on_submit=True):
            barcode_input = st.text_input("약품 바코드를 스캔하세요:", key="barcode")
            submit_button = st.form_submit_button(label="스캔")

        if submit_button and barcode_input:
            std_code, exp_date, s_num = parse_gs1_details(barcode_input.strip())
            
            target_idx = st.session_state.order_df.index[st.session_state.order_df["표준코드"] == std_code].tolist()
            
            if not target_idx:
                msg = f"❌ [오입고/미등록] 명세서에 없는 품목입니다. (코드: {std_code})"
                st.session_state.logs.insert(0, ("error", msg))
            else:
                # [요구사항 1] 일련번호 중복 필터 기능
                unique_key = f"{std_code}_{s_num}" if s_num else None
                if unique_key and unique_key in st.session_state.scanned_serials:
                    msg = f"🚨 [중복 차단] 이미 검수된 박스입니다! (S/N: {s_num})"
                    st.session_state.logs.insert(0, ("error", msg))
                else:
                    if unique_key:
                        st.session_state.scanned_serials.add(unique_key)
                    
                    idx = target_idx[0]
                    st.session_state.order_df.loc[idx, "스캔수량"] += 1
                    
                    drug_name = st.session_state.order_df.loc[idx, "제품전산출력명"]
                    scanned = st.session_state.order_df.loc[idx, "스캔수량"]
                    expected = st.session_state.order_df.loc[idx, "수량"]
                    
                    msg = f"✅ 스캔 완료: {drug_name} ({scanned}/{expected})"
                    st.session_state.logs.insert(0, ("success", msg))
                
                st.rerun()

        for log_type, log_msg in st.session_state.logs[:5]:
            if log_type == "error": st.error(log_msg)
            else: st.success(log_msg)

    with col_right:
        st.subheader("📋 대시보드")
        
        view_df = st.session_state.order_df[["No", "제품코드", "제품전산출력명", "제품규격", "수량", "스캔수량", "등록여부"]].copy()

        # [요구사항 2] 미등록 셀 소프트 핑크 하이라이트 및 상태 색상
        def highlight_status(row):
            if not row["등록여부"]:
                return ['background-color: #FFB6C1; color: black;'] * len(row) # Soft Pink
            
            expected, scanned = row["수량"], row["스캔수량"]
            if scanned == expected and expected > 0:
                return ['background-color: #D4EFDF; color: black;'] * len(row) # 초록 (완료)
            elif scanned > expected:
                return ['background-color: #FADBD8; color: black;'] * len(row) # 빨강 (초과)
            elif scanned > 0:
                return ['background-color: #FCF3CF; color: black;'] * len(row) # 노랑 (진행)
            return [''] * len(row)

        # 표시할 때는 '등록여부' 컬럼은 숨김 처리
        styled_df = view_df.style.apply(highlight_status, axis=1).hide(subset=["등록여부"], axis=1)
        st.dataframe(styled_df, use_container_width=True, height=600)
