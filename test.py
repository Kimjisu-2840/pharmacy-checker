import streamlit as st
import pandas as pd
import datetime
import io

# ReportLab 라이브러리 (PDF 생성용)
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# ---------------------------------------------------------
# 페이지 기본 설정
# ---------------------------------------------------------
st.set_page_config(
    page_title="의약품 입고 정밀 검수 시스템",
    page_icon="💊",
    layout="wide"
)

# ---------------------------------------------------------
# 화면 우측 하단 제작자 서명 고정 CSS (MADE BY JISU.KIM)
# ---------------------------------------------------------
st.markdown(
    """
    <style>
    .footer-signature {
        position: fixed;
        bottom: 12px;
        right: 20px;
        z-index: 9999;
        font-family: 'Arial', sans-serif;
        font-size: 13px;
        font-weight: bold;
        color: #4A5568;
        background-color: rgba(255, 255, 255, 0.85);
        padding: 6px 14px;
        border-radius: 20px;
        border: 1px solid #E2E8F0;
        box-shadow: 0px 2px 6px rgba(0, 0, 0, 0.08);
        letter-spacing: 1px;
    }
    </style>
    <div class="footer-signature">
        ✨ MADE BY JISU.KIM
    </div>
    """,
    unsafe_allow_html=True
)

# 세션 상태(Session State) 초기화
if "master_db" not in st.session_state:
    st.session_state.master_db = {}
if "order_df" not in st.session_state:
    st.session_state.order_df = None
if "logs" not in st.session_state:
    st.session_state.logs = []

# ---------------------------------------------------------
# 바코드 파싱 함수 (GS1-128 / DataMatrix)
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
# 화면 표(A~F열) 데이터 1:1 매핑 PDF 생성 함수
# ---------------------------------------------------------
def generate_pdf_from_view_df(order_df, today_str):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=25, leftMargin=25, topMargin=30, bottomMargin=30)
    story = []

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'TitleStyle',
        parent=styles['Heading1'],
        fontSize=16,
        spaceAfter=15,
        alignment=1
    )

    story.append(Paragraph(f"[{today_str}] 지오영 입고 검수 현황 보고서", title_style))
    story.append(Spacer(1, 10))

    table_data = [["No", "제품코드", "제품전산출력명", "제품규격", "수량", "스캔수량", "검수상태"]]

    total_expected = 0
    total_scanned = 0

    for _, row in order_df.iterrows():
        expected = int(row["수량"])
        scanned = int(row["스캔수량"])
        total_expected += expected
        total_scanned += scanned
        
        if scanned == expected and expected > 0:
            status = "완료"
        elif scanned < expected:
            status = f"미달(-{expected - scanned})"
        else:
            status = f"초과(+{scanned - expected})"

        table_data.append([
            str(row["No"]),
            str(row["제품코드"]),
            str(row["제품전산출력명"])[:16],
            str(row["제품규격"]),
            str(expected),
            str(scanned),
            status
        ])

    summary_status = "검수 완료" if total_scanned == total_expected else "진행 중 / 불일치"
    table_data.append(["총계", "-", "-", "-", str(total_expected), str(total_scanned), summary_status])

    pdf_table = Table(table_data, colWidths=[30, 65, 180, 80, 50, 55, 75])
    pdf_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#2C3E50")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor("#EAEDED")),
        ('TEXTCOLOR', (0, -1), (-1, -1), colors.black),
    ]))

    story.append(pdf_table)
    doc.build(story)
    buffer.seek(0)
    return buffer

# ---------------------------------------------------------
# 사이드바: 1차/2차 엑셀 업로드 & PDF 다운로드
# ---------------------------------------------------------
with st.sidebar:
    st.header("📁 데이터 파일 업로드")
    
    master_file = st.file_uploader("1차: 사내 마스터 엑셀 (.xlsx, .xls)", type=["xlsx", "xls"])
    if master_file and not st.session_state.master_db:
        try:
            df_m = pd.read_excel(master_file, usecols="A,C,H,L,Q,AS", dtype={'A': str, 'Q': str})
            df_m.columns = ["제품코드", "제품전산출력명", "제품규격", "매입처이름", "표준코드", "제약사명"]
            df_m = df_m.dropna(subset=["표준코드"])
            df_m["표준코드"] = df_m["표준코드"].astype(str).str.replace('.0', '', regex=False).str.strip().str.zfill(14)
            df_m["제품코드"] = df_m["제품코드"].astype(str).str.replace('.0', '', regex=False).str.strip().str.zfill(6)
            df_m = df_m.drop_duplicates(subset=["표준코드"], keep="first")
            
            st.session_state.master_db = df_m.set_index("표준코드").to_dict('index')
            st.success(f"✅ 마스터 DB 준비 완료 ({len(st.session_state.master_db):,}건)")
        except Exception as e:
            st.error(f"마스터 로드 실패: {e}")

    order_file = st.file_uploader("2차: 제약사 명세서 엑셀 (.xlsx, .xls)", type=["xlsx", "xls"])
    if order_file and st.session_state.master_db and st.session_state.order_df is None:
        try:
            df_o = pd.read_excel(order_file)
            header_idx = 4
            df_o.columns = df_o.iloc[header_idx].values
            df_o_data = df_o.iloc[header_idx + 1:].copy()
            
            df_clean = df_o_data[['표준코드', '수량']].dropna(subset=['표준코드']).copy()
            df_clean['표준코드'] = df_clean['표준코드'].astype(str).str.replace('.0', '', regex=False).str.strip().str.zfill(14)
            df_clean['수량'] = pd.to_numeric(df_clean['수량'], errors='coerce').fillna(0).astype(int)
            
            grouped_order = df_clean.groupby('표준코드', as_index=False)['수량'].sum()

            list_data = []
            no = 1
            for _, row in grouped_order.iterrows():
                std_code = row['표준코드']
                qty = row['수량']
                
                master_info = st.session_state.master_db.get(std_code, {})
                prod_code = master_info.get("제품코드", "미등록")
                prod_name = master_info.get("제품전산출력명", "마스터 미등록 품목")
                prod_spec = master_info.get("제품규격", "-")

                list_data.append({
                    "No": no,
                    "제품코드": prod_code,
                    "제품전산출력명": prod_name,
                    "제품규격": prod_spec,
                    "수량": qty,
                    "스캔수량": 0,
                    "표준코드": std_code
                })
                no += 1

            st.session_state.order_df = pd.DataFrame(list_data)
            st.success(f"✅ 명세서 로드 완료 ({len(st.session_state.order_df)}품목)")
        except Exception as e:
            st.error(f"명세서 로드 실패: {e}")

    if st.session_state.order_df is not None:
        st.markdown("---")
        st.subheader("📄 검수 보고서 출력")
        today_str = datetime.datetime.now().strftime("%Y-%m-%d")
        
        pdf_data = generate_pdf_from_view_df(st.session_state.order_df, today_str)
        
        st.download_button(
            label="📥 화면 대시보드 표 PDF 다운로드",
            data=pdf_data,
            file_name=f"{today_str}_지오영_입고검수현황.pdf",
            mime="application/pdf"
        )

# ---------------------------------------------------------
# 메인 화면 영역
# ---------------------------------------------------------
today_str = datetime.datetime.now().strftime("%Y-%m-%d")

if st.session_state.master_db and st.session_state.order_df is not None:
    st.header(f"📊 [{today_str}] 지오영 입고 검수 현황")
    
    total_items = len(st.session_state.order_df)
    total_expected = st.session_state.order_df["수량"].sum()
    total_scanned = st.session_state.order_df["스캔수량"].sum()
    
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("총 입고 예정 품목", f"{total_items:,} 품목")
    k2.metric("총 입고 예정 수량", f"{total_expected:,} 개")
    k3.metric("현재 총 스캔 수량", f"{total_scanned:,} 개")
    k4.metric("검수 진행률", f"{(total_scanned/total_expected*100) if total_expected > 0 else 0:.1f} %")

    st.markdown("---")

    col_left, col_right = st.columns([1, 1.2])

    with col_left:
        st.subheader("🔍 실시간 바코드 스캔")
        with st.form(key="scan_form", clear_on_submit=True):
            barcode_input = st.text_input("약품 바코드를 스캔하고 Enter를 누르세요:", key="barcode")
            submit_button = st.form_submit_button(label="스캔 입력")

        if submit_button and barcode_input:
            std_code, exp_date, s_num = parse_gs1_details(barcode_input.strip())
            
            if std_code not in st.session_state.master_db:
                msg = f"❌ [미등록 오류] 사내 마스터 DB에 없는 약품입니다. ({std_code})"
                st.session_state.logs.insert(0, ("error", msg))
            else:
                target_idx = st.session_state.order_df.index[st.session_state.order_df["표준코드"] == std_code].tolist()
                
                if not target_idx:
                    info = st.session_state.master_db[std_code]
                    msg = f"⚠️ [오입고 경고] 금일 지오영 명세서에 없는 품목입니다!\n   - 약품명: {info['제품전산출력명']}"
                    st.session_state.logs.insert(0, ("warning", msg))
                else:
                    idx = target_idx[0]
                    st.session_state.order_df.loc[idx, "스캔수량"] += 1
                    
                    drug_name = st.session_state.order_df.loc[idx, "제품전산출력명"]
                    expected = st.session_state.order_df.loc[idx, "수량"]
                    scanned = st.session_state.order_df.loc[idx, "스캔수량"]
                    
                    if scanned > expected:
                        msg = f"🚨 [초과 스캔] {drug_name} ({scanned}/{expected}개)"
                        st.session_state.logs.insert(0, ("error", msg))
                    elif scanned == expected:
                        msg = f"✅ [검수 완료] {drug_name} ({expected}/{expected}개 입고 완료)"
                        st.session_state.logs.insert(0, ("success", msg))
                    else:
                        msg = f"🟢 [스캔 완료] {drug_name} ({scanned}/{expected}개 진행 중) | SN: {s_num}"
                        st.session_state.logs.insert(0, ("info", msg))
                    
                    st.rerun()

        st.subheader("📋 실시간 스캔 로그")
        for log_type, log_msg in st.session_state.logs[:10]:
            if log_type == "error": st.error(log_msg)
            elif log_type == "warning": st.warning(log_msg)
            elif log_type == "success": st.success(log_msg)
            else: st.info(log_msg)

    with col_right:
        st.subheader("📋 [실시간] 입고 검수 대시보드")
        
        view_df = st.session_state.order_df[["No", "제품코드", "제품전산출력명", "제품규격", "수량", "스캔수량"]].copy()

        def highlight_status(row):
            expected = row["수량"]
            scanned = row["스캔수량"]
            if scanned == expected and expected > 0:
                return ['background-color: #D4EFDF; color: black;'] * len(row)
            elif scanned > expected:
                return ['background-color: #FADBD8; color: black;'] * len(row)
            elif scanned > 0:
                return ['background-color: #FCF3CF; color: black;'] * len(row)
            return [''] * len(row)

        styled_df = view_df.style.apply(highlight_status, axis=1)
        st.dataframe(styled_df, use_container_width=True, height=600)

else:
    st.title("💊 의약품 입고 정밀 검수 시스템")
    st.info("👈 좌측 사이드바에서 [1차: 사내 마스터 엑셀]과 [2차: 제약사 명세서 엑셀]을 차례로 업로드해주세요.")
