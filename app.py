# app.py 상단 마스터 로드 부분 교체

import os

if not st.session_state.master_db:
    current_dir = os.path.dirname(os.path.abspath(__file__))
    master_path = os.path.join(current_dir, "data", "master_data.xlsx")
    
    st.sidebar.markdown("### 🔍 서버 내부 파일 진단")
    st.sidebar.write(f"**1. 현재 실행 경로:** `{current_dir}`")
    
    # 루트 폴더 파일 목록 확인
    if os.path.exists(current_dir):
        st.sidebar.write(f"**2. 최상위 폴더 목록:** {os.listdir(current_dir)}")
    
    # data 폴더 존재 여부 및 내부 파일 확인
    data_dir = os.path.join(current_dir, "data")
    if os.path.exists(data_dir):
        st.sidebar.write(f"**3. data 폴더 내부 파일:** {os.listdir(data_dir)}")
    else:
        st.sidebar.error("❌ 서버에 'data' 폴더 자체가 존재하지 않습니다!")

    try:
        df_m = pd.read_excel(master_path, usecols="A,C,H,L,Q,AS", dtype={'A': str, 'Q': str})
        # ... (이하 동일한 로드 로직)
        st.sidebar.success("✅ 마스터 로드 성공!")
    except Exception as e:
        st.sidebar.error(f"⚠️ 로드 실패 원인: {e}")
