import pandas as pd
from .base_parser import BaseInvoiceParser

class GeoYoungParser(BaseInvoiceParser):
    def parse(self, file_obj) -> pd.DataFrame:
        df = pd.read_excel(file_obj)
        header_idx = 4  # 지오영 엑셀은 5번째 줄(인덱스 4)이 헤더
        df.columns = df.iloc[header_idx].values
        df_data = df.iloc[header_idx + 1:].copy()
        
        # 품목명 열 이름 찾기 (엑셀 버전에 따라 다를 수 있으므로 유연하게 대처)
        name_col = '품목명' if '품목명' in df.columns else '제품명' if '제품명' in df.columns else df.columns[2]
        
        # 필요한 열만 추출
        df_clean = df_data[['표준코드', '수량', name_col]].dropna(subset=['표준코드']).copy()
        df_clean = df_clean.rename(columns={name_col: '명세서제품명'})
        
        df_clean['표준코드'] = df_clean['표준코드'].astype(str).str.replace('.0', '', regex=False).str.strip().str.zfill(14)
        df_clean['수량'] = pd.to_numeric(df_clean['수량'], errors='coerce').fillna(0).astype(int)
        
        # 중복 품목은 수량을 합산하여 반환
        grouped = df_clean.groupby(['표준코드', '명세서제품명'], as_index=False)['수량'].sum()
        
        return grouped