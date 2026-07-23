from abc import ABC, abstractmethod
import pandas as pd

class BaseInvoiceParser(ABC):
    @abstractmethod
    def parse(self, file_obj) -> pd.DataFrame:
        """
        엑셀 파일을 읽어 반드시 아래 컬럼을 포함한 DataFrame을 반환해야 합니다.
        필수 컬럼: ['표준코드', '수량', '명세서제품명']
        """
        pass