"""
Excel 파일 파싱 서비스 로직
"""

import logging
import os

import pandas as pd

logging = logging.getLogger(__name__)

class ExcelParseService:
    """
    Excel 파일 파싱 서비스
    """

    def __init__(self):
        return
        
    def read_excel(self, file_path):
        """
        Excel 파일을 읽어 DataFrame으로 변환
        
        Args:
            file_path(str): 엑셀 파일 경로
            
        Returns:
            pandas.DataFrame: 엑셀 데이터
        """
        try:
            logging.info(f"엑셀 파일 읽기 시작: {file_path}")
            return pd.read_excel(file_path)
        except Exception as e:
            logging.error(f"엑셀 파일 읽기 오류: {str(e)}")
            raise
        
    def write_excel(self, df, output_file):
        try:
            logging.info(f"엑셀 파일 저장 시작; {output_file}")
            output_path = os.path.join(self.output_folder, output_file)
            df.to_excel(output_path, index=False)
            logging.info(f"{output_file} 엑셀 파일 저장 완료")
        except Exception as e:
            logging.error(f"엑셀 파일 저장 오류: {str(e)}")
            raise
        
    def parse_data(self, file_name, column_num):
        """
        입력된 엑셀 파일 처리

        Args:
            file_name (str): 처리할 파일 이름
            column_num (int): 처리할 열 번호 (1부터 시작))
        """
        input_path = os.path.join(self.input_folder, file_name)
        
        # 엑셀 파일 읽기
        df = self.read_excel(input_path)
        
        # 사용자 입력 값 -> 인덱스로 변환
        column_idx = column_num - 1
        
        # 열 번호 유효성 검사
        if column_idx < 0 or column_idx >= len(df.columns):
            raise ValueError(f"요청된 열 번호 {column_num}이 유효하지 않습니다. 열 범위는 1-{len(df.columns)} 입니다.")
        
        # 해당 열의 고유 값 추출
        unique_values = df.iloc[:, column_idx].unique()
        logging.info(f"고유 값 {len(unique_values)}개 추출: {unique_values}")
        
        # 원본 파일명에서 확장자 추출
        # file_base: 파일명
        # file_ext: 확장자
        file_base, file_ext = os.path.splittext(file_name)
        
        # 각 고유 값에 대해 별도의 엑셀 파일 생성
        for value in unique_values:
            # 값이 NaN인 경우
            if pd.isna(value):
                value_str = "NA"
            else:
                value_str = str(value)
            
            # 파일명에 포함할 수 없는 문자 제거
            safe_value = "".join(c for c in value_str if c not in r'\/:*?"<>|')
            
            # 해당 값을 가진 데이터만 필터링
            filtered_df = df[df.iloc[:, column_idx] == value]
            
            # 엑셀 파일 저장
            self.write_excel(filtered_df, f"{file_base}_{safe_value}{file_ext}")
            

