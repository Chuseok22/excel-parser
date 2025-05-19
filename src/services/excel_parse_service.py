"""
Excel 파일 파싱 서비스 로직
"""

import logging
import os
import json
import csv

import pandas as pd

logging = logging.getLogger(__name__)

class ExcelParseService:
    """
    Excel 파일 파싱 서비스
    """

    def __init__(self, input_folder=None, output_folder=None):
        """
        ExcelParseService 초기화
        
        Args:
            input_folder (str): 입력 엑셀 파일 폴더 경로
            output_folder (str): 출력 엑셀 파일 폴더 경로
        """
        self.input_folder = input_folder or os.path.join(os.getcwd(), "input")
        self.output_folder = output_folder or os.path.join(os.getcwd(), "output")
        
        # 출력 폴더가 없으면 생성
        if not os.path.exists(self.output_folder):
            os.makedirs(self.output_folder, exist_ok=True)
        
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
        file_base, file_ext = os.path.splitext(file_name)
        
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
            output_file = f"{file_base}_{safe_value}{file_ext}"
            self.write_excel(filtered_df, output_file)
        
        # 생성된 파일 경로 반환
        output_files = []
        for value in unique_values:
            # 값이 NaN인 경우
            if pd.isna(value):
                value_str = "NA"
            else:
                value_str = str(value)
            
            # 파일명에 포함할 수 없는 문자 제거
            safe_value = "".join(c for c in value_str if c not in r'\/:*?"<>|')
            
            # 출력 파일 경로 추가
            output_files.append(os.path.join(self.output_folder, f"{file_base}_{safe_value}{file_ext}"))
        
        return output_files
        
    def parse_excel(self, file_path, progress_callback=None):
        """
        엑셀 파일의 헤더를 파싱하는 기능
        
        Args:
            file_path (str): 엑셀 파일 경로
            progress_callback (callable): 진행률을 보고할 콜백 함수
            
        Returns:
            dict: {헤더명: {'data_type': 데이터타입, 'description': 설명}}
        """
        try:
            # 엑셀 파일 읽기
            df = pd.read_excel(file_path)
            
            # 결과 저장을 위한 딕셔너리
            result = {}
            
            # 각 열에 대한 정보 추출
            total_columns = len(df.columns)
            for i, col in enumerate(df.columns):
                # 데이터 타입 추론
                col_type = str(df[col].dtype)
                
                # 설명 - 샘플 데이터 등을 포함
                non_na_values = df[col].dropna()
                sample_data = "데이터 없음"
                if len(non_na_values) > 0:
                    sample_values = non_na_values.head(3).tolist()
                    sample_data = ", ".join(str(val) for val in sample_values)
                
                # 결과 저장
                result[col] = {
                    'data_type': col_type,
                    'description': f"샘플 데이터: {sample_data}"
                }
                
                # 진행률 업데이트
                if progress_callback:
                    progress = int((i + 1) / total_columns * 100)
                    progress_callback(progress)
            
            return result
            
        except Exception as e:
            logging.error(f"엑셀 파싱 오류: {str(e)}")
            raise
    
    def save_as_csv(self, data, file_path):
        """
        결과를 CSV 파일로 저장
        
        Args:
            data (dict): 저장할 데이터
            file_path (str): 저장할 파일 경로
        """
        try:
            with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                # 헤더 작성
                writer.writerow(['헤더명', '데이터타입', '설명'])
                
                # 데이터 작성
                for header, info in data.items():
                    writer.writerow([
                        header,
                        info.get('data_type', ''),
                        info.get('description', '')
                    ])
                    
            return True
        except Exception as e:
            logging.error(f"CSV 저장 오류: {str(e)}")
            raise
    
    def save_as_json(self, data, file_path):
        """
        결과를 JSON 파일로 저장
        
        Args:
            data (dict): 저장할 데이터
            file_path (str): 저장할 파일 경로
        """
        try:
            with open(file_path, 'w', encoding='utf-8') as jsonfile:
                json.dump(data, jsonfile, ensure_ascii=False, indent=2)
                
            return True
        except Exception as e:
            logging.error(f"JSON 저장 오류: {str(e)}")
            raise


