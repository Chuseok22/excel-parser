import sys
import os

# 상위 디렉토리의 모듈을 import 하기 위한 경로 추가
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# UI 모듈 가져오기
from src.ui.excel_parser_ui import run_ui


def main():
    """
    애플리케이션 메인 함수
    Excel Parser UI를 실행합니다.
    """
    run_ui()


if __name__ == "__main__":
    main()
