# Excel Parser

Excel 파일 파싱 및 처리를 위한 Python 어플리케이션입니다.

## 기능

- Excel 파일 읽기 및 쓰기
- 데이터 정제 및 변환
- 파일 병합
- 결과 저장 및 내보내기

## 요구사항

- Python 3.8 이상
- pandas
- openpyxl
- 기타 requirements.txt에 명시된 패키지들

## 설치 방법

### 1. 가상환경 설정

```bash
# 가상환경 생성
python -m venv venv

# 가상환경 활성화 (Windows)
venv\Scripts\activate

# 가상환경 활성화 (macOS/Linux)
source venv/bin/activate
```

### 2. 의존성 패키지 설치

```bash
pip install -r requirements.txt
```

## 사용 방법

### 기본 실행

```bash
python app.py
```

### 테스트 실행

```bash
python -m pytest
```

### EXE 파일 빌드 (Windows)

```bash
pyinstaller excel_parser.spec
```

## 프로젝트 구조

```
excel-parser/
├── app.py                  # 애플리케이션 진입점
├── config.json             # 설정 파일
├── requirements.txt        # 의존성 패키지 목록
├── excel_parser.spec       # PyInstaller 스펙 파일
├── src/                    # 소스 코드
│   ├── configs/            # 설정 관련 모듈
│   ├── models/             # 데이터 모델
│   ├── services/           # 비즈니스 로직
│   └── utils/              # 유틸리티 함수
├── tests/                  # 테스트 코드
├── input/                  # 입력 파일 디렉토리
└── output/                 # 출력 파일 디렉토리
```

## 라이센스

이 프로젝트는 MIT 라이센스를 따릅니다.
