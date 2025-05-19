# Excel 헤더 파서

Excel 파일의 특정 열을 기준으로 데이터를 분류하고 별도의 엑셀 파일로 추출하는 Python 애플리케이션입니다.

## 주요 기능

- 엑셀 파일의 특정 열(사용자 지정)을 기준으로 데이터 그룹화
- 그룹화된 데이터를 각각 별도의 엑셀 파일로 추출
- 직관적인 GUI 인터페이스 제공
- 처리 과정 실시간 모니터링

## 요구사항

- Python 3.10 이상 (GitHub Actions 빌드용)
- pandas: 엑셀 파일 처리를 위한 데이터 분석 라이브러리
- openpyxl: Excel 파일 읽기/쓰기 라이브러리
- tkinter: GUI 인터페이스 (Python 표준 라이브러리)
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
python main.py
```

GUI 애플리케이션이 실행되며, 다음 단계에 따라 사용할 수 있습니다:

1. **엑셀 파일 선택**: "찾아보기" 버튼을 클릭하여 처리할 엑셀 파일을 선택합니다.
2. **출력 폴더 설정**: 결과 파일을 저장할 폴더를 선택합니다. 미지정 시 입력 파일 위치에 output 폴더가 자동 생성됩니다.
3. **데이터 열 번호 입력**: 그룹화 기준이 될 열 번호를 입력합니다 (1부터 시작).
   - 예: 부서명이 1번 열에 있다면 "1"을 입력
4. **처리 시작**: "처리 시작" 버튼을 클릭하여 파일 처리를 시작합니다.
5. **처리 결과 확인**: 처리 로그 영역에서 진행 상황 및 결과를 확인할 수 있습니다.

### 예시

다음과 같은 엑셀 데이터가 있을 때:

| 부서명 | 직원명 | 급여 |
|--------|--------|------|
| 부서1  | 직원A  | 30000|
| 부서1  | 직원B  | 50000|
| 부서2  | 직원C  | 20000|
| 부서2  | 직원D  | 10000|

데이터 열 번호에 "1"을 입력하면(부서명 기준):
- "부서1" 데이터만 담긴 엑셀 파일 1개
- "부서2" 데이터만 담긴 엑셀 파일 1개

총 2개의 파일이 생성됩니다.

### EXE 파일 실행 (Windows)

배포된 EXE 파일을 더블 클릭하여 직접 실행할 수 있습니다.

### EXE 파일 빌드 (개발자용)

```bash
# PyInstaller 설치
pip install pyinstaller

# EXE 파일 빌드
pyinstaller --onefile --windowed --icon=app_icon.ico main.py
```

## 프로젝트 구조

```
excel-header-parser/
├── main.py                 # 애플리케이션 진입점
├── requirements.txt        # 의존성 패키지 목록
├── README.md               # 프로젝트 설명 문서
├── src/                    # 소스 코드
│   ├── models/             # 데이터 모델
│   ├── services/           # 비즈니스 로직
│   │   └── excel_parse_service.py  # 엑셀 파싱 서비스
│   ├── ui/                 # UI 관련 모듈
│   │   └── ui_main.py      # 메인 UI 클래스
│   └── utils/              # 유틸리티 함수
├── input/                  # 입력 파일 디렉토리 (자동 생성)
└── output/                 # 출력 파일 디렉토리 (자동 생성)
```

## 주요 모듈 설명

### ExcelParseService

`src/services/excel_parse_service.py` - 엑셀 파일 처리 핵심 로직을 담당하는 클래스
- `read_excel()`: 엑셀 파일을 읽어 DataFrame으로 변환
- `parse_data()`: 지정된 열을 기준으로 데이터를 그룹화하고 별도의 파일로 저장

### UiMain

`src/ui/ui_main.py` - 사용자 인터페이스를 담당하는 클래스
- 파일 선택 및 출력 폴더 지정 기능
- 열 번호 입력 필드
- 처리 진행 상태 표시
- 로그 메시지 실시간 출력

## 라이센스

이 프로젝트는 MIT 라이센스를 따릅니다.
