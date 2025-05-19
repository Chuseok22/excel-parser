import os
import subprocess
import sys

# Window CP1252 환경 stdout 인코딩 설정
try:
    sys.stdout.reconfigure(encoding="utf-8")
except AttributeError:
    pass

def main():
    print("Excel Parser EXE 빌드 시작")

    # PyInstaller 설치 확인
    try:
        import PyInstaller
        print(f"PyInstaller 버전: {PyInstaller.__version__}")
    except ImportError:
        print("PyInstaller 설치 중...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "PyInstaller"])

    # 빌드 명령 실행
    build_command = [
        "pyinstaller",
        "--onefile", # 단일 exe 파일로 빌드
        "--windowed", # 콘솔 창 없이 실행
        "--name", "excel-parser", # 출력 파일 이름
        "main.py" # 메인 스크립트 파일
    ]
    print(f"빌드 명령 실행: {' '.join(build_command)}")
    subprocess.check_call(build_command)
    
    # 빌드 결과 확인
    dist_path = os.path.join(os.getcwd(), "dist", "excel-parser.exe")
    if os.path.exists(dist_path):
        print(f"빌드 성공: {dist_path}")
        # dist 디렉토리 내용 표시
        print("생성된 파일 목록:")
        for file in os.listdir("dist"):
            print(f" - {file}")
        return 0
    else:
        print("빌드 실패")
        return 1
    
if __name__ == "__main__":
    sys.exit(main())