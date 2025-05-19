import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import logging

from src.services.excel_parse_service import ExcelParseService

# 로깅 설정
logging.basicConfig(level=logging.INFO, 
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class UiMain:
    def __init__(self, root):
        self.root = root
        self.root.title("엑셀 헤더 파서")
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        
        # ExcelParseService 인스턴스 생성
        self.excel_service = ExcelParseService()
        
        # UI 요소 생성
        self.create_widgets()
        
    def create_widgets(self):
        # 상단 프레임 - 파일 선택 영역
        top_frame = ttk.LabelFrame(self.root, text="파일 선택")
        top_frame.pack(fill="x", expand=False, padx=10, pady=5)
        
        # 파일 선택 버튼과 경로 표시
        ttk.Label(top_frame, text="엑셀 파일:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.file_path_var = tk.StringVar()
        ttk.Entry(top_frame, textvariable=self.file_path_var, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(top_frame, text="찾아보기", command=self.browse_file).grid(row=0, column=2, padx=5, pady=5)
        
        # 출력 폴더 선택 버튼과 경로 표시
        ttk.Label(top_frame, text="출력 폴더:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.output_path_var = tk.StringVar()
        ttk.Entry(top_frame, textvariable=self.output_path_var, width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(top_frame, text="찾아보기", command=self.browse_output_folder).grid(row=1, column=2, padx=5, pady=5)
        
        # 열 번호 입력 영역
        ttk.Label(top_frame, text="데이터 열 번호:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        self.column_num_var = tk.StringVar(value="1")
        ttk.Spinbox(top_frame, from_=1, to=100, textvariable=self.column_num_var, width=5).grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 처리 시작 버튼
        self.process_button = ttk.Button(top_frame, text="처리 시작", command=self.start_processing)
        self.process_button.grid(row=3, column=0, columnspan=3, pady=10)
        
        # 진행 상태 표시
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(top_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=4, column=0, columnspan=3, padx=5, pady=5, sticky="ew")
        
        # 하단 프레임 - 로그 출력 영역
        log_frame = ttk.LabelFrame(self.root, text="처리 로그")
        log_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        # 로그 출력 텍스트 영역
        self.log_text = tk.Text(log_frame, wrap=tk.WORD, height=10)
        self.log_text.pack(fill="both", expand=True, padx=5, pady=5)
        
        # 스크롤바 추가
        scrollbar = ttk.Scrollbar(self.log_text, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        # 로그 핸들러 설정
        self.log_handler = TextHandler(self.log_text)
        self.log_handler.setLevel(logging.INFO)
        logger.addHandler(self.log_handler)
        logging.getLogger('src.services.excel_parse_service').addHandler(self.log_handler)
    
    def browse_file(self):
        """파일 선택 대화상자 열기"""
        filename = filedialog.askopenfilename(
            title="엑셀 파일 선택",
            filetypes=[("Excel 파일", "*.xlsx *.xls"), ("모든 파일", "*.*")]
        )
        if filename:
            self.file_path_var.set(filename)
    
    def browse_output_folder(self):
        """출력 폴더 선택 대화상자 열기"""
        folder = filedialog.askdirectory(title="출력 폴더 선택")
        if folder:
            self.output_path_var.set(folder)
    
    def start_processing(self):
        """엑셀 파일 처리 시작"""
        file_path = self.file_path_var.get()
        output_path = self.output_path_var.get()
        
        try:
            column_num = int(self.column_num_var.get())
        except ValueError:
            messagebox.showerror("입력 오류", "열 번호는 숫자여야 합니다.")
            return
        
        if not file_path:
            messagebox.showerror("입력 오류", "엑셀 파일을 선택해주세요.")
            return
        
        if not os.path.exists(file_path):
            messagebox.showerror("파일 오류", "선택한 엑셀 파일이 존재하지 않습니다.")
            return
        
        if not output_path:
            # 출력 폴더가 지정되지 않았을 경우 기본값 사용
            output_path = os.path.join(os.path.dirname(file_path), "output")
            self.output_path_var.set(output_path)
        
        # 처리 버튼 비활성화
        self.process_button.config(state=tk.DISABLED)
        
        # 로그 초기화
        self.log_text.delete(1.0, tk.END)
        
        # 출력 폴더 설정
        self.excel_service.output_folder = output_path
        
        # 부모 폴더 경로와 파일명 분리
        file_dir, file_name = os.path.split(file_path)
        self.excel_service.input_folder = file_dir
        
        # 백그라운드 스레드에서 처리 시작
        thread = threading.Thread(target=self.process_file, args=(file_name, column_num))
        thread.daemon = True
        thread.start()
    
    def process_file(self, file_name, column_num):
        """파일 처리 실행 (백그라운드 스레드)"""
        try:
            logger.info(f"파일 처리 시작: {file_name}, 열 번호: {column_num}")
            
            # 진행 상황 표시
            self.progress_var.set(10)
            
            # 파일 처리 실행
            output_files = self.excel_service.parse_data(file_name, column_num)
            
            # 진행 상황 완료
            self.progress_var.set(100)
            
            # 처리 완료 메시지
            if output_files:
                logger.info(f"처리 완료: {len(output_files)}개의 파일이 생성되었습니다.")
                
                # GUI 스레드에서 메시지 박스 표시
                self.root.after(0, lambda: messagebox.showinfo(
                    "처리 완료", 
                    f"{len(output_files)}개의 파일이 생성되었습니다.\n저장 위치: {self.excel_service.output_folder}"
                ))
            
        except Exception as e:
            logger.error(f"처리 중 오류 발생: {str(e)}")
            # GUI 스레드에서 오류 메시지 표시
            self.root.after(0, lambda: messagebox.showerror("오류", f"처리 중 오류가 발생했습니다.\n{str(e)}"))
        finally:
            # GUI 스레드에서 버튼 활성화
            self.root.after(0, lambda: self.process_button.config(state=tk.NORMAL))


class TextHandler(logging.Handler):
    """로그 메시지를 텍스트 위젯으로 리디렉션하는 핸들러"""
    
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget
        
    def emit(self, record):
        msg = self.format(record)
        
        def append():
            self.text_widget.insert(tk.END, msg + '\n')
            self.text_widget.see(tk.END)  # 항상 마지막 라인이 보이도록
        
        # GUI 스레드에서 실행
        self.text_widget.after(0, append)