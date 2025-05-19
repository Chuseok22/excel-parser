#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Excel 파서 UI 모듈
PyQt5를 사용하여 사용자 인터페이스를 제공합니다.
"""

import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QFileDialog, 
                           QLabel, QVBoxLayout, QHBoxLayout, QWidget, QTableWidget, 
                           QTableWidgetItem, QMessageBox, QProgressBar)
from PyQt5.QtCore import Qt, QThread, pyqtSignal

# 상위 디렉토리의 모듈을 import 하기 위한 경로 추가
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Excel 파싱 서비스 가져오기 (구현이 필요함)
from services.excel_parse_service import ExcelParseService


class WorkerThread(QThread):
    """
    백그라운드에서 엑셀 파싱을 수행하는 작업 스레드
    """
    progress_signal = pyqtSignal(int)
    finished_signal = pyqtSignal(dict)
    error_signal = pyqtSignal(str)
    
    def __init__(self, file_path):
        super().__init__()
        self.file_path = file_path
        self.excel_service = ExcelParseService()
        
    def run(self):
        try:
            # 엑셀 파일 파싱 작업 수행
            result = self.excel_service.parse_excel(self.file_path, progress_callback=self.update_progress)
            self.finished_signal.emit(result)
        except Exception as e:
            self.error_signal.emit(str(e))
    
    def update_progress(self, value):
        self.progress_signal.emit(value)


class ExcelParserUI(QMainWindow):
    """
    Excel 파서 메인 UI 클래스
    """
    def __init__(self):
        super().__init__()
        self.init_ui()
        
    def init_ui(self):
        """UI 초기화 메서드"""
        self.setWindowTitle('Excel Header Parser')
        self.setGeometry(100, 100, 800, 600)
        
        # 메인 위젯과 레이아웃
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        
        # 파일 선택 영역
        file_layout = QHBoxLayout()
        self.file_label = QLabel("파일을 선택해주세요")
        self.browse_button = QPushButton("파일 찾기")
        self.browse_button.clicked.connect(self.browse_file)
        file_layout.addWidget(self.file_label)
        file_layout.addWidget(self.browse_button)
        main_layout.addLayout(file_layout)
        
        # 진행 상황 표시
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)
        
        # 결과 테이블
        self.result_table = QTableWidget()
        self.result_table.setColumnCount(3)
        self.result_table.setHorizontalHeaderLabels(['헤더', '데이터 타입', '설명'])
        self.result_table.horizontalHeader().setStretchLastSection(True)
        main_layout.addWidget(self.result_table)
        
        # 실행 및 저장 버튼
        button_layout = QHBoxLayout()
        self.parse_button = QPushButton("헤더 파싱")
        self.parse_button.setEnabled(False)
        self.parse_button.clicked.connect(self.parse_excel)
        
        self.save_button = QPushButton("결과 저장")
        self.save_button.setEnabled(False)
        self.save_button.clicked.connect(self.save_results)
        
        button_layout.addWidget(self.parse_button)
        button_layout.addWidget(self.save_button)
        main_layout.addLayout(button_layout)
        
        # 상태바 추가
        self.statusBar().showMessage('준비됨')
        
        # 클래스 변수 초기화
        self.selected_file = None
        self.parse_results = None
        self.worker_thread = None
    
    def browse_file(self):
        """파일 탐색기 열기"""
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName(
            self, '엑셀 파일 선택', '', 'Excel Files (*.xlsx *.xls)')
        
        if file_path:
            self.selected_file = file_path
            self.file_label.setText(f"선택된 파일: {os.path.basename(file_path)}")
            self.parse_button.setEnabled(True)
            self.statusBar().showMessage(f"파일이 선택됨: {file_path}")
    
    def parse_excel(self):
        """엑셀 파일 파싱 시작"""
        if not self.selected_file:
            QMessageBox.warning(self, '경고', '파일을 먼저 선택하세요.')
            return
            
        # UI 상태 업데이트
        self.parse_button.setEnabled(False)
        self.browse_button.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.statusBar().showMessage('파싱 중...')
        
        # 워커 스레드 시작
        self.worker_thread = WorkerThread(self.selected_file)
        self.worker_thread.progress_signal.connect(self.update_progress)
        self.worker_thread.finished_signal.connect(self.parsing_finished)
        self.worker_thread.error_signal.connect(self.parsing_error)
        self.worker_thread.start()
    
    def update_progress(self, value):
        """진행 상황 업데이트"""
        self.progress_bar.setValue(value)
    
    def parsing_finished(self, results):
        """파싱 작업 완료 처리"""
        self.parse_results = results
        self.display_results()
        
        # UI 상태 업데이트
        self.browse_button.setEnabled(True)
        self.parse_button.setEnabled(True)
        self.save_button.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.statusBar().showMessage('파싱 완료')
    
    def parsing_error(self, error_message):
        """파싱 오류 처리"""
        QMessageBox.critical(self, '오류', f'파싱 중 오류가 발생했습니다: {error_message}')
        
        # UI 상태 업데이트
        self.browse_button.setEnabled(True)
        self.parse_button.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.statusBar().showMessage('파싱 오류')
    
    def display_results(self):
        """결과를 테이블에 표시"""
        if not self.parse_results:
            return
            
        # 테이블 초기화
        self.result_table.setRowCount(0)
        
        # 결과 데이터 추가
        for i, (header, info) in enumerate(self.parse_results.items()):
            self.result_table.insertRow(i)
            self.result_table.setItem(i, 0, QTableWidgetItem(header))
            self.result_table.setItem(i, 1, QTableWidgetItem(info.get('data_type', '')))
            self.result_table.setItem(i, 2, QTableWidgetItem(info.get('description', '')))
    
    def save_results(self):
        """결과를 파일로 저장"""
        if not self.parse_results:
            QMessageBox.warning(self, '경고', '저장할 결과가 없습니다.')
            return
            
        file_dialog = QFileDialog()
        save_path, _ = file_dialog.getSaveFileName(
            self, '결과 저장', '', 'CSV Files (*.csv);;JSON Files (*.json)')
            
        if not save_path:
            return
            
        try:
            # 파일 형식에 따라 저장 로직 구현
            file_ext = os.path.splitext(save_path)[1].lower()
            
            if file_ext == '.csv':
                self.excel_service.save_as_csv(self.parse_results, save_path)
            elif file_ext == '.json':
                self.excel_service.save_as_json(self.parse_results, save_path)
            else:
                QMessageBox.warning(self, '경고', '지원되지 않는 파일 형식입니다.')
                return
                
            QMessageBox.information(self, '저장 완료', f'결과가 {save_path}에 저장되었습니다.')
            self.statusBar().showMessage(f'결과 저장됨: {save_path}')
        except Exception as e:
            QMessageBox.critical(self, '오류', f'저장 중 오류가 발생했습니다: {str(e)}')


def run_ui():
    """UI 실행 함수"""
    app = QApplication(sys.argv)
    window = ExcelParserUI()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    run_ui()
