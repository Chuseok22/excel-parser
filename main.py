import sys

import tkinter as tk
from ui.ui_main import UiMain

def main():
    """
    애플리케이션 메인 함수
    """
    root = tk.Tk()
    app = UiMain(root)
    root.mainloop()
    return 0
    
if __name__ == "__main__":
    sys.exit(main())