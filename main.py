import tkinter as tk
from app_window import TaskAnalyzerGUI


def main():
    root = tk.Tk()
    TaskAnalyzerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()