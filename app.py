import tkinter as tk
from gui import ExcelMapperGUI


def main():
    root = tk.Tk()
    app = ExcelMapperGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
