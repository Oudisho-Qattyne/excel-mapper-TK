from tkinterdnd2 import TkinterDnD
from gui import ExcelMapperGUI


def main():
    root = TkinterDnD.Tk()
    app = ExcelMapperGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
