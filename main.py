import sys
from PyQt5.QtWidgets import QApplication
from main_window import Main


def main():
    app = QApplication(sys.argv)
    mn = Main()
    mn.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
