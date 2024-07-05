import sys
import logging
from PyQt5.QtWidgets import QApplication
from gui import MainWindow

logging.basicConfig(filename='app.log', level=logging.ERROR)

def main():

    try:
        app = QApplication(sys.argv)
        main_window = MainWindow()
        main_window.show()
        sys.exit(app.exec_())
    except Exception as e:
        logging.error("오류가 발생했습니다: %s", e, exc_info=True)


if __name__ == "__main__":
    main()