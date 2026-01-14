import logging

try:
    from src.ui.main_window import MainWindow
except ModuleNotFoundError:
    import os
    import sys

    sys.path.append(os.path.dirname(os.path.dirname(__file__)))
    from src.ui.main_window import MainWindow

def main():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )
    w = MainWindow()
    w.run()

if __name__ == "__main__":
    main()
