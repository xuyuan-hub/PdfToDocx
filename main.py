# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

import sys
import logging
from PyQt5.QtCore import pyqtSignal,QThread
from PyQt5.QtWidgets import QApplication, QMainWindow,QFileDialog
from pdftoword import pdf_to_word
from GUI.MainWindow import Ui_officeConverter

class ConvertThread(QThread):
    convert_finished = pyqtSignal()  # 自定义信号，用于通知主线程转换完成

    def __init__(self, input_file,docx_file):
        super(ConvertThread, self).__init__()
        self.input_file = input_file
        self.docx_file = docx_file

    def run(self):
        pdf_to_word(self.input_file, self.docx_file)
        self.convert_finished.emit()  # 发送转换完成信号

class QTextEditHandler(logging.Handler):
    def __init__(self, text_edit):
        super(QTextEditHandler, self).__init__()
        self.text_edit = text_edit
        self.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

    def emit(self, record):
        msg = self.format(record)
        print(msg)
        self.text_edit.append(msg)


class myMainWindow(Ui_officeConverter,QMainWindow):
    _signal = pyqtSignal

    def __init__(self):

        super(Ui_officeConverter,self).__init__()
        self.setupUi(self)
        self.slot_init()
        self.init_logging()

    def slot_init(self):
        self.openDirectory.clicked.connect(self.openDirectoryEvent)
        self.launch.clicked.connect(self.start_convert)

    def openDirectoryEvent(self):
        """
        Open directory
        :return: file path
        """
        self.input_file, _ = QFileDialog.getOpenFileName(self, "选择输入文件")
        self.inputPath.setText(self.input_file)


    def start_convert(self):
        """
        Start the conversion in a separate thread
        :return:
        """
        if self.comboBox.currentIndex() == 0:
            docx_file = ".".join(self.input_file.split(".")[0:-1]) +".docx"
            self.convert_thread = ConvertThread(self.input_file,docx_file)
            self.convert_thread.convert_finished.connect(self.on_convert_finished)
            self.convert_thread.start()

    def on_convert_finished(self):
        """
        Slot for handling conversion finished signal
        :return:
        """
        self.progressBar.setProperty("value", 100)
        pass

    def init_logging(self):
        logger = logging.getLogger()
        logger.setLevel(logging.INFO)

        # 添加 QTextEditHandler
        text_edit_handler = QTextEditHandler(self.textBrowser)
        logger.addHandler(text_edit_handler)

    def log_info(self):
        logging.info("This is an example log message.")

if __name__ == "__main__":
    # main()
    app = QApplication(sys.argv)
    count_gui = myMainWindow()
    count_gui.show()
    sys.exit(app.exec_())

