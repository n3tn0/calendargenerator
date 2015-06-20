__author__ = 'n3tn0'
__version__ = 'BETA'

from PyQt5.QtWidgets import *
from main_auto import Ui_Sumy
import sys
from schedulecreate import *

class ImageDialog(QDialog, Ui_Sumy):
    def __init__(self):
        super(ImageDialog, self).__init__()
        self.setupUi(self)

        # Connect the buttons.
        self.browse.clicked.connect(self.selectfile)
        self.submit.clicked.connect(self.parse)
        self.info.clicked.connect(self.infor)
        self.quit.clicked.connect(QApplication.exit)

    def selectfile(self):
        spreadsheet = QFileDialog.getOpenFileName(self, "Select Spreadsheet")
        self.filepath.setText(spreadsheet[0])

    def parse(self):
        import os
        if self.all.isChecked():
            period = 'all'
        if self.nextmonth.isChecked():
            period = 'next'
        if self.current.isChecked():
            period = 'current'

        global workbook, calendarfile
        workbook = self.filepath.displayText()
        calendarfile = os.path.split(workbook)[0]+'\calendar.xlsx'
        schedulemaker = CreateSchedule()
        schedulemaker.build(workbook, calendarfile)

        QMessageBox.information(self, 'Success!', 'Calendar created and saved!')
        QApplication.exit()

    def infor(self):
        QMessageBox.information(self, 'Sumy v' + __version__, '<p align=center> CALENDAR GENERATOR <br> Made by Timothy Noto <br> Version: %s <br> Released Under MIT Licence' % (__version__))

def main():
    app = QApplication(sys.argv)
    form = ImageDialog()
    form.show()
    app.exec_()

main()