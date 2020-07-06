# -*- coding: utf-8 -*-
"""
Created on Wed Mar 11 11:28:16 2020

@author: Andrew-V.Young

Adding comment for new pull

Adding another comment for newer pull

Added extra #2 comment for pull request
"""


import sys
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QPushButton, QLabel, QTextEdit, QFileDialog, QApplication
import pandas as pd
from openpyxl import load_workbook


def reformat_drift_table(inFileName = 'Book1.xlsx'):
    # function to reformat story drift spreadsheet created by etabs
    ofile = pd.read_excel(inFileName, header = 1)
    
    # remove first and third row which are etabs titles
    delete_rows = [0, 2]
    rmvTitleRows = ofile.drop(delete_rows) 
    
    # insert initial excel row number
    excelRow = rmvTitleRows.apply(lambda row: row.name + len(delete_rows) + 1, axis = 1)
    rmvTitleRows.insert(0,'Initial Row', excelRow, True)
    
    # filter for only rows that contain drift combos
    driftRows = rmvTitleRows.loc[rmvTitleRows['Load Case/Combo'].str.contains('drift', case = False)]
        
    # calculate story drift DCR
    maxDrift = 0.01
    dcrSeries = driftRows.apply(lambda row: row.Drift/maxDrift, axis = 1)
    
    # insert DCR column and sort largest to smallest
    driftRows.insert(len(driftRows.columns), 'DCR', dcrSeries, True)
    dfSort = driftRows.sort_values(by=['DCR'], ascending = False)
    # print(dfSort.head())
    
    book = load_workbook(inFileName)  # new data entry without deleting existing
    
    # add sorted data to new sheet
    with pd.ExcelWriter(inFileName, engine = 'openpyxl') as writer:
        writer.book = book
        dfSort.to_excel(writer, sheet_name = 'Drift Sorted')
        writer.save()
        writer.close()

    return 'reformatting complete'


class get_file_dialog(QWidget):
    def __init__(self, parent=None):
        super(get_file_dialog, self).__init__(parent)
        self.initUI()

    def initUI(self):

        # add window title and prompt label text
        self.setWindowTitle("Story Drift Reformatting Tool")
        layout = QVBoxLayout()
        self.le = QLabel("Select Story Drift File To Reformat")
        layout.addWidget(self.le)

        # add button to open file name, connect to open file function
        self.btn = QPushButton("Choose File")
        self.btn.clicked.connect(lambda: self.getfile())
        layout.addWidget(self.btn)

        # add text box to use as status notification, enter initial text
        self.statustext = QTextEdit()
        self.statustext.setText('Please use button above to choose a file')
        layout.addWidget(self.statustext)

        self.setLayout(layout)

    def getfile(self):
        # function that pulls up get open file name window and re-formats selected file
        dlg = QFileDialog()
        dlg.setFileMode(QFileDialog.AnyFile)
        dlg.setNameFilter("Excel files (*.xlsx)")

        # open window and extract file name from outputs
        fileName, others = dlg.getOpenFileName(self, "Choose File")

        # run reformatting if file chosen, otherwise no action
        if fileName:
            mess1 = "Selected File: \n %s \n\n" % fileName
            file_format = reformat_drift_table(fileName)
            self.statustext.setText(mess1 + file_format)
        else:
            not_opened = "No file was opened"
            self.statustext.setText(not_opened)

# set up main application with get_file_dialog widget
def main():
   app = QApplication(sys.argv)
   ex = get_file_dialog()
   ex.show()
   sys.exit(app.exec_())

# run widget
if __name__ == '__main__':
    main()
