import csv

from PyQt5 import QtCore, QtGui, QtWidgets
import pandas as pd
import pyodbc
import urllib
from sqlalchemy import create_engine
import ntpath
import sys
#sys.stdout = open('CSV_Reader_Log.txt', 'w')



class MyWindow(QtWidgets.QWidget):
    def __init__(self, parent=None):
        super(MyWindow, self).__init__(parent)
        self.setFixedWidth= 800
        self.setFixedHeight=600

        #self.model = QtGui.QStandardItemModel(self)
        self.model = QtGui.QStandardItemModel(self)

        self.tableView = QtWidgets.QTableView(self)
        self.tableView.setModel(self.model)
        self.tableView.horizontalHeader().setStretchLastSection(True)

        self.pushButtonLoad = QtWidgets.QPushButton(self)
        self.pushButtonLoad.setText("Load Csv File!")
        self.pushButtonLoad.clicked.connect(self.on_pushButtonLoad_clicked)

        self.pushButtonWrite = QtWidgets.QPushButton(self)
        self.pushButtonWrite.setText("Write SQL Server!")
        self.pushButtonWrite.clicked.connect(self.on_pushButtonWrite_clicked)

        self.layoutVertical = QtWidgets.QVBoxLayout(self)
        self.layoutVertical.addWidget(self.tableView)
        self.layoutVertical.addWidget(self.pushButtonLoad)
        self.layoutVertical.addWidget(self.pushButtonWrite)

    def loadCsv(self, fileName):
        self.tableView.reset()
        self.model = QtGui.QStandardItemModel(self)
        self.tableView.setModel(self.model)
        count = 0
        with open(fileName, "r") as fileInput:
            for row in csv.reader(fileInput):    
                items = [
                    QtGui.QStandardItem(field)
                    for field in row
                ]
                self.model.appendRow(items)
                count += 1
        
        msg = 'Csv data read ' + str(count) + ' records'
        QtWidgets.QMessageBox.information(self, 'Information', msg) 

        #self.saveToSQLServer(self.fileName, tName)


    def saveToSQLServer(self, fileName, tableName):
        try:
            data = pd.read_csv (fileName) 
            df = pd.DataFrame(data)

            con = pyodbc.connect('Driver={SQL Server};'
                                  'Server=DESKTOP-CME8ABH\DPS_SQL_SERVER;'
                                  'Database=TestDb;'
                                  'Trusted_Connection=yes;')

            server = 'DESKTOP-CME8ABH\DPS_SQL_SERVER'
            database = 'TestDb'
            username = 'PafWriter'
            password = 'P1unpaussC_1105='

            params= urllib.parse.quote_plus('DRIVER={SQL Server};' + 'SERVER='+ server +';DATABASE='+ database +';UID='+username+';PWD='+ password)
            engine = create_engine("mssql+pyodbc:///?odbc_connect=%s"%params)

            try:
                df.to_sql(tableName, con = engine, if_exists='append', index=False)  
                #print('Data saved: ', str(len(data)))
                msg = 'SQL data written ' + str(len(df)) + ' records'
                QtWidgets.QMessageBox.information(self, 'Information', msg) 


            except Exception as ex:
                print('Data failed')
                print(ex)
        except Exception as ex:
            print('Error reading file')
            print(ex)


    def writeCsv(self, fileName):
        with open(fileName, "w") as fileOutput:
            writer = csv.writer(fileOutput)
            for rowNumber in range(self.model.rowCount()):
                fields = [
                    self.model.data(
                        self.model.index(rowNumber, columnNumber),
                        QtCore.Qt.DisplayRole
                    )
                    for columnNumber in range(self.model.columnCount())
                ]
                writer.writerow(fields)

    @QtCore.pyqtSlot()
    def on_pushButtonWrite_clicked(self):
        head, tail = ntpath.split(self.fileName)
        fName = tail or ntpath.basename(head)
        tName = fName.split('.')[0]

        self.saveToSQLServer(self.fileName, tName)



        #self.writeCsv(self.fileName)

    @QtCore.pyqtSlot()
    def on_pushButtonLoad_clicked(self):
        result = QtWidgets.QFileDialog.getOpenFileName(self, 'Open csv file', '.', 'Csv files (*.csv)')
        self.fileName=result[0]
        self.loadCsv(self.fileName)

if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    app.setApplicationName('CSV Reader')


    main = MyWindow()
    main.resize(800,600)
    main.show()

    sys.exit(app.exec_())
