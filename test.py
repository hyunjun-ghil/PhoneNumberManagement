import pymssql
import sys
import os
import easygui
import time
import pandas as pd
import paramiko
from PyQt5.QtWidgets import *
from PyQt5 import uic
from PyQt5.QtGui import *
from functools import partial


#UI파일 연결
#단, UI파일은 Python 코드 파일과 같은 디렉토리에 위치해야한다.

global g_FLAG, LoginId
g_FLAG = ""
LoginId = ""

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

recipupload_class = uic.loadUiType(resource_path("RECIPUpload.ui"))[0]

class RECIPUploadDialog(QDialog, recipupload_class):
    def __init__(self) :
        super().__init__()
        self.setupUi(self)
        self.FileUploadBtn.clicked.connect(self.fileUpload)
        self.saveBtn.clicked.connect(self.saveToDB)

    def fileUpload(self):
        path = easygui.fileopenbox()
        file = open(path, "r");
        row_index = 0

        conn = pymssql.connect(server='192.168.35.155', user='cms', password='cms', database='CMS',
                               charset='cp949')
        cursor = conn.cursor(as_dict=True)

        sql = "delete from TRCCallCenterRecordInfo_KCPM"

        try:
            cursor.execute(sql)
            conn.commit()
        except pymssql.DatabaseError as e:
            QMessageBox.warning(self, "Error", e)

        for data in file:
            str = data.split()
            print(str[0] + ' ' + str[1] + ' ' + str[2] + ' ' + str[5])

            self.ip_table.setItem(row_index, 0, QTableWidgetItem(str[0]))
            self.ip_table.setItem(row_index, 1, QTableWidgetItem(str[1]))
            self.ip_table.setItem(row_index, 2, QTableWidgetItem(str[2]))
            self.ip_table.setItem(row_index, 3, QTableWidgetItem(str[5]))
            row_count = self.ip_table.rowCount()
            self.ip_table.setRowCount(row_count + 1)  # row 추가
            row_index = row_index + 1
        conn.close()

    def saveToDB(self):

        answer = QMessageBox.question(self, "notice", "저장하시겠습니까?")
        if answer == QMessageBox.Yes:
            conn = pymssql.connect(server='192.168.35.155', user='cms', password='cms', database='CMS',
                                   charset='cp949')
            cursor = conn.cursor(as_dict=True)

            try:
                row_count = self.ip_table.rowCount() - 1
                for i in range(0, row_count):
                    t_seq = self.ip_table.item(i, 0).text()
                    t_dn = self.ip_table.item(i, 1).text()
                    t_record_server = self.ip_table.item(i, 2).text()
                    t_phone_ip = self.ip_table.item(i, 3).text()
                    sql = "insert into TRCCallCenterRecordInfo_KCPM(SEQ, DN, Record_Server, Phone_IP) values('{SEQ}', '{DN}', '{Record_Server}', '{Phone_IP}')".format(
                        SEQ=t_seq, DN=t_dn, Record_Server=t_record_server, Phone_IP=t_phone_ip)
                    cursor.execute(sql)
                    conn.commit()

            except pymssql.DatabaseError as e:
                QMessageBox.warning(self, "Error", e)

            QMessageBox.information(self, "notice", "저장되었습니다.")
            conn.close()
        else:
            QMessageBox.information(self, "notice", "취소되었습니다.")


if __name__ == "__main__" :
    app = QApplication(sys.argv)
    myWindow = RECIPUploadDialog()
    myWindow.show()
    app.exec_()

