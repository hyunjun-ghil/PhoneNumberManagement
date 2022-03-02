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
import sql
import cube

#UI파일 연결
#단, UI파일은 Python 코드 파일과 같은 디렉토리에 위치해야한다.

global g_FLAG, LoginId
g_FLAG = ""
LoginId = ""

# DB 접속정보
cmslogin = sql.CMSlnfo
cubelogin_1 = cube.CubeInfo_1
cubelogin_2 = cube.CubeInfo_2


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


form_class = uic.loadUiType(resource_path("main.ui"))[0]
showModifyWindow_class = uic.loadUiType(resource_path("modifyWindow.ui"))[0]
searchDN_class = uic.loadUiType(resource_path("searchDN.ui"))[0]
login_class = uic.loadUiType(resource_path("login.ui"))[0]
softphonelogout_class = uic.loadUiType(resource_path("softphoneLogout.ui"))[0]
recipupload_class = uic.loadUiType(resource_path("RECIPUpload.ui"))[0]

class ModifyDialog(QDialog, showModifyWindow_class):
    def __init__(self, DN, i, j, flag):
        super().__init__()
        self.setupUi(self)
        self.DN = DN
        self.i = i
        self.j = j
        self.flag = flag
        self.checkDN(DN)

        _conn = pymssql.connect(server=cmslogin['host'], user=cmslogin['user'], password=cmslogin['passwd'], database=cmslogin['db'], charset=cmslogin['charset'])
        cursor = _conn.cursor(as_dict=True)
        sql = "exec SCMKCPMModQry_V2 'Q', '{DN}', '{CoordX}', '{CoordY}', '', '', '', '{flag}'".format(DN=self.DN, CoordX=self.i, CoordY=self.j, flag=self.flag)
        print(sql)
        cursor.execute(sql)
        result = cursor.fetchall()
        _conn.close()

        print(result)

        self.DN_LineEdit.setText(result[0]["DN"])
        self.PHONEIP_LineEdit.setText(result[0]["Phone_IP"])
        self.PCIP_LineEdit.setText(result[0]["PC_IP"])
        self.ETC_LineEdit.setText(result[0]["Remark1"])

        self.saveButton.clicked.connect(self.save)
        self.deleteButton.clicked.connect(self.delete)

    def save(self):
        response = QMessageBox.question(self, 'save', '저장하시겠습니까?',QMessageBox.Yes | QMessageBox.No)

        if response == QMessageBox.Yes:
            conn = pymssql.connect(server=cmslogin['host'], user=cmslogin['user'], password=cmslogin['passwd'], database=cmslogin['db'], charset=cmslogin['charset'])
            cursor = conn.cursor(as_dict=True)

            sql = "exec SCMKCPMModQry_V2 'S', '{DN}', '{CoordX}', '{CoordY}', '{Remark}', '{PhoneIP}', '{PCIP}', '{flag}', '{UserID}', '{UserPW}'".format(
                DN=self.DN_LineEdit.text(), CoordX=self.i, CoordY=self.j, Remark=self.ETC_LineEdit.text(), PhoneIP=self.PHONEIP_LineEdit.text(),
                PCIP=self.PCIP_LineEdit.text(), flag=self.flag, UserID="", UserPW="")
            try:
                cursor.execute(sql)
                result = cursor.fetchall()
                conn.commit()
            except pymssql.DatabaseError as e:
                QMessageBox.warning(self, "Error", e)
            conn.close()
            if result[0]['result'] == 'DUP':
                QMessageBox.warning(self, "Error", "중복된 내선이 존재합니다. Menu-Search 기능을 통해 내선을 검색하세요.")
            elif result[0]['result'] == 'INSERT':
                QMessageBox.information(self, "Inform", "추가되었습니다.")
            elif result[0]['result'] == 'UPDATE':
                QMessageBox.information(self, "Inform", "수정되었습니다.")
            self.close()
        else:
            self.close()

    def delete(self):
        response = QMessageBox.question(self, 'delete', '삭제하시겠습니까?',
                                     QMessageBox.Yes | QMessageBox.No)

        if response == QMessageBox.Yes:
            conn = pymssql.connect(server=cmslogin['host'], user=cmslogin['user'], password=cmslogin['passwd'], database=cmslogin['db'], charset=cmslogin['charset'])
            cursor = conn.cursor(as_dict=True)

            sql = "exec SCMKCPMModQry_V2 'D', '{DN}', '{CoordX}', '{CoordY}', '{Remark}', '{PhoneIP}', '{PCIP}', '{flag}'".format(
                DN=self.DN_LineEdit.text(), CoordX=self.i, CoordY=self.j, Remark=self.ETC_LineEdit.text(), PhoneIP=self.PHONEIP_LineEdit.text(),
                PCIP=self.PCIP_LineEdit.text(), flag=self.flag)
            print(sql)
            try:
                cursor.execute(sql)
                result = cursor.fetchall()
                conn.commit()
            except pymssql.DatabaseError as e:
                QMessageBox.warning(self, "Error", e)
            conn.close()
            if result[0]['result'] == 'DELETE':
                QMessageBox.information(self, "Inform", "삭제되었습니다.")
            conn.close()
            self.close()
        else:
            self.close()

    def checkDN(self, DN):
        conn = pymssql.connect(server=cmslogin['host'], user=cmslogin['user'], password=cmslogin['passwd'], database=cmslogin['db'], charset=cmslogin['charset'])
        cursor = conn.cursor(as_dict=True)
        sql = "exec SCMCallCenterInfoCheckQry_V2 'O', '{}' ".format(DN)
        cursor.execute(sql)
        result = cursor.fetchall()
        conn.close()

        pbx_color = 'lightgreen'
        rec_color = 'lightgreen'
        pbx_title = 'OK'
        rec_title = 'OK'

        print(result)
        if result:
            if result[0]['STATUS'] == 'NPBX':
                pbx_color = 'darkorange'
                pbx_title = 'X'
            elif result[0]['STATUS'] == 'NREC':
                rec_color = 'darkorange'
                rec_title = 'X'
            elif result[0]['STATUS'] == 'NBOTH':
                pbx_color = 'orangered'
                pbx_title = 'X'
                rec_color = 'orangered'
                rec_title = 'X'
        else:
            pbx_title = ''
            rec_title = ''

        PBX_RGB = getattr(self, "PBX_RGB")
        PBX_RGB.setStyleSheet("background-color: " + pbx_color)
        PBX_RGB.setText(pbx_title)
        REC_RGB = getattr(self, "REC_RGB")
        REC_RGB.setStyleSheet("background-color: " + rec_color)
        REC_RGB.setText(rec_title)

class SearchDNDialog(QDialog, searchDN_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.findButton.clicked.connect(self.find)

        self.FLAG = None
        self.CoordX = None
        self.CoordY = None

    def find(self):
        conn = pymssql.connect(server=cmslogin['host'], user=cmslogin['user'], password=cmslogin['passwd'], database=cmslogin['db'], charset=cmslogin['charset'])
        cursor = conn.cursor(as_dict=True)
        sql = "exec SCMKCPMModQry_V2 'Q', '{DN}', '', '', '', '{PhoneIP}', '{PCIP}', '', ''".format(
            DN=self.DN_LineEdit.text(), PhoneIP=self.PHONEIP_LineEdit.text(), PCIP=self.PCIP_LineEdit.text())
        print(sql)
        try:
            cursor.execute(sql)
            result = cursor.fetchall()
            conn.commit()
        except pymssql.DatabaseError as e:
            QMessageBox.warning(self, "Error", e)

        conn.close()

        if result:
            self.FLAG = result[0]["FLAG"]
            self.CoordX = result[0]["CoordX"]
            self.CoordY = result[0]["CoordY"]
            QMessageBox.information(self, "result", "내선번호 : " + result[0]["DN"] + " " + "위치 : " + result[0]["FLAG"] + " " + "좌표 : ( "+ result[0]["CoordX"] + ", "  + result[0]["CoordY"] + ")" )
            self.close()
        else:
            QMessageBox.information(self, "result", "등록된 내선번호가 없습니다.")

# class refreshTab_using_thread(QThread):
#     def __init__(self, parent=None):
#         super().__init__(parent)
#     def run(self):
#         win = WindowClass()
#         win.refresh_tab()

class WindowClass(QMainWindow, form_class) :
    def __init__(self) :
        super().__init__()
        self.setupUi(self)
        self.initUI()

    def initUI(self):
        self.init_tab_name()
        self.refresh_tab()
        self.actionSearchDN.triggered.connect(self.searchDNMenuClicked)
        self.actionCheckStatus.triggered.connect(self.checkStatus)
        self.actionRefresh.triggered.connect(self.refresh_tab)
        self.actionSoftPhoneLogOut.triggered.connect(self.SoftPhoneLogoutClicked)
        self.actionRECIPUpload.triggered.connect(self.RECIPUploadClicked)

    def refresh_tab(self):
        self.make_tab("D12F")
        self.make_tab("D9F")
        self.make_tab("TECHD")
        self.make_tab("Y3F")
        self.make_tab("TECH")
        self.make_tab("CHUNCHEON")
        self.make_tab("BUSAN")
        self.make_tab("SEONGSU")
        self.make_tab("MULLAE")
        self.make_tab("REPAIR")
        self.make_tab("WORK_HOME")
        self.make_tab("CENTER")
        self.make_rec_tab()

    def init_tab_name(self):
        self.tabWidget.setTabText(0, "대전12층")
        tab1_Textlabel1 = self.tab1_label1
        tab1_Textlabel1.setText("대전12층")

        self.tabWidget.setTabText(1, "대전9층")
        tab2_Textlabel1 = self.tab2_label1
        tab2_Textlabel1.setText("대전9층")

        self.tabWidget.setTabText(2, "기술상담센터_대전")
        tab3_Textlabel1 = self.tab3_label1
        tab3_Textlabel1.setText("기술상담센터_대전")

        self.tabWidget.setTabText(3, "여의도3층")
        tab4_Textlabel1 = self.tab4_label1
        tab4_Textlabel1.setText("여의도3층")

        self.tabWidget.setTabText(4, "기술상담센터_여의도")
        tab5_Textlabel1 = self.tab5_label1
        tab5_Textlabel1.setText("기술상담센터_여의도")

        self.tabWidget.setTabText(5, "춘천_한국고용정보")
        tab6_Textlabel1 = self.tab6_label1
        tab6_Textlabel1.setText("춘천_한국고용정보")

        self.tabWidget.setTabText(6, "부산_한국고용정보")
        tab7_Textlabel1 = self.tab7_label1
        tab7_Textlabel1.setText("부산_한국고용정보")

        self.tabWidget.setTabText(7, "성수_한국고용정보")
        tab8_Textlabel1 = self.tab8_label1
        tab8_Textlabel1.setText("성수_한국고용정보")

        self.tabWidget.setTabText(8, "문래_한국코퍼레이션")
        tab9_Textlabel1 = self.tab9_label1
        tab9_Textlabel1.setText("문래_한국코퍼레이션")

        self.tabWidget.setTabText(9, "리페어센터")
        tab10_Textlabel1 = self.tab10_label1
        tab10_Textlabel1.setText("리페어센터")

        self.tabWidget.setTabText(10, "재택근무")
        tab11_Textlabel1 = self.tab11_label1
        tab11_Textlabel1.setText("재택근무")

        self.tabWidget.setTabText(11, "센터")
        tab11_Textlabel1 = self.tab12_label1
        tab11_Textlabel1.setText("센터")

        self.tabWidget.setTabText(12, "녹취 현황")

    def search_flag(self, flag):
        if flag == "D12F":
            return 1
        elif flag == "D9F":
            return 2
        elif flag == "TECHD":
            return 3
        elif flag == "Y3F":
            return 4
        elif flag == "TECH":
            return 5
        elif flag == "CHUNCHEON":
            return 6
        elif flag == "BUSAN":
            return 7
        elif flag == "SEONGSU":
            return 8
        elif flag == "MULLAE":
            return 9
        elif flag == "REPAIR":
            return 10
        elif flag == "CENTER":
            return 11
        elif flag == "WORK_HOME":
            return 12


    def make_tab(self, flag):
        blank_line = [2, 5, 8, 11, 14, 17, 20, 23, 26, 29]

        index = self.search_flag(flag)

        _conn = pymssql.connect(server=cmslogin['host'], user=cmslogin['user'], password=cmslogin['passwd'], database=cmslogin['db'], charset=cmslogin['charset'])
        cursor = _conn.cursor(as_dict=True)

        sql = "exec SCMKCPMModQry_V2 'M', '', '', '', '', '', '', '{}'".format(flag)
        cursor.execute(sql)
        result = cursor.fetchall()
        _conn.close()
        Seat_Total = str(len(result))

        tab_grid1 = getattr(self, "tab{}_grid1".format(index))

        for i in range(0, 10):
            for j in range(0, 29):
                if j in blank_line:
                    globals()["tab{}_seat{}_{}".format(index, i, j)] = QLabel()
                    globals()["tab{}_seat{}_{}".format(index, i, j)].setMaximumHeight(500)
                    tab_grid1.addWidget(globals()["tab{}_seat{}_{}".format(index, i, j)], i, j)

                else:
                    globals()["tab{}_seat{}_{}".format(index, i, j)] = QPushButton()
                    globals()["tab{}_seat{}_{}".format(index, i, j)].setMaximumHeight(500)
                    globals()["tab{}_seat{}_{}".format(index, i, j)].setMaximumWidth(50)


                    df = pd.DataFrame(result)
                    if df['CoordX'] is not None:
                        if df['CoordY'] is not None:
                            result_df = df.loc[(df['CoordX'] == str(i)) & (df['CoordY'] == str(j))]
                    list_df = result_df.to_dict('records')

                    if list_df:
                        DN = list_df[0]['DN']
                        phoneip = list_df[0]['Phone_IP']
                        pcip = list_df[0]['PC_IP']
                        remark = list_df[0]['Remark1'],
                        flag = list_df[0]['Flag']
                        result_text = str(DN) + '\n' + str(phoneip) + '\n' + str(pcip) + '\n' + str(remark)
                    else:
                        DN = ''
                        phoneip = ''
                        pcip = ''
                        remark = ''
                        flag = flag
                        result_text = ''
                    globals()["tab{}_seat{}_{}".format(index, i, j)].setText(result_text)
                    tab_grid1.addWidget(globals()["tab{}_seat{}_{}".format(index, i, j)], i, j)
                    globals()["tab{}_seat{}_{}".format(index, i, j)].clicked.connect(partial(self.pushSeatButtonClicked, DN, i, j, flag))

        tab1_Textlabel2 = getattr(self, "tab{}_label2".format(index))
        tab1_Textlabel2.setText(Seat_Total)

    def make_rec_tab(self):
        _conn = pymssql.connect(server=cmslogin['host'], user=cmslogin['user'], password=cmslogin['passwd'], database=cmslogin['db'], charset=cmslogin['charset'])
        cursor = _conn.cursor(as_dict=True)

        sql = "exec SCMKCPMModQry 'R', '', '', '', '', '', '', ''"
        cursor.execute(sql)
        result = cursor.fetchall()

        sql = "exec SCMKCPMModQry 'R2', '', '', '', '', '', '', 'IPREC_01'"
        cursor.execute(sql)
        result2 = cursor.fetchall()

        sql = "exec SCMKCPMModQry 'R2', '', '', '', '', '', '', 'IPREC_02'"
        cursor.execute(sql)
        result3 = cursor.fetchall()

        sql = "exec SCMKCPMModQry 'R2', '', '', '', '', '', '', 'IPREC_03'"
        cursor.execute(sql)
        result4 = cursor.fetchall()

        sql = "select ROW_NUMBER() OVER (PARTITION BY 1 ORDER BY DN) AS SEQ, DN from TRCCallCenterPBXInfo_KCPM where left(DN, 1) <> '3'"
        cursor.execute(sql)
        result5 = cursor.fetchall()
        _conn.close()

        rec1_status = "사용현황 : " + str(result[0]['REC1_USE']) + "/" + str(result[0]['REC1_TOTAL'])

        tab13_rec1_label = getattr(self, "tab13_rec1_label")
        tab13_rec1_label.setText(rec1_status)

        for i in range(0, len(result2)):
            self.tab13_rec1.setItem(i, 0, QTableWidgetItem(str(result2[i]["SEQ"])))
            self.tab13_rec1.setItem(i, 1, QTableWidgetItem(result2[i]["DN"]))
            self.tab13_rec1.setItem(i, 2, QTableWidgetItem(result2[i]["Phone_IP"]))
            if i != len(result2)-1:
                row_count = self.tab13_rec1.rowCount()
                self.tab13_rec1.setRowCount(row_count + 1)  # row 추가

        rec2_status = "사용현황 : " + str(result[0]['REC2_USE']) + "/" + str(result[0]['REC2_TOTAL'])

        tab13_rec2_label = getattr(self, "tab13_rec2_label")
        tab13_rec2_label.setText(rec2_status)

        for i in range(0, len(result3)):
            self.tab13_rec2.setItem(i, 0, QTableWidgetItem(str(result3[i]["SEQ"])))
            self.tab13_rec2.setItem(i, 1, QTableWidgetItem(result3[i]["DN"]))
            self.tab13_rec2.setItem(i, 2, QTableWidgetItem(result3[i]["Phone_IP"]))
            if i != len(result3) - 1:
                row_count = self.tab13_rec2.rowCount()
                self.tab13_rec2.setRowCount(row_count + 1)  # row 추가

        rec3_status = "사용현황 : " + str(result[0]['REC3_USE']) + "/" + str(result[0]['REC3_TOTAL'])

        tab13_rec3_label = getattr(self, "tab13_rec3_label")
        tab13_rec3_label.setText(rec3_status)

        for i in range(0, len(result4)):
            self.tab13_rec3.setItem(i, 0, QTableWidgetItem(str(result4[i]["SEQ"])))
            self.tab13_rec3.setItem(i, 1, QTableWidgetItem(result4[i]["DN"]))
            self.tab13_rec3.setItem(i, 2, QTableWidgetItem(result4[i]["Phone_IP"]))
            if i != len(result4) - 1:
                row_count = self.tab13_rec3.rowCount()
                self.tab13_rec3.setRowCount(row_count + 1)  # row 추가

        PBX_status = "사용현황 : " + str(len(result5))

        tab13_PBX_label = getattr(self, "tab13_PBX_label")
        tab13_PBX_label.setText(PBX_status)

        for i in range(0, len(result5)):
            self.tab13_PBX.setItem(i, 0, QTableWidgetItem(str(result5[i]["SEQ"])))
            self.tab13_PBX.setItem(i, 1, QTableWidgetItem(result5[i]["DN"]))
            if i != len(result5) - 1:
                row_count = self.tab13_PBX.rowCount()
                self.tab13_PBX.setRowCount(row_count + 1)  # row 추가



    def pushSeatButtonClicked(self, DN, i, j, flag):
        ModDlg = ModifyDialog(DN, i, j, flag)
        ModDlg.exec_()
        self.make_tab(flag)

    def searchDNMenuClicked(self):
        SearchDlg = SearchDNDialog()
        SearchDlg.exec_()
        global g_FLAG
        if SearchDlg.FLAG is not None and SearchDlg.CoordX is not None and SearchDlg.CoordY is not None:
            CoordX = SearchDlg.CoordX
            CoordY = SearchDlg.CoordY
            index = self.search_flag(SearchDlg.FLAG)
            self.tabWidget.setCurrentIndex(index-1)
            if g_FLAG:
                self.make_tab(g_FLAG)
            DN_Button = globals()["tab{}_seat{}_{}".format(index, CoordX, CoordY)]
            DN_Button.setStyleSheet("background-color: red")
            g_FLAG = SearchDlg.FLAG

    def checkStatus(self):
        conn = pymssql.connect(server=cmslogin['host'], user=cmslogin['user'], password=cmslogin['passwd'], database=cmslogin['db'], charset=cmslogin['charset'])
        cursor = conn.cursor(as_dict=True)
        sql = "exec SCMCallCenterInfoCheckQry_V2 'A', '' "
        cursor.execute(sql)
        result = cursor.fetchall()
        conn.close()
        for i in range(0, len(result)):
            pageNum = self.search_flag(result[i]['FLAG'])
            if result[i]['STATUS'] == 'NPBX':
                color = 'darkorange'
            elif result[i]['STATUS'] == 'NREC':
                color = 'darkorange'
            elif result[i]['STATUS'] == 'NBOTH':
                color = 'orangered'
            elif result[i]['STATUS'] == 'OK':
                color = 'lightgreen'
            Button = globals()["tab{}_seat{}_{}".format(pageNum, result[i]['CoordX'], result[i]['CoordY'])]
            Button.setStyleSheet("background-color: " + color)

    def SoftPhoneLogoutClicked(self):
        SoftPhoneLogoutDlg = SoftPhoneLogOutDialog()
        SoftPhoneLogoutDlg.exec_()

    def RECIPUploadClicked(self):
        RECIPUploadDlg = RECIPUploadDialog()
        RECIPUploadDlg.exec_()


class SoftPhoneLogOutDialog(QDialog, softphonelogout_class):
    def __init__(self) :
        super().__init__()
        self.setupUi(self)
        self.DN_LineEdit.setFocus()
        self.LogoutButton.clicked.connect(self.logout)

    def logout(self):
        DN = self.DN_LineEdit.text()
        AgentID = self.AgentID_LineEdit.text()

        if DN == '' or AgentID == '':
            QMessageBox.information(self, "error", "내선번호와 AgentID를 둘 다 입력하세요")

        else:
            cli = paramiko.SSHClient()
            cli.set_missing_host_key_policy(paramiko.AutoAddPolicy)
            cli.connect(hostname=cubelogin_1['host'], username=cubelogin_1['user'], password=cubelogin_1['passwd'])

            channel = cli.invoke_shell()
            channel.send("home\n")
            time.sleep(0.5)
            channel.send("bin\n")
            time.sleep(0.5)
            channel.send("emglogout -dn {}\n".format(DN))
            time.sleep(1)
            channel.send("y\n")
            time.sleep(0.5)
            channel.send("emglogout -emp {}\n".format(AgentID))
            time.sleep(1)
            channel.send("y\n")
            time.sleep(0.5)

            QMessageBox.information(self, "notice", "내선번호 [" + DN + "] , Agent ID [" + AgentID + "] 1번 서버에서 정상적으로 로그아웃 되었습니다.")

            cli2 = paramiko.SSHClient()
            cli2.set_missing_host_key_policy(paramiko.AutoAddPolicy)
            cli2.connect(hostname=cubelogin_2['host'], username=cubelogin_2['user'], password=cubelogin_2['passwd'])

            channel = cli2.invoke_shell()
            channel.send("home\n")
            time.sleep(0.5)
            channel.send("bin\n")
            time.sleep(0.5)
            channel.send("emglogout -dn {}\n".format(DN))
            time.sleep(1)
            channel.send("y\n")
            time.sleep(0.5)
            channel.send("emglogout -emp {}\n".format(AgentID))
            time.sleep(1)
            channel.send("y\n")
            time.sleep(0.5)

            QMessageBox.information(self, "notice", "내선번호 [" + DN + "] , Agent ID [" + AgentID + "] 2번 서버에서 정상적으로 로그아웃 되었습니다.")

class RECIPUploadDialog(QDialog, recipupload_class):
    def __init__(self) :
        super().__init__()
        self.setupUi(self)
        self.FileUploadBtn.clicked.connect(self.fileUpload)
        self.saveBtn.clicked.connect(self.saveToDB)
        self.refreshBtn.clicked.connect(self.refreshRow)

    def fileUpload(self):
        path = easygui.fileopenbox()
        file = open(path, "r");
        row_index = 0

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

    def saveToDB(self):

        answer = QMessageBox.question(self, "notice", "저장하시겠습니까?")
        if answer == QMessageBox.Yes:
            conn = pymssql.connect(server=cmslogin['host'], user=cmslogin['user'], password=cmslogin['passwd'], database=cmslogin['db'], charset=cmslogin['charset'])
            cursor = conn.cursor(as_dict=True)

            sql = "delete from TRCCallCenterRecordInfo_KCPM"

            try:
                cursor.execute(sql)
                conn.commit()
            except pymssql.DatabaseError as e:
                QMessageBox.warning(self, "Error", e)
            conn.close()

            conn = pymssql.connect(server=cmslogin['host'], user=cmslogin['user'], password=cmslogin['passwd'], database=cmslogin['db'], charset=cmslogin['charset'])
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

    def refreshRow(self):
        for i in reversed(range(self.ip_table.rowCount())):
            self.ip_table.removeRow(i)

class LoginDialog(QMainWindow, login_class):
    def __init__(self) :
        super().__init__()
        self.setupUi(self)
        self.ID_LineEdit.setFocus()
        self.qLoginButton.clicked.connect(self.login)
        self.PW_LineEdit.returnPressed.connect(self.login)

        self.qPixmapFileVar = QPixmap()
        self.qPixmapFileVar.load(resource_path("KD.png"))
        self.KD_label.setPixmap(self.qPixmapFileVar)

    def login(self):
        _conn = pymssql.connect(server=cmslogin['host'], user=cmslogin['user'], password=cmslogin['passwd'], database=cmslogin['db'], charset=cmslogin['charset'])
        cursor = _conn.cursor(as_dict=True)

        global LoginId
        LoginId = self.ID_LineEdit.text()

        sql = "exec SCTUserLoginQryNew_KCPM '{ID}', '{PW}'".format(ID=self.ID_LineEdit.text(), PW=self.PW_LineEdit.text())
        cursor.execute(sql)
        result = cursor.fetchall()
        _conn.close()

        if result[0]['rType'] == 'OK':
            self.close()
            MainWindowClass = WindowClass()
            MainWindowClass.show()
        elif result[0]['rType'] == 'WU':
            QMessageBox.information (self, "error", "ID가 일치하지 않습니다. CIC 계정을 확인하세요.")
        elif result[0]['rType'] == 'WP':
            QMessageBox.information(self, "error", "PW가 일치하지 않습니다. CIC 계정을 확인하세요.")
        elif result[0]['rType'] == 'NE':
            QMessageBox.information(self, "error", "계정이 존재하지 않습니다. CIC 계정을 확인하세요.")
        elif result[0]['rType'] == 'NA':
            QMessageBox.information(self, "error", "권한이 없습니다. CIC 권한을 확인하세요.")
    # def refresh_tab_using_thread(self):
    #     x = refreshTab_using_thread(self)
    #     x.start()


if __name__ == "__main__" :
    app = QApplication(sys.argv)
    myWindow = LoginDialog()
    myWindow.show()
    app.exec_()

