import glob
import os
import sys
from pathlib import PurePath

import pythoncom
import win32clipboard
import win32com.client
import yaml
from PyQt6 import QtWidgets, uic


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        uic.loadUi("MainWindow.ui", self)

        self.excel = None

        # combobox
        settings_list = [PurePath(path).stem for path in glob.glob(f"{os.getcwd()}/settings/*.yaml")]
        self.setting_cb.addItems(settings_list)
        self.setting_cb.currentTextChanged.connect(self.setting_cb_channged)

        # detail ckb
        self.exist_detail_ckb.stateChanged.connect(self.detail_ckb_changed)

        # save button
        self.blackacre_setting_save_btn.clicked.connect(self.blackacre_setting_saved)
        self.position_setting_save_btn.clicked.connect(self.position_setting_saved)

        # generator button
        self.generator_btn.clicked.connect(self.generator)

    def setting_cb_channged(self):
        self.path = os.getcwd() + "\\settings\\" + self.setting_cb.currentText() + ".yaml"
        try:
            self.log_pte.appendPlainText(f"{self.setting_cb.currentText()} 설정 로딩중..")
            with open(self.path, "r", encoding="utf-8") as f:
                self.setting = yaml.load(f, Loader=yaml.FullLoader)

            self.item_name_te.setText(self.setting["blackacre"]["item_name"])
            self.grade_te.setText(self.setting["blackacre"]["grade"])
            self.quantity_te.setText(self.setting["blackacre"]["quantity"])
            self.serial_no_te.setText(self.setting["blackacre"]["serial_no"])
            self.lot_no_te.setText(self.setting["blackacre"]["lot_no"])
            self.due_date_te.setText(self.setting["blackacre"]["due_date"])
            self.delivery_date_te.setText(self.setting["blackacre"]["delivery_date"])
            self.inspected_te.setText(self.setting["blackacre"]["inspected_by"])
            self.approved_te.setText(self.setting["blackacre"]["approved_by"])
            self.serial_num_loc_te.setText(self.setting["raw"]["serial_num_loc"])
            self.raw_data_loc_te.setText(self.setting["raw"]["raw_data_loc"])
            self.copy_data_loc_te.setText(self.setting["raw"]["copy_data_loc"])
            self.inspection_serial_loc_te.setText(self.setting["inspection"]["inspection_serial_loc"])
            self.inspection_data_loc_te.setText(self.setting["inspection"]["inspection_data_loc"])
            self.inspection_data_num_te.setText(self.setting["inspection"]["inspection_data_num"])
            if bool(self.setting["detail"]["exist_detail"]):
                self.detail_raw_data_loc_te.setDisabled(False)
                self.detail_copy_data_loc_te.setDisabled(False)

                self.exist_detail_ckb.setChecked(bool(self.setting["detail"]["exist_detail"]))
                self.detail_raw_data_loc_te.setText(self.setting["detail"]["detail_raw_data_loc"])
                self.detail_copy_data_loc_te.setText(self.setting["detail"]["detail_copy_data_loc"])
            else:
                self.detail_raw_data_loc_te.setDisabled(True)
                self.detail_copy_data_loc_te.setDisabled(True)

                self.exist_detail_ckb.setChecked(bool(self.setting["detail"]["exist_detail"]))
                self.detail_raw_data_loc_te.clear()
                self.detail_copy_data_loc_te.clear()

            self.log_pte.appendPlainText("로딩 완료!")
        except:
            pass

    def detail_ckb_changed(self, state):
        if state:
            self.detail_raw_data_loc_te.setDisabled(False)
            self.detail_copy_data_loc_te.setDisabled(False)
        else:
            self.detail_raw_data_loc_te.setDisabled(True)
            self.detail_copy_data_loc_te.setDisabled(True)

    def blackacre_setting_saved(self):
        self.setting["blackacre"]["item_name"] = self.item_name_te.toPlainText().upper()
        self.setting["blackacre"]["grade"] = self.grade_te.toPlainText().upper()
        self.setting["blackacre"]["quantity"] = self.quantity_te.toPlainText().upper()
        self.setting["blackacre"]["serial_no"] = self.serial_no_te.toPlainText().upper()
        self.setting["blackacre"]["lot_no"] = self.lot_no_te.toPlainText().upper()
        self.setting["blackacre"]["due_date"] = self.due_date_te.toPlainText().upper()
        self.setting["blackacre"]["delivery_date"] = self.delivery_date_te.toPlainText().upper()
        self.setting["blackacre"]["inspected_by"] = self.inspected_te.toPlainText().upper()
        self.setting["blackacre"]["approved_by"] = self.approved_te.toPlainText().upper()

        self.log_pte.appendPlainText(f"{self.setting_cb.currentText()} 설정 저장중..")
        with open(self.path, "w", encoding="utf-8") as f:
            yaml.dump(self.setting, f)
        self.log_pte.appendPlainText("저장 완료!")

    def position_setting_saved(self):
        self.setting["raw"]["serial_num_loc"] = self.serial_num_loc_te.toPlainText().upper()
        self.setting["raw"]["raw_data_loc"] = self.raw_data_loc_te.toPlainText().upper()
        self.setting["raw"]["copy_data_loc"] = self.copy_data_loc_te.toPlainText().upper()
        self.setting["inspection"]["inspection_serial_loc"] = self.inspection_serial_loc_te.toPlainText().upper()
        self.setting["inspection"]["inspection_data_loc"] = self.inspection_data_loc_te.toPlainText().upper()
        self.setting["inspection"]["inspection_data_num"] = self.inspection_data_num_te.toPlainText().upper()
        self.setting["detail"]["exist_detail"] = self.exist_detail_ckb.isChecked()
        self.setting["detail"]["detail_raw_data_loc"] = self.detail_raw_data_loc_te.toPlainText().upper()
        self.setting["detail"]["detail_copy_data_loc"] = self.detail_copy_data_loc_te.toPlainText().upper()

        self.log_pte.appendPlainText(f"{self.setting_cb.currentText()} 설정 저장중..")
        with open(self.path, "w", encoding="utf-8") as f:
            yaml.dump(self.setting, f)
        self.log_pte.appendPlainText("저장 완료!")

    def generator(self):
        if self.excel != None:
            self.excel.Quit()
        self.set_clipboard()  # 클립보드 초기화

        self.log_pte.appendPlainText(f"{self.setting_cb.currentText()} 성적서 생성중..")
        self.main_pgb.reset()
        self.item_name = self.setting_cb.currentText()

        pythoncom.CoInitialize()
        self.excel = win32com.client.Dispatch("Excel.Application")  # 엑셀 실행
        self.excel.Visible = False

        self.blackacre_wb = self.excel.Workbooks.Open(f"{os.getcwd()}\\성적서 양식\\갑지.xls")  # 갑지 양식 열기
        self.tck_template_wb = self.excel.Workbooks.Open(f"{os.getcwd()}\\TCK 양식\\T-{self.item_name}.xls")  # TCK RAW 데이터 양식 열기
        if self.setting["detail"]["exist_detail"]:
            self.tck_detail_template_wb = self.excel.Workbooks.Open(f"{os.getcwd()}\\TCK 양식\\TD-{self.item_name}.xls")  # TCK Detail RAW 데이터 양식 열기

        self.data_list = glob.glob(f"{os.getcwd()}\\측정 데이터\\*.xls")  # 측정 데이터 리스트 불러오기

        workbook = self.excel.Workbooks.Add()

        # 갑지 시트 열기
        blackacre_sheet = self.blackacre_wb.Worksheets("갑지")

        # 갑지 시트 데이터 입력
        blackacre_sheet.Range("F11").value = self.setting["blackacre"]["item_name"]  # DWG NO
        blackacre_sheet.Range("F12").value = self.setting["blackacre"]["grade"]  # Grade
        blackacre_sheet.Range("F13").value = self.setting["blackacre"]["quantity"]  # Quantity
        blackacre_sheet.Range("F14").value = self.setting["blackacre"]["serial_no"]  # Serial NO
        blackacre_sheet.Range("F15").value = self.setting["blackacre"]["lot_no"]  # Material LOT NO
        blackacre_sheet.Range("F16").value = self.setting["blackacre"]["due_date"]  # due date
        blackacre_sheet.Range("F17").value = self.setting["blackacre"]["delivery_date"]  # delivery date
        blackacre_sheet.Range("E27").value = self.setting["blackacre"]["inspected_by"]  # inspected by
        blackacre_sheet.Range("O27").value = self.setting["blackacre"]["approved_by"]  # Approved BY

        # 갑지 시트 복사
        ws = workbook.Worksheets.Add(Before=workbook.Worksheets(workbook.Worksheets.Count))
        ws.Name = "갑지"
        blackacre_sheet.Cells.Copy()
        ws.Cells.PasteSpecial(Paste=14)
        self.set_clipboard()
        self.blackacre_wb.Close(False)

        # 데이터 시트 열기
        tck_sheet1 = self.tck_template_wb.Worksheets("Sheet1")
        if self.setting["detail"]["exist_detail"]:
            tck_detail_sheet1 = self.tck_detail_template_wb.Worksheets("Sheet1")

        data_cnt = 0
        sheet_cnt = 1
        pgb_value = 0
        for file_name in self.data_list:
            if data_cnt == 0:
                self.inspection_wb = self.excel.Workbooks.Open(f"{os.getcwd()}\\성적서 양식\\{self.item_name}.xls")  # 데이터 양식 열기
                self.inspection_sheet = self.inspection_wb.Worksheets("DATA")

            raw_data_wb = self.excel.Workbooks.Open(file_name)  # raw data 열기
            raw_sheet = raw_data_wb.Worksheets("Sheet1")  # raw data sheet 열기

            if self.setting["detail"]["exist_detail"] and raw_sheet.Range(self.setting["detail"]["detail_raw_data_loc"].split(":")[-1]).Value != None:
                raw_sheet.Range(self.setting["detail"]["detail_raw_data_loc"]).Copy()  # raw data 복사
                tck_detail_sheet1.Range(self.setting["detail"]["detail_raw_data_loc"]).PasteSpecial(Paste=-4163)  # raw data를 tck 양식에 복사
                raw_sheet.Range(self.setting["raw"]["serial_num_loc"]).Copy()  # raw data serial num 복사
                row = self.setting["inspection"]["inspection_serial_loc"][1:]
                col = self.setting["inspection"]["inspection_serial_loc"][0]
                loc = chr(ord(col) + data_cnt).upper() + row
                self.inspection_sheet.Range(loc).PasteSpecial(Paste=-4163)  # raw data serial num 붙여넣기

                tck_cmm_sheet = self.tck_detail_template_wb.Worksheets("CMM")
                tck_cmm_sheet.Range(self.setting["detail"]["detail_copy_data_loc"]).Copy()  # 최종 데이터 복사
                row = self.setting["inspection"]["inspection_data_loc"][1:]
                col = self.setting["inspection"]["inspection_data_loc"][0]
                loc = chr(ord(col) + data_cnt).upper() + row
                self.inspection_sheet.Range(loc).PasteSpecial(Paste=-4163)  # 성적서 시트에 붙여넣기
            else:
                raw_sheet.Range(self.setting["raw"]["raw_data_loc"]).Copy()  # raw data 복사
                tck_sheet1.Range(self.setting["raw"]["raw_data_loc"]).PasteSpecial(Paste=-4163)  # raw data를 tck 양식에 복사
                raw_sheet.Range(self.setting["raw"]["serial_num_loc"]).Copy()  # raw data serial num 복사
                row = self.setting["inspection"]["inspection_serial_loc"][1:]
                col = self.setting["inspection"]["inspection_serial_loc"][0]
                loc = chr(ord(col) + data_cnt).upper() + row
                self.inspection_sheet.Range(loc).PasteSpecial(Paste=-4163)  # raw data serial num 붙여넣기

                tck_cmm_sheet = self.tck_template_wb.Worksheets("CMM")
                tck_cmm_sheet.Range(self.setting["raw"]["copy_data_loc"]).Copy()  # 최종 데이터 복사
                row = self.setting["inspection"]["inspection_data_loc"][1:]
                col = self.setting["inspection"]["inspection_data_loc"][0]
                loc = chr(ord(col) + data_cnt).upper() + row
                self.inspection_sheet.Range(loc).PasteSpecial(Paste=-4163)  # 성적서 시트에 붙여넣기

            data_cnt += 1
            self.set_clipboard()
            raw_data_wb.Close(False)

            if data_cnt == int(self.setting["inspection"]["inspection_data_num"]):
                data_cnt = 0
                ws = workbook.Worksheets.Add(Before=workbook.Worksheets(workbook.Worksheets.Count))
                ws.Name = f"DATA{sheet_cnt}"
                self.inspection_sheet.Cells.Copy()
                ws.Cells.PasteSpecial(Paste=14)
                sheet_cnt += 1
                self.set_clipboard()
                self.inspection_wb.Close(False)

            if pgb_value < 100:
                pgb_value += (1 / len(self.data_list)) * 100
                self.main_pgb.setValue(int(pgb_value))

        if data_cnt != 0:
            ws = workbook.Worksheets.Add(Before=workbook.Worksheets(workbook.Worksheets.Count))
            ws.Name = f"DATA{sheet_cnt}"
            self.inspection_sheet.Cells.Copy()
            ws.Cells.PasteSpecial(Paste=14)
            self.set_clipboard()
            self.inspection_wb.Close(False)

        self.main_pgb.setValue(100)

        self.set_clipboard()
        self.tck_template_wb.Close(False)
        if self.setting["detail"]["exist_detail"]:
            self.tck_detail_template_wb.Close(False)

        self.excel.DisplayAlerts = False
        workbook.Worksheets("Sheet1").Delete()
        self.excel.DisplayAlerts = True

        workbook.SaveAs(f"{os.getcwd()}\\성적서.xls")
        self.log_pte.appendPlainText(f"{self.setting_cb.currentText()} 성적서 생성 완료!")
        self.excel.Quit()
        pythoncom.CoUninitialize()

    def set_clipboard(self):
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.CloseClipboard()


app = QtWidgets.QApplication(sys.argv)
window = MainWindow()
window.show()
app.exec()
