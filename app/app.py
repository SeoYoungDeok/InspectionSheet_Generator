import glob
import os
import sys
from pathlib import PurePath

import yaml
from PyQt6 import QtWidgets, uic


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        uic.loadUi("MainWindow.ui", self)

        # combobox
        settings_list = [
            PurePath(path).stem for path in glob.glob("../app/settings/*.yaml")
        ]
        self.setting_cb.addItems(settings_list)
        self.setting_cb.currentTextChanged.connect(self.setting_cb_channged)

        # detail ckb
        self.exist_detail_ckb.stateChanged.connect(self.detail_ckb_changed)

        # save button
        self.blackacre_setting_save_btn.clicked.connect(self.blackacre_setting_saved)
        self.position_setting_save_btn.clicked.connect(self.position_setting_saved)

    def setting_cb_channged(self):
        self.path = (
            os.getcwd() + "\\settings\\" + self.setting_cb.currentText() + ".yaml"
        )
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
            self.inspection_serial_loc_te.setText(
                self.setting["inspection"]["inspection_serial_loc"]
            )
            self.inspection_data_loc_te.setText(
                self.setting["inspection"]["inspection_data_loc"]
            )
            self.inspection_data_num_te.setText(
                self.setting["inspection"]["inspection_data_num"]
            )
            if bool(self.setting["detail"]["exist_detail"]):
                self.detail_raw_data_loc_te.setDisabled(False)
                self.detail_copy_data_loc_te.setDisabled(False)

                self.exist_detail_ckb.setChecked(
                    bool(self.setting["detail"]["exist_detail"])
                )
                self.detail_raw_data_loc_te.setText(
                    self.setting["detail"]["detail_raw_data_loc"]
                )
                self.detail_copy_data_loc_te.setText(
                    self.setting["detail"]["detail_copy_data_loc"]
                )
            else:
                self.detail_raw_data_loc_te.setDisabled(True)
                self.detail_copy_data_loc_te.setDisabled(True)

                self.exist_detail_ckb.setChecked(
                    bool(self.setting["detail"]["exist_detail"])
                )
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
        self.setting["blackacre"][
            "delivery_date"
        ] = self.delivery_date_te.toPlainText().upper()
        self.setting["blackacre"][
            "inspected_by"
        ] = self.inspected_te.toPlainText().upper()
        self.setting["blackacre"][
            "approved_by"
        ] = self.approved_te.toPlainText().upper()

        self.log_pte.appendPlainText(f"{self.setting_cb.currentText()} 설정 저장중..")
        with open(self.path, "w", encoding="utf-8") as f:
            yaml.dump(self.setting, f)
        self.log_pte.appendPlainText("저장 완료!")

    def position_setting_saved(self):
        self.setting["raw"][
            "serial_num_loc"
        ] = self.serial_num_loc_te.toPlainText().upper()
        self.setting["raw"]["raw_data_loc"] = self.raw_data_loc_te.toPlainText().upper()
        self.setting["raw"][
            "copy_data_loc"
        ] = self.copy_data_loc_te.toPlainText().upper()
        self.setting["inspection"][
            "inspection_serial_loc"
        ] = self.inspection_serial_loc_te.toPlainText().upper()
        self.setting["inspection"][
            "inspection_data_loc"
        ] = self.inspection_data_loc_te.toPlainText().upper()
        self.setting["inspection"][
            "inspection_data_num"
        ] = self.inspection_data_num_te.toPlainText().upper()
        self.setting["detail"]["exist_detail"] = self.exist_detail_ckb.isChecked()
        self.setting["detail"][
            "detail_raw_data_loc"
        ] = self.detail_raw_data_loc_te.toPlainText().upper()
        self.setting["detail"][
            "detail_copy_data_loc"
        ] = self.detail_copy_data_loc_te.toPlainText().upper()

        self.log_pte.appendPlainText(f"{self.setting_cb.currentText()} 설정 저장중..")
        with open(self.path, "w", encoding="utf-8") as f:
            yaml.dump(self.setting, f)
        self.log_pte.appendPlainText("저장 완료!")


app = QtWidgets.QApplication(sys.argv)
window = MainWindow()
window.show()
app.exec()
