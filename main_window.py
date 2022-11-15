from contextlib import suppress
from UI_main import Ui_Dialog
from PyQt5.QtWidgets import QWidget, QFileDialog, QMessageBox
from pandas import ExcelFile, DataFrame


class Main(QWidget, Ui_Dialog):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        self.pushButton_3.hide()

        self.pushButton.clicked.connect(lambda: self.openFileNameDialog(self.pushButton))
        self.pushButton_2.clicked.connect(lambda: self.openFileNameDialog(self.pushButton_2))
        self.files = {}

        self.data = []
        self.pushButton_3.clicked.connect(self.process)

    def openFileNameDialog(self, btn):
        file, _ = QFileDialog.getOpenFileName(self, 'Open File', './', 'Excel File (*.xlsx *.xls)')
        if file:
            self.files[btn.objectName()] = file
            btn.setText(file.split("/")[-1][:30])
        if len(self.files) == 2:
            self.pushButton_3.show()

    def process(self):
        barcode_column = self.text.text()
        if not(barcode_column.isalpha() and barcode_column.isascii()):
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Ошибка")
            msg.setInformativeText("Не все поля заполнены или неверный фопмат данных.")
            msg.setWindowTitle("Ошибка")
            msg.exec_()
            return
        res_dict = self.create_barcode_rest_dict()

        df = ExcelFile(self.files['pushButton_2']).parse(header=8, usecols=barcode_column.upper())
        barcodes = DataFrame(df).to_numpy()
        rest = ['', '', 'Остаток в офисе']

        for barcode in barcodes:
            barcode = self.cut_barcode(barcode[0])
            if str(barcode) in res_dict:
                rest.append(res_dict[str(barcode)] if res_dict[str(barcode)] != '' else 0)
            else:
                rest.append('не найден')
        df = ExcelFile(self.files['pushButton_2']).parse(header=5)
        df.insert(len(df.columns), ' ', rest)
        df.to_excel('result.xlsx', index=False)

        self.close()

    def create_barcode_rest_dict(self):
        df = ExcelFile(self.files['pushButton']).parse(header=1, usecols='D,E').fillna(0)
        pairs = DataFrame(df).to_numpy()
        data = {}
        for pair in pairs[2:-1]:
            if type(pair[1]) == str:
                barcodes = pair[1].split()
                for barcode in barcodes:
                    if barcode:
                        data[barcode] = pair[0]
        return data

    def header(self):
        pass

    def column(self):
        pass

    def cut_barcode(self, barcode):
        barcode = str(barcode).replace('.', '')[0:-1]
        if len(barcode) <= 13:
            return int(barcode)
        elif len(barcode) == 14:
            return int(barcode[1:])
        elif len(barcode) > 14:
            return int(barcode[:13]) 
