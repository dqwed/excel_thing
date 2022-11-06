from contextlib import suppress
from UI_main import Ui_Dialog
from PyQt5.QtWidgets import QWidget, QFileDialog
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
        res_dict = self.create_barcode_rest_dict()

        df = ExcelFile(self.files['pushButton_2']).parse(header=8, usecols='F')
        barcodes = DataFrame(df).to_numpy(dtype=int)
        rest = ['', '', 'Остаток в офисе']

        for barcode in barcodes:
            if str(barcode[0]) in res_dict:
                rest.append(res_dict[str(barcode[0])] if res_dict[str(barcode[0])] != '' else 0)
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
