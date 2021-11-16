import requests
import openpyxl

class Planilha:
    def __init__(self,name):
        self.fileName = name
        self.workbook = openpyxl.load_workbook(f"planilhas/{self.fileName}.xlsx")
        self.worksheet = self.workbook["Planilha1"]
        self.api = "https://minhareceita.org/"

    def init(self):
        for col in self.worksheet.iter_cols(0, self.worksheet.max_column):
            if col[0].value == "cnpj":
                for cel in col[1:]:
                    res = requests.get(self.api+cel.value)
                    cell = self.worksheet.cell(row=cel.row, column=cel.column+1)
                    cell.value = res.json()["cep"]
        self.workbook.save(f"planilhas/{self.fileName}-Modified.xlsx")