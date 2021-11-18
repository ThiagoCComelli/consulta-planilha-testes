from pycep_correios import get_address_from_cep, WebService
import openpyxl

class Planilha:
    def __init__(self,name):
        self.fileName = name
        self.workbook = openpyxl.load_workbook(f"planilhas/{self.fileName}.xlsx")
        self.worksheet = self.workbook["Orçamento "]

    def init(self):
        for index, row in enumerate(self.worksheet.iter_rows(9,539)):
            print(index)
            cep = str(row[3].value)[:5] + "-" + str(row[3].value)[5:]

            try:
                address = get_address_from_cep(cep, webservice=WebService.CORREIOS)
                self.newCells(address,row)
            except Exception as err:
                row[0].value = str(err)

        self.workbook.save(f"planilhas/{self.fileName}-Modified.xlsx")
    
    def newCells(self,address,row):
        row[4].value = address["logradouro"]
        row[5].value = address["bairro"]
        row[6].value = address["cidade"]