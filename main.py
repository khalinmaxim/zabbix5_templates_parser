from pyzabbix import ZabbixAPI
from openpyxl import *


class Parser:
    def __init__(self, hostid, templateids):
        self.hostid = hostid
        self.templateids = templateids
        self.z = ZabbixAPI("https", user="user", password="password")
        self.items = self.z.item.get(hostid=self.hostid, templateids=self.templateids)

    def excel(self):
        wb = Workbook()
        sheet = wb.active
        sheet.title = 'TemplateID ' + str(self.templateids)
        column = 1
        for item in self.items[0].keys():
            sheet.cell(row=1, column=column).value = item
            column += 1
        for row in range(len(self.items)):
            column = 1
            for keys in self.items[row].keys():
                sheet.cell(row=row+2, column=column).value = str(self.items[row][keys])
                column += 1
        wb.save('parser.xlsx')


if __name__ == "__main__":
    parser = Parser("hostid", "templateids")
    parser.excel()
