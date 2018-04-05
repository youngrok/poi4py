import os
import unittest

import poi4py


class ExcelTest(unittest.TestCase):

    def test_read(self):
        poi4py.start_jvm()
        workbook = poi4py.open_workbook(os.path.dirname(__file__) + '/소득자별근로소득원천징수부_20180403133614.xls', password='1111')
        sheet = workbook.getSheetAt(0)
        for i in range(3, sheet.getLastRowNum() + 1):
            row = sheet.getRow(i)
            print(row.getCell(0).getStringCellValue(), row.getCell(1).getStringCellValue(), row.getCell(2).getNumericCellValue())

        poi4py.shutdown_jvm()
