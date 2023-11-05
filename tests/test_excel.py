import unittest
from poexcel.api.excel import *


class TestExcel(unittest.TestCase):
    def test_fake2excel(self):
        fake2excel(language='fdsa', columns=['name', 'text'], rows=200)

    def test_split_excel_by_column(self):
        split_excel_by_column(filepath=r'../../contributors/bulabean/sedemo.xls',
                              column=6)

    def test_sheet2excel(self):
        sheet2excel(file_path=r'/tests/excel_files/fake2excel.xlsx',
                    output_path=r'D:\workplace\code\github\poexcel\tests\output_path')

    def test_merge2sheet(self):
        merge2sheet(dir_path=r'D:\workplace\code\github\python-office\tests\test_files\excel\merge2sheet')

    def test_merge2excel(self):
        merge2excel(dir_path=r'../../contributors/bulabean', output_file='test_merge2excel.xlsx', )

    def test_find_excel_data(self):
        find_excel_data(search_key='刘家站垦殖场', target_dir=r'../../contributors/bulabean')

    def test_query4excel(self):
        # query4excel(query_content='程序员晚枫', query_path=r'D:\test\py310\excel_test')
        query4excel(query_content='程序员晚枫', query_path=r'D:\test\py310\excel_test\course',
                    output_path=r'D:\test\py310\excel_test\output_path', output_name='晚枫')

    def test_excel2pdf(self):
        excel2pdf(excel_path=r'./excel_files/fake2excel.xlsx', pdf_path=r'./output_path')
