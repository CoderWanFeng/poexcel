import search4file
import win32com
from faker import Faker
import pandas as pd

import os
from pathlib import Path
from openpyxl import load_workbook

from pofile import get_files
from poprogress import simple_progress
from tqdm import tqdm
from pathlib import Path
# 忽略waring警告
from poexcel.lib import pandas_mem
from poexcel.lib.excel import SplitExcel
import xlwings as xw


class MainExcel():

    def fake2excel(self, columns, rows, path, language):
        """
        @Author & Date  : CoderWanFeng 2022/5/13 0:12
        @Desc  : columns:list，每列的数据名称，默认是名称
                rows：多少行，默认是1
                language：什么语言，可以填english，默认是中文
                path：输出excel的位置，有默认值
        @Ref  : 可以fake的数据类型有：https://mp.weixin.qq.com/s/xVwEjXu58WovgSi4ZTtVQw
        """
        # 可以选择英语
        language = 'en_US' if language.lower() == 'english' else 'zh_CN'
        fake = Faker(language)
        excel_dict = {}
        try:
            for column in simple_progress(columns, desc=f'columns'):
                excel_dict[column] = list()
                for _ in simple_progress(range(0, rows), desc='rows'):
                    excel_dict[column].append(eval('fake.{func}()'.format(func=column)))
        except AttributeError:
            print("输入列名有误，请检查columns的输入list值")
            print("详细参考: https://mp.weixin.gq.com/s/xVwEjXu58WovgSi4ZTtVQw")
            exit(1)
        # 用pandas，将模拟数据，写进excel里面
        res_excel_file = pd.ExcelWriter(str(Path(path).absolute()))
        res_data = pd.DataFrame(excel_dict)
        res_data = pandas_mem.reduce_pandas_mem_usage(res_data)
        res_data.to_excel(res_excel_file, index=False)
        # writer.save()
        res_excel_file.close()

    def merge2excel(self, dir_path, output_file, xlsxSuffix=".xlsx"):
        """
        :param dir_path: 存放excel文件的位置
        :param output_file: 输出合并后excel文件的位置
        :return: 没有返回值
        """
        if not output_file.endswith(xlsxSuffix):
            raise Exception(f'您自定义的输出文件名，不是以{xlsxSuffix}结尾的')
        file_path_dict = self.getfile(dir_path)  # excel文件所在的文件夹
        try:
            writer = pd.ExcelWriter(output_file)  # 合并后的excel名称
        except PermissionError:
            raise Exception(f'小可爱，你的输出文件，是不是上次打开了没关闭呀？这是你自己指定的输出文件名称：{output_file}')
        for file, path in file_path_dict.items():
            # 将所有excel sheet 全部加进来，然后利用报错结束 InvalidWorksheetName
            if file.endswith("xlsx"):
                for i in range(0, 100):
                    try:
                        df = pd.read_excel(path, sheet_name=i)
                        df.to_excel(writer, sheet_name=file.split('.')[0] + '_(' + str(i + 1) + ')', index=False)
                    # index 超出范围 - 结束
                    except ValueError:
                        break
                    except Exception as e:
                        raise Exception(f"文件名弧过长，请修改。")
            if file.endswith("csv"):
                df = pd.read_csv(path)
                df.to_excel(writer, sheet_name=file.split('.')[0], index=False)
        print(f'您指定的Excel文件已经合并完毕，合并后的文件名是{output_file}')
        writer.save()

    def getfile(self, dirpath):
        path = Path(dirpath)
        file_path_dict = {}
        for root, dirs, files in os.walk(dirpath):
            for file in files:
                if file.endswith("xlsx") or file.endswith("csv"):
                    file_path_dict[file] = (path / file)
        return file_path_dict

    def sheet2excel(self, file_path, output_path: str):
        # 先读取一次文件，获取sheet表的名称

        origin_excel = load_workbook(filename=file_path)  # 读取原excel文件
        origin_sheet_names = origin_excel.sheetnames  # 获取sheet的名称
        print(f'一共有{len(origin_sheet_names)}个sheet，名称分别为：{origin_sheet_names}')
        print('拆分开始')

        if len(origin_sheet_names) > 1:  # 如果sheetnames小于1，报错：该文件不需要拆分

            for j in tqdm(range(len(origin_sheet_names))):

                wb = load_workbook(filename=file_path)  # 再读取一次文件，由于每次删除后需要保存一次，所以不能与上一次一样
                sheet = wb[origin_sheet_names[j]]
                wb.copy_worksheet(sheet)

                new_filename = Path(output_path).joinpath(origin_sheet_names[j] + '.xlsx')  # 新建一个sheet命名的excel文件

                for i in tqdm(range(len(origin_sheet_names))):
                    sheet1 = wb[origin_sheet_names[i]]
                    wb.remove(sheet1)

                wb.save(filename=new_filename)

                # 由于使用copy_worksheet后，sheet表名有copy字段，这里做个调整

                new = load_workbook(filename=new_filename)
                news = new.active
                news.title = origin_sheet_names[j]
                new.save(filename=new_filename)
            print('拆分结束')
        else:
            raise Exception(f"你的文件只有一个sheet，难道还要拆分吗？我做不到啊~~~，你的文件名{file_path}")

    def merge2sheet(self, dir_path, output_sheet_name: str, output_excel_name):
        # 多个excel文件，合并到一个sheet里面
        for root, dirs, files in os.walk(dir_path):
            path = Path(dir_path)
            print(f'正在合并的文件有：{files}')
            print(f'合并后的文件名是：{output_excel_name}')
            print(f'合并后的sheet名是：{output_sheet_name}')
            df_list = []
            for file in files:
                if file.endswith("xlsx") or file.endswith("xls"):
                    excel_path = (path / file)
                    df_list.append(pd.read_excel(excel_path))
            res = pd.concat(df_list)
            res.to_excel(
                (path / (output_excel_name + '.xlsx')),
                sheet_name=output_sheet_name,
                index=False  # 不保留index
            )

            pass

    def find_excel_data(self, search_key, target_dir):
        """
        检索指定目录下的excel文件和过滤
        参数：
            search_key：检索的关键词
            target_dir：目标文件夹
        """
        search4file.find_excel_data(search_key, target_dir)

    def split_excel_by_column(self, filepath, column, worksheet_name):
        SplitExcel.split_excel_by_column(filepath, column, worksheet_name)

    def excel2pdf(self, excel_path, pdf_path, sheet_id):
        """
        https://blog.csdn.net/qq_57187936/article/details/125605967
        """
        abs_excel_path = str(Path(excel_path).absolute())
        input_excel_path_list1 = get_files(abs_excel_path, suffix='.xls')
        input_excel_path_list2 = get_files(abs_excel_path, suffix='.xlsx')
        input_excel_path_list1.extend(input_excel_path_list2)
        output_pdf_path = Path(pdf_path).absolute()
        for excel_file in input_excel_path_list1:
            with xw.App() as app:  # 下列来源：https://www.qiniu.com/qfans/qnso-57724345#comments
                app.visible = False
                # Initialize new excel workbook
                book = app.books.open(str(excel_file))
                sheet = book.sheets[sheet_id]
                # Construct path for pdf file
                pdf_path_name = os.path.join(str(output_pdf_path), Path(excel_file).stem + '.pdf')
                sheet.to_pdf(path=pdf_path_name, show=False)
    def count4page(self,input_path):
        """
        统计Excel文件打印的页数
        :author Cai-cy
        :param input_path:
        :return:
        """
        # 指定文件夹路径
        # 打开 Excel 应用程序
        excel = win32com.client.Dispatch("Excel.Application")

        # 遍历文件夹下的所有文件
        for file_name in os.listdir(input_path):
            # 判断文件是否是 Excel 文件
            if file_name.endswith(".xlsx") or file_name.endswith(".xls"):
                # 打开 Excel 文件
                file_path = os.path.join(input_path, file_name)
                workbook = excel.Workbooks.Open(file_path)
                # 获取 Excel 文件的打印页数
                page_count = workbook.ActiveSheet.PageSetup.Pages.Count
                # 输出 Excel 文件的打印页数
                print(f"{file_name}: {page_count}页")
                # 关闭 Excel 文件
                workbook.Close()

        # 关闭 Excel 应用程序
        excel.Quit()