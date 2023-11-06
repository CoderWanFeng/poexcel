import xlwings as xw

from poexcel.core.ExcelType import MainExcel


def centimeters2points(cm: float or int) -> float:
    """
    用于行高和页边距等的厘米转磅,换算公式见下:
    https://learn.microsoft.com/zh-cn/office/vba/api/excel.application.centimeterstopoints
    Args:
        cm:厘米数

    Returns:磅值

    """
    return cm / 0.035


def get_sheet_index(excel_file: str, sheet_list: str or int or [str] or [int] or None, return_name: bool = True,
                    sort: bool = False) -> list:
    """
        根据从1开始的工作表序号或工作表名,从核查并获取sheet的索引号(从0开始)或工作表名
    Args:
       excel_file:必填,指定excel文件路径
       sheet_list:必填,当为 str 视为工作表名, 当为 int 时,视为工作表序号(从1开始),当为-1 或 None 时,为所有工作表
       return_name:可选,默认为True,返回工作表名,当为False时,返回序号号(从0开始)
       sort:可选,默认为False,根据传入的顺序返回,当为True时,按工作表的顺序返回

    Returns:从工作表名或序列号组成的列表

    """
    with xw.App() as app:
        app.visible = False
        book = app.books.open(str(excel_file))
        sheet_names: list = book.sheet_names  # 所有工作表名
    r = []  # 要返回的列表
    if not isinstance(sheet_list, list):  # 如果不为列表转换为列表
        if sheet_list == -1 or sheet_list is None:  # 如果是-1或None 返回所有工作表名
            return sheet_names
        else:
            sheet_list = [sheet_list]
    for i in range(len(sheet_list)):
        if isinstance(sheet_list[i], int):  # 如果为int,且在正确范围内,直接减1取值
            if sheet_list[i] - 1 not in range(len(sheet_names)):
                raise IndexError(f"序号{sheet_list[i]}超过工作簿序号1-{len(sheet_names)}范围.")
            n = sheet_list[i] - 1
        elif isinstance(sheet_list[i], str):  # 如果为str,寻找序列号
            if sheet_list[i] not in sheet_names:
                raise IndexError(f"未在工作簿中找到工作表名{sheet_list[i]}.")
            n = sheet_names.index(sheet_list[i])
        else:  # 即不为int 也不是str 报错
            raise TypeError(f"{sheet_list[i]}应为str或int,实际为{type(sheet_list[i])}")
        if n in r:  # 提示但不报错
            print(f'参数中出现了重复工作表名{sheet_names[n]}')
        else:
            r.append(n)

    if sort:  # 需要排序
        r.sort()
    if return_name:  # 需要返回工作表名
        return [sheet_names[i] for i in r]
    else:
        return r


class NewMainExcel(MainExcel):

    def page_setup(self, excel_file: str, new_file: str or None = None,
                   sheet_list: str or int or [str] or [int] or None = None,
                   paper_size: str or None = None,
                   Zoom: int or False or None = None, FitToPagesTall: int or False or None = None,
                   FitToPagesWide: int or False or None = None,
                   BlackAndWhite: bool or None = None, Margin: [int] or None = None, Orientation: bool or None = None,
                   **kwargs):

        # 对工作表列表处理并转换
        sheet_list = get_sheet_index(excel_file, sheet_list)

        # 几种常用的纸张尺寸
        paper_size_dict = dict(A3=xw.constants.PaperSize.xlPaperA3, A4=xw.constants.PaperSize.xlPaperA4,
                               A5=xw.constants.PaperSize.xlPaperA5, B4=xw.constants.PaperSize.xlPaperB4,
                               B5=xw.constants.PaperSize.xlPaperB5)

        # 处理纸张尺寸
        if paper_size:
            if paper_size.upper() not in paper_size_dict:
                raise ValueError(f"paper_size参数的值{paper_size}不是常用值,请使用常用值或使用PaperSize参数")
            paper_size = paper_size_dict[paper_size.upper()]
        # 处理纸张方向
        if Orientation is None:
            pass
        else:
            if Orientation is True:
                Orientation = 2  # 横向打印
            elif Orientation is False:
                Orientation = 1  # 纵向打印
            else:
                raise ValueError(f"Orientation 参数的值{Orientation}不是bool")
        # 缩放
        if Zoom:
            if not isinstance(Zoom, int):
                raise ValueError(f"Zoom参数的值{Zoom}不是int")
            if FitToPagesWide or FitToPagesTall:
                raise ValueError(f"当使用Zoom参数时,不可同时使用FitToPagesWide或FitToPagesTall参数")
        else:
            if FitToPagesWide and not isinstance(FitToPagesWide, int):  # 设置打印工作表时将宽度缩放到的页数
                raise ValueError(f"FitToPagesWide 参数的值{FitToPagesWide}不是int")
            if FitToPagesTall and not isinstance(FitToPagesTall, int):  # 设置打印工作表时将高度缩放到的页数
                raise ValueError(f"FitToPagesTall 参数的值{FitToPagesTall}不是int")
        # 黑白打印
        if BlackAndWhite is None:
            pass
        else:
            if not isinstance(BlackAndWhite, bool):
                raise ValueError(f"BlackAndWhite 参数的值{BlackAndWhite}不是bool")
        # 页边距
        if Margin is None:
            pass
        else:
            if not (isinstance(Margin, list) and len(Margin) == 4):
                raise ValueError(f"Margin 参数的值{Margin}应是长度为4的list")
            if not all(map(lambda x: isinstance(x, int) or isinstance(x, float), Margin)):
                raise ValueError(f"Margin 参数的值{Margin}均应是int或float")

        with xw.App() as app:
            app.visible = False
            book = app.books.open(str(excel_file))

            for sheet_id in sheet_list:
                sheet = book.sheets[sheet_id]

                # 纸张尺寸
                if paper_size:
                    sheet.api.PageSetup.PaperSize = paper_size
                # 纸张方向
                if Orientation:
                    sheet.api.PageSetup.Orientation = Orientation
                # 缩放
                if not Zoom is None:
                    sheet.api.PageSetup.Zoom = Zoom
                if not Zoom:  # 如果Zoom为None或False时
                    if FitToPagesWide or FitToPagesTall:  # 如果这两个缩放有任意一个有值
                        sheet.api.PageSetup.Zoom = False  # 如果有Zoom那么宽度和高度缩放不生效,要先重置
                    if not FitToPagesWide is None:  # 宽度缩放
                        sheet.api.PageSetup.FitToPagesWide = FitToPagesWide
                    if not FitToPagesTall is None:  # 高度缩放
                        sheet.api.PageSetup.FitToPagesTall = FitToPagesTall
                # 黑白打印
                if BlackAndWhite is None:
                    pass
                else:
                    sheet.api.PageSetup.BlackAndWhite = BlackAndWhite
                # 页边距
                if Margin:
                    sheet.api.PageSetup.TopMargin = centimeters2points(Margin[0])  # 上
                    sheet.api.PageSetup.BottomMargin = centimeters2points(Margin[1])  # 下
                    sheet.api.PageSetup.LeftMargin = centimeters2points(Margin[2])  # 左
                    sheet.api.PageSetup.RightMargin = centimeters2points(Margin[3])  # 右

                # 不定长参数
                # 不做处理,只做错误捕捉
                try:
                    if kwargs:
                        for k, v in kwargs.items():
                            setattr(sheet.api.PageSetup, k, v)
                except Exception as err:
                    print(err)
                    raise ValueError(
                        '请参照 https://learn.microsoft.com/zh-cn/office/vba/api/excel.pagesetup 使用其他参数')
            # 保存文件
            if new_file:
                book.save(new_file)
            else:
                book.save()


mainExcel = NewMainExcel()


def page_setup(excel_file: str, new_file: str or None = None,
               sheet_list: str or int or [str] or [int] or None = None,
               paper_size: str or None = None,
               Zoom: int or False or None = None, FitToPagesTall: int or False or None = None,
               FitToPagesWide: int or False or None = None,
               BlackAndWhite: bool or None = None, Margin: [int] or None = None, Orientation: bool or None = None,
               **kwargs):
    """
    批量修改工作表的页面设置,仅对一些常用设置做了判断处理,更多的属性可以到Excel vba 官方文档查询,用法见下面的示例.
    https://learn.microsoft.com/zh-cn/office/vba/api/excel.pagesetup
    paper_size将部分常用纸张尺寸设置了别名,更多纸张尺寸可以通过不定参PaperSize指定.
    除此以外的参数均沿用VBA的大驼峰命名,便于与不定参统一.
    Args:
        excel_file:必选,Excel文件的路径,仅支持单个Excel工作簿
        new_file:可选,修改后的Excel的保存路径,默认值为None,覆盖保存
        sheet_list:可选,当为 str 视为工作表名, 当为 int 时,视为工作表序号(从1开始),当为-1 或 None 时,为所有工作表.默认为None
        paper_size:可选,纸张尺寸的别名,可选值包括A3,A4,A5,B4,B5,不区分大小写,默认为None,不修改.
        Zoom:可选,缩放比例,当为int时,认定为缩放的百分比数,例如100即缩放100%.默认为None,不修改.也可为False,重置缩放.当Zoom不为False时,不能指定FitToPagesTall及FitToPagesWide
        FitToPagesTall:可选,设置打印工作表时根据页高缩放到的页数,只能是整数,当为False时为自动.默认为None,不修改.当指定了该参数且不为False时,会自动将Zoom修改为False.
        FitToPagesWide:可选,设置打印工作表时根据页宽缩放到的页数,只能是整数,当为False时为自动.默认为None,不修改.当指定了该参数且不为False时,会自动将Zoom修改为False.
        BlackAndWhite:可选,是否单色打印,当为False时关闭单色打印,当为True时开启.默认为None,不修改.
        Margin:可选,四个方向页边距组成的列表,格式为[上,下,左,右],单位为厘米.如仅需修改若干个或页眉页脚,可以通过不定参修改.默认为None,不修改.
        Orientation:可选,纸张方向,当为True时横向,当为False时纵向,默认为None,不修改.
        **kwargs:其他VBA中PageSetup对象的属性,参数名为VBA中的属性名,值见VBA官方文档,当采用常量时,可在xlwings\constants.py 中查找

    Returns:None

    """

    mainExcel.page_setup(excel_file, new_file, sheet_list, paper_size,
                         Zoom, FitToPagesTall, FitToPagesWide, BlackAndWhite, Margin, Orientation, **kwargs)


if __name__ == '__main__':
    # nm = NewMainExcel()
    excel_files = r".\excel_files\page_setup.xlsx" # 测试文件
    new_files = r".\excel_files\page_setup 修改{}.xlsx"

    # 以从1开始的工作表序号指定的单个工作表 (第1个)
    # 缩放调整为80 纸张尺寸为B5(通过可选参数paper_size) 其他不变
    page_setup(excel_files, new_files.format(1), sheet_list=1, Zoom=80, paper_size='b5')
    # 以从工作表名称指定的单个工作表 (第2个)
    # 缩放调整为70 纸张尺寸为B4(通过不定长参数PaperSize,值采用xw的常量) 页边距修改为 2,2,2,2cm 其他不变
    page_setup(excel_files, new_files.format(2), sheet_list="工作表2", Zoom=70,
               PaperSize=xw.constants.PaperSize.xlPaperB4, Margin=[2, 2, 2, 2])
    # 以从工作表名称和序号混合的方式指定的多个工作表 (第1个和第3个)
    # 缩放调整为宽度1页 纸张尺寸为B4(通过不定长参数PaperSize,值采用枚举值)  页眉高度修改为4cm 其他不变
    page_setup(excel_files, new_files.format(3), sheet_list=["工作表3", 1], FitToPagesWide=1, PaperSize=12,
               HeaderMargin=centimeters2points(4))
    # 以默认值的方式修改所有工作表
    # 缩放调整为高度1页 宽度为自动 纸张尺寸为B4(通过可选参数paper_size) 纸张横向 单色打印  其他不变
    page_setup(excel_files, new_files.format(4), FitToPagesTall=1, FitToPagesWide=False, paper_size='b4',
               Orientation=True,BlackAndWhite=True)
