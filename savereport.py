
import xlsxwriter
import os.path

class Savereport(object):
    """
    此类专门用于写入Excel
    """
    # __workbook用于存放excel文件的对象
    __workbook = None
    # __sheet用于存放excel文件中一张表格的对象，文件操作时主要操作该对象
    __sheets = {}

 

    def __init__(self,fliename,sheet_name):
        path=os.path.relpath(fliename)
        print("报告路径:"+path)
        self.__workbook = xlsxwriter.Workbook(path)
        self.addSheet(sheet_name)

        

    def addSheet(self, sheet_name):
        """
        添加单元格
        :param sheet_name: 单元格名称
        """
        self.__sheets[sheet_name] = self.__workbook.add_worksheet(sheet_name)


    def set_sheet_style(self,sheet_name):
        self.__sheets[sheet_name].set_column(0, 5, 20)
        bold=self.__workbook.add_format({'font_size':10, 'bold':1,'bg_color':'#FEFEFE', 'font_color':'#101010','align':'left', 'top':2,'left':2,'right':2,'bottom':2,'text_wrap': True, 'num_format':'yyyy-mm-dd' })
        self.__sheets[sheet_name].set_default_row(100)
        self.__sheets[sheet_name].set_row(0,20)
        return bold

    def close_excel(self):
        self.__workbook.close()

    def put_value_in_cell(self, row_index, col_index, value,bold, sheet_name):
        """
        把字符串填入表格的单元格
        :param value:要填入的值
        :param row_index:要填入值所在的行号
        :param col_index:要填入值所在的列号
        :param sheet_name 单元格名称
        """
        self.__sheets[sheet_name].write(row_index, col_index, value,bold)



    

