import xlrd
import logging





logging.basicConfig(format='%(asctime)s - %(pathname)s[line:%(lineno)d] '
                           '- %(levelname)s: %(message)s',level=logging.DEBUG)

class Readconfig(object):

    def __init__(self,filepath,testsystem):
        self.filepath=filepath
        self.testsystem=testsystem
        self.all_sheets=None
        self.configs=None
        self._read_file()
        self.parse_configs()



    
    def _read_file(self):
        print('执行')
        all_sheets={}
        try:
            workbook=xlrd.open_workbook(self.filepath)
        except Exception as e:
            logging.error("open testcase file fail, please check!"+str(e))
            return False
        
        sheet_names = workbook.sheet_names()
        print(sheet_names)
        for sheet_name in sheet_names:
            if '国网配置' in sheet_name:
                guowang_configs=workbook.sheet_by_name('国网配置')
                all_sheets['guowang_configs']=guowang_configs
            elif 'gcloud配置' in sheet_name:
                gcloud_configs=workbook.sheet_by_name('gcloud配置')
                all_sheets['gcloud_configs']=gcloud_configs

        self.all_sheets=all_sheets
        print(all_sheets)
        return all_sheets


    def parse_configs(self):
        configs_dict={}
        sheet=None
        print('系统类型：'+self.testsystem)
        if self.testsystem=='1':             #根据你输入的系统代表值去获取对应的系统配置
            sheet=self.all_sheets['guowang_configs']
        elif self.testsystem=='2':
            sheet=self.all_sheets['gcloud_configs']
        if not sheet:
            logging.error('测试系统选择不在范围内！')
            return None

        rows=sheet.get_rows()
        for row in list(rows)[1:]:
            key_content = row[0].value
            value_content = row[1].value
            if key_content:
                if key_content=='system_url':
                    configs_dict['system_url']=value_content
                elif key_content=='username':
                    configs_dict['username']=value_content
                elif key_content=='password':
                    configs_dict['password']=value_content
                elif key_content=='uername_element_xpath':
                    configs_dict['uername_element_xpath']=value_content
                elif key_content=='password_element_xpath':
                    configs_dict['password_element_xpath']=value_content
                elif key_content=='commitbutton_xpath':
                    configs_dict['commitbutton_xpath']=value_content
                elif key_content=='clickcount':
                    configs_dict['clickcount']=value_content
        self.configs=configs_dict
