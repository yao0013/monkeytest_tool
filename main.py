from monkeytest import *


if __name__=="__main__":
    file_path = os.path.abspath("monkeyconfig.xls")
    systemkind=input("请输入你要测试的系统(其中1代表国网系统，2代表gcloud系统):")    #手动输入你要测试的系统代表值
    configs=Readconfig(file_path,systemkind).configs
    monkeytest=Monkeytest(configs)
    proxy=monkeytest.get_proxy()
    driver=monkeytest.get_driver(proxy)
    proxy.new_har("douyin", options={'captureHeaders': True, 'captureContent': True})
    monkeytest.login(driver)
    monkeytest.autoclick(driver,int(configs['clickcount']))
    result = proxy.har
    now_time = datetime.now(tz=pytz.timezone("Asia/Shanghai")).strftime("%Y-%m-%d_%H-%M-%S")
    
    if systemkind=='1':    #根据你输入的系统代表值去定义测试报告中sheet名称    
        sheet_name='国网系统'
    elif systemkind=='2':
        sheet_name='gcloud系统'
    else:
        sheet_name='其他'
    reportfile_path=str('report/')
    report=Savereport(reportfile_path+now_time+ ".xlsX",sheet_name)
    bold=report.set_sheet_style(sheet_name)   
    set_sheet_title(report,bold,sheet_name)

    parse_result(result,report,bold,sheet_name)
    report.close_excel()         
    server.stop
    print("测试完毕")