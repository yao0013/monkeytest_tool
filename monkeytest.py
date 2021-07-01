# -*- coding: utf-8 -*-
import time
import datetime
import pytz
from selenium import webdriver
import pyautogui as page
import random
from browsermobproxy import Server, server
import os
from readconfig import Readconfig
from savereport import Savereport
from datetime import datetime
import logging
import re
from responsecode_dict import *


logging.basicConfig(format='%(asctime)s - %(pathname)s[line:%(lineno)d] '
                           '- %(levelname)s: %(message)s',level=logging.DEBUG)
#server = Server(r'e:\gcloudworkplace\monkeytest\browsermobproxy\browsermob-proxy-2.1.4\bin\browsermob-proxy.bat')
server = Server(r'externaltool\browsermob-proxy-2.1.4\bin\browsermob-proxy.bat')
class Monkeytest(object):
    def __init__(self,configs):
        self.url=configs['system_url']
        self.username=configs['username']
        self.password=configs['password']
        self.uername_element_xpath=configs['uername_element_xpath']
        self.password_element_xpath=configs['password_element_xpath']
        self.commitbutton_xpath=configs['commitbutton_xpath']
  
    def get_proxy(self):
        server.start()
        proxy = server.create_proxy()
        return proxy
    
    def get_driver(self,proxy):
        options = webdriver.ChromeOptions()
        options.add_argument('--ignore-certificate-errors')
        options.add_argument('--proxy-server={0}'.format(proxy.proxy))
        #driver = webdriver.Chrome(executable_path='e:\gcloudworkplace\monkeytest\\v90chromedriver\chromedriver.exe',chrome_options=options)
        driver = webdriver.Chrome(executable_path='externaltool\\v90chromedriver\chromedriver.exe',chrome_options=options)
        driver.maximize_window()
        return driver

    def login(self,driver):
        driver.get(self.url)
        time.sleep(2)
        user = driver.find_element_by_xpath(self.uername_element_xpath)   #定位用户名控件
        user.send_keys(self.username)
        password = driver.find_element_by_xpath(self.password_element_xpath)  #定位密码控件
        try:
            password.send_keys(str(int(self.password)))
        except:
            password.send_keys(self.password)
        button = driver.find_element_by_xpath(self.commitbutton_xpath)   #定位确认按钮控件
        button.click()
        time.sleep(4)
    
    def autoclick(self,driver,count):
        for i in range(count):
            i = i+1
            x = random.randint(0, driver.get_window_size()["width"])
            y = random.randint(0, driver.get_window_size()["height"])
           #限制鼠标点击的区域，防止鼠标点击到系统退出登录按钮或者浏览器关闭按钮等
            if (x>1788):
                if (y<=175):
                    y=176
                if (y>=993):
                    y=993
            if (x<312):
                if(y<95):
                    y=95
            if (y<114):
                y=114
            if (y>1012):
                y=1012
            
            #print(x,y,i)
            page.PAUSE = 0.1
            page.click(x,y)

def set_sheet_title(Savereport,bold,sheet_name):
    sheet_title=['url','method','request','response','status_code','初步诊断']
    for i in range(6):
        Savereport.put_value_in_cell(0, i, sheet_title[i],bold, sheet_name)


def parse_request(request,i,Savereport,bold,sheet_name):
    request_str=str(request)
    if '\'postData\':'  in request_str:
        try:
            Savereport.put_value_in_cell(i,2,str(request['postData']),bold,sheet_name)
        except:
            logging.warning('请求没有postData属性')
    else:
        Savereport.put_value_in_cell(i,2,str(request),bold,sheet_name) 

def parse_response(response,i,Savereport,bold,sheet_name):
    response_str=str(response)
    if '\'content\':'  in response_str and '\'text\':'  in response_str:
        Savereport.put_value_in_cell(i,3,str(response['content']['text']),bold,sheet_name) 
        #print('响应的text:'+str(response['content']['text']))
    else:
        Savereport.put_value_in_cell(i,3,str(response),bold,sheet_name)
        logging.warning('响应没有text属性')



def parse_response_code(response,_response_statuscode,i,Savereport,bold,sheet_name):
    Found_out_forcode=False
    response_str=str(response)
    response_code_final=None
    if '\'content\':'  in response_str and '\'text\':'  in response_str and '\"code\":' in response_str:
        _response_code=re.findall(r"code\":(.+?),",str(response['content']['text']))
        Savereport.put_value_in_cell(i,4,str(_response_code[0]),bold,sheet_name)
        response_code_final=str(_response_code[0])
    else:
        Savereport.put_value_in_cell(i,4,str(_response_statuscode),bold,sheet_name)
        response_code_final=str(_response_statuscode)
        logging.warning('响应没有code属性')
    for key in responsecode_dict.keys():
        if key==response_code_final:
            Savereport.put_value_in_cell(i,5,str(responsecode_dict[key]),bold,sheet_name)
            Found_out_forcode=True
    if not Found_out_forcode:
        Savereport.put_value_in_cell(i,5,'未知原因',bold,sheet_name)
        

    


def parse_result(result,Savereport,bold,sheet_name):
    i=1
    for entry in result['log']['entries']:
        _url = entry['request']['url']
        _response_statuscode=entry['response']['status']
        if str(_response_statuscode) != '200':
            Savereport.put_value_in_cell(i,0,_url,bold,sheet_name)
            try:
                Savereport.put_value_in_cell(i,1,entry['request']['method'],bold,sheet_name)
            except:
                logging.warning('请求没有method属性')
            try:
                parse_request(entry['request'],i,Savereport,bold,sheet_name)
            except:
                logging.warning('请求没有request属性')
            try:
                parse_response(entry['response'],i,Savereport,bold,sheet_name)
            except:
                logging.warning('请求没有response属性')
            parse_response_code(entry['response'],_response_statuscode,i,Savereport,bold,sheet_name)
            i=i+1
    


