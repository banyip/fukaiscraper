# coding=utf-8
# 封装selenium常用操作

from selenium import webdriver
from selenium.webdriver.ie.service import Service
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support.select import Select



class seleniumUtil():
    def __init__(self,ouput_folder_path='output',explorer='chrome'):
        options = webdriver.ChromeOptions()
        #设置为0表示禁止弹出窗口，                     #设置文件下载路径
        prefs = { 'profile.default_content_settings.popups':0 ,"profile.content_settings.exceptions.automatic_downloads.*.setting":1}
        #prefs.append({'download.default_directory': output_folder_path})
        options.add_experimental_option('prefs',prefs)
        options.add_argument("--no-sandbox")
        self.drive=webdriver.Chrome(chrome_options=options)


    def openUrl(self,url):
        try:      
            self.drive.implicitly_wait(20)
            self.drive.maximize_window()
            self.drive.get(url)
        except Exception as e:
            # print(e)
            print(traceback.print_exc())
            raise e


    def input(self,strElement,value):
        input=None
        if strElement[:1]=='#':
            input=self.drive.find_element_by_id(strElement[1:])
        elif strElement[:1]=='/':
            input=self.drive.find_element_by_xpath(strElement)
        elif strElement[:1]=='$':
            input=self.drive.find_element_by_name(strElement[1:])
        if input is not None:
            input.click()
            input.clear()
            input.send_keys(value)

    def clickElement(self,strElement,tagname=None,clickType=None):
        input = None
        while True:
            try:
                if tagname is None:
                    if strElement[:1]=='#':
                        input=self.drive.find_element_by_id(strElement[1:])
                    elif strElement[:1]=='/':
                        input=self.drive.find_element_by_xpath(strElement)
                    elif strElement[:1]=='$':
                        input=self.drive.find_element_by_name(strElement[1:])
                    elif strElement[:1]=='!':
                        input=self.drive.find_element_by_link_text(strElement[1:])
                    else:
                        buttons=self.drive.find_elements_by_tag_name('button')
                        for button in buttons:
                            print(button.text)
                            if button.text==strElement:
                                input=button
                                break
                else:
                    buttons=self.drive.find_elements_by_tag_name(tagname)
                    for button in buttons:
                        #print(button.text)
                        if button.text==strElement:
                            input=button

                break
            except Exception as e:
                print(e)
                time.sleep(2)
                continue    
        if input is not None:
            if clickType is None:
                self.drive.execute_script("$(arguments[0]).click()",input)
            else:
                input.click()
            



