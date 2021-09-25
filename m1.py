       
from selenium.webdriver.support.ui import Select
from selenium import webdriver
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import Chrome

import time
import sys
import pandas as pd
import winsound
import re
drive = Chrome()
path = 'chromedriver.exe'
bro = webdriver.Chrome(path)
print('done importing')

def loader():
    while True:
        if bro.find_element_by_xpath('''//*[@id="UpdateProgress"]/div/div''').is_displayed():
            print('Loading')
            time.sleep(1)
            continue
        return 
def wait(time_in_sec):
    bro.set_script_timeout(time_in_sec)
    return 

def Handle(AMOUNT):
    loader()
    time.sleep(2)
    try:
        body=bro.find_element_by_xpath('''//*[@id="billBox"]/div[1]/div[3]/div''')
        print('body found')
        if body.is_displayed():
            print('body displayed')
            bill=int(body.text.split('TOTALAMOUNT')[1])
            print('got bill amount')
            if AMOUNT == bill :
                print('amount same')
                bro.find_element_by_xpath('''//*[@id="cpBody_divPayment"]/div/div[5]''').click()
                print('pending second pay button')
            else:
                print('amount not same')
                bro.find_element_by_xpath('''//*[@id="cpBody_btn_Reset"]''').click()
        else:
            print('not found body')
            
    except Exception as e:
        print('Exception from 39',e)
        time.sleep(1)
        button=bro.find_elements_by_tag_name('button')
        for k in button:
            if k.is_displayed():
                print(k.text)
                k.click()





def set_network(offline=False,latency=5,download_throughput=500*1024,upload_throughput=500*1024):
    bro.set_network_conditions(
            offline=offline,
            latency=latency,  # additional latency (ms)
            download_throughput=download_throughput,  # maximal throughput
            upload_throughput=upload_throughput)


def start():
    set_network(latency=10,download_throughput=50000,upload_throughput=50000)
    print('net set')
    bro.get("https://sso.rajasthan.gov.in/dashboard")
    bro.maximize_window()
    bro.find_element_by_xpath('''//*[@id="cpBody_txt_Data1"]''').send_keys('seshmoub')
    bro.find_element_by_xpath('''//*[@id="cpBody_txt_Data2"]''').send_keys('shubham@1234')
    winsound.Beep(1223,100)
    time.sleep(12)
    bro.find_element_by_xpath('''//*[@id="cpBody_btn_LDAPLogin"]''').click()
    loader()
    if bro.find_element_by_xpath('''//*[@id="cpBody_txt_Data2"]''').is_displayed():
        bro.find_element_by_xpath('''//*[@id="cpBody_txt_Data2"]''').send_keys('shubham@1234')
        bro.find_element_by_xpath('''//*[@id="cpBody_cbx_newsession"]''').click()
        bro.find_element_by_xpath('''//*[@id="cpBody_btn_LDAPLogin"]''').click()
    loader()
def dashboard():
    # DASHBOARD
    bro.find_element_by_xpath('''//*[@id="billapp"]/a/span''').click()
    bro.find_element_by_xpath('''//*[@id="cpBody_dlBillPayment_lnkBills_1"]/div/div/div[2]/div[1]/img''').click()
    loader()
def main_menu(k_no,AMOUNT,provider='DISCOM'):
    # MAINMENU
    Select(bro.find_element_by_xpath('''//*[@id="cpBody_ddl_List"]''')).select_by_visible_text('DISCOM')
    loader()
    if bro.find_element_by_xpath('''//*[@id="cpBody_txt_Search"]''').get_attribute('value') != '':
        bro.find_element_by_xpath('''//*[@id="cpBody_txt_Search"]''').clear()
    bro.find_element_by_xpath('''//*[@id="cpBody_txt_Search"]''').send_keys(str(k_no))
    bro.find_element_by_xpath('''//*[@id="cpBody_btn_Submit"]''').click()
    loader()
    time.sleep(1)
    Handle(AMOUNT)
    time.sleep(1)
    
    


if __name__=='__main__':
    start()
    dashboard()
    file=pd.read_excel('BILL CHECKED.xlsx')
    K_no=list(file.get('Discom'))
    print(K_no)
    print('--------')
    Amount=list(file.get('Amount'))
    print(Amount)
    main_menu(320242001721,6893)
    for k in range(len(K_no)):
        print(K_no[k],Amount[k])
        main_menu(K_no[k],int(Amount[k]))
    print('Congrats TEst Passed')
    

    

    