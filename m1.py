
from selenium.webdriver.support.ui import Select
from selenium import webdriver
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import Chrome
from selenium.webdriver.support import expected_conditions as EC

import time
import openpyxl
import sys
import pandas as pd
import winsound
import re
drive = Chrome()
path = 'chromedriver.exe'
bro = webdriver.Chrome(path)
print('done importing')


# add input here for the file name
file_name = 'BILL CHECKED.xlsx'
wb = openpyxl.load_workbook(file_name)
sheet = wb.active
sheet2 = wb.create_sheet()
sheet2.append(['Sr No', 'Discom', 'K No', 'Amount',
               'status', 'Website amount'])
wb.save(file_name)


def del_row(index, sheet=sheet):
    print('delete row command for index', index)
    sheet.delete_rows(index)
    wb.save(file_name)
    print('done deleting and saving')


def append_data(sheet=sheet2, appends=['Sr No', 'Discom', 'K No', 'Amount', 'status', 'Website amount']):
    print('append command for sheet', sheet, 'data', appends)
    sheet2.append(appends)
    wb.save(file_name)
    print('done appending and saving')

# for deleter of rows and colums use sheet.delete_row or col and give it any index and don't forget to save


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


def Handle(sr, aMOUNT, k_no, provider='DISCOM'):
    loader()
    time.sleep(2)
    try:
        body = bro.find_element_by_xpath(
            '''//*[@id="billBox"]/div[1]/div[3]/div''')
        print('body found')
        if body.is_displayed():

            bill = int(body.text.split('TOTALAMOUNT')[1])

            if aMOUNT == bill:
                print('amount same')
                bro.find_element_by_xpath(
                    '''//*[@id="cpBody_divPayment"]/div/div[5]''').click()
                print('click to pay1')
                bro.find_element_by_xpath('''//*[@id="cpBody_btn_PayBill"]''')
                # loading is pending
                bro.find_element_by_xpath(
                    '''//*[@id="paymentModeTypeNetBId"]''').click()

                # loading after pay2
                time.sleep(2)
                try:
                    if bro.find_element_by_xpath(
                            '''//*[@id="cpBody_btnBack"]''').is_displayed():
                        print('Back To SSo Present')
                        bro.find_element_by_xpath(
                            '''//*[@id="cpBody_btnBack"]''').click()
                        data = [sr, provider, k_no, aMOUNT, 'PAID', bill]
                        append_data(sheet2, data)

                except Exception as e:
                    print('Exception bcz of back to sso', e)
                    data = [sr, provider, k_no, aMOUNT, 'PAID', bill]
                    append_data(sheet2, data)

            else:
                print('amount not same')
                bro.find_element_by_xpath(
                    '''//*[@id="cpBody_btn_Reset"]''').click()
                data = [sr, provider, k_no, aMOUNT,
                        'Not Having same amount', bill]
                append_data(sheet2, data)
        else:
            print('not found body')
            data = [sr, provider, k_no, aMOUNT,
                    'data not found', 'not found bill']
            append_data(sheet2, data)

    except Exception as e:
        print('Exception from 39', e)
        time.sleep(1)
        button = bro.find_elements_by_tag_name('button')
        try:

            for k in button:
                if k.is_displayed():
                    print(k.text)
                    k.click()
                    data = [sr, provider, k_no, aMOUNT,
                            'Not Found', 'Not Found']
                    append_data(sheet2, data)
        except Exception as e:
            print('Exception from line 91', e)
            bro.find_element_by_xpath(
                '''//*[@id="cpBody_btnBack"]''').click()
            data = [sr, provider, k_no, aMOUNT, 'PAID', 'data not found ']
            append_data(sheet2, data)
            dashboard()


def set_network(offline=False, latency=5, download_throughput=500*1024, upload_throughput=500*1024):
    bro.set_network_conditions(
        offline=offline,
        latency=latency,  # additional latency (ms)
        download_throughput=download_throughput,  # maximal throughput
        upload_throughput=upload_throughput)


def start():
    set_network(latency=0, download_throughput=500 *
                1024, upload_throughput=500*1024)
    print('net set')
    bro.get("https://sso.rajasthan.gov.in/dashboard")
    bro.maximize_window()
    bro.find_element_by_xpath(
        '''//*[@id="cpBody_txt_Data1"]''').send_keys('seshmoub')
    bro.find_element_by_xpath(
        '''//*[@id="cpBody_txt_Data2"]''').send_keys('shubham@1234')
    winsound.Beep(1223, 100)
    time.sleep(12)

    try:
        loader()
        bro.find_element_by_xpath(
            '''//*[@id="cpBody_btn_LDAPLogin"]''').click()

        if bro.find_element_by_xpath('''//*[@id="cpBody_txt_Data2"]''').is_displayed():
            bro.find_element_by_xpath(
                '''//*[@id="cpBody_txt_Data2"]''').send_keys('shubham@1234')
            bro.find_element_by_xpath(
                '''//*[@id="cpBody_cbx_newsession"]''').click()
            bro.find_element_by_xpath(
                '''//*[@id="cpBody_btn_LDAPLogin"]''').click()
    except Exception as e:
        print('Exception raised on line no 164', e)
        dashboard()
    loader()


def dashboard():
    # DASHBOARD
    bro.find_element_by_xpath('''//*[@id="billapp"]/a/span''').click()
    bro.find_element_by_xpath(
        '''//*[@id="cpBody_dlBillPayment_lnkBills_1"]/div/div/div[2]/div[1]/img''').click()
    loader()


def main_menu(sr, k_no, amount, provider='DISCOM'):
    # MAINMENU
    Select(bro.find_element_by_xpath(
        '''//*[@id="cpBody_ddl_List"]''')).select_by_visible_text('DISCOM')
    loader()
    if bro.find_element_by_xpath('''//*[@id="cpBody_txt_Search"]''').get_attribute('value') != '':
        bro.find_element_by_xpath('''//*[@id="cpBody_txt_Search"]''').clear()
    bro.find_element_by_xpath(
        '''//*[@id="cpBody_txt_Search"]''').send_keys(str(k_no))
    bro.find_element_by_xpath('''//*[@id="cpBody_btn_Submit"]''').click()
    loader()
    time.sleep(1)
    Handle(sr, amount, k_no, provider='DISCOM')
    time.sleep(1)


if __name__ == '__main__':
    start()
    dashboard()
    file = pd.read_excel('BILL CHECKED.xlsx')
    K_no = list(file.get('Discom'))
    print(K_no)
    print('--------')
    Amount = list(file.get('Amount'))

    for k in range(len(K_no)):
        print(K_no[k], Amount[k])
        #  main_menu(sr, k_no, AMOUNT, provider='DISCOM'):
        main_menu(k+1, K_no[k], int(Amount[k]))
    bro.quit()
    print('Congrats TEst Passed')
