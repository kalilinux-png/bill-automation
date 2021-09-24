
# try using bro.switch_to_active_elements()
# use id to check for similarity
# screenshot of cache

print('Starting the Program......')
# file_name=input('Please Enter The Excel File Name In Which Data is \n make sure the file is in same folder as the module1 \nfor more details kindly contact the developer ')

# username=input("Sir,Madam Please Enter Your UserName")
       
# password=input("Sir,Madam Please Enter Your LoginPassword")
print('importing files')
        
from selenium.webdriver.support.ui import Select
from selenium import webdriver
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import Chrome
import time
import sys
import pandas as pd
import winsound
drive = Chrome()
path = 'chromedriver.exe'
bro = webdriver.Chrome(path)
print('done importing')


def login(name, password):
    '''This Function is used to login with a name and password'''

    # bro.add_cookie({'name' : 'SSOAUTH', 'value' : '3F473ADCA829A8A4B7F97C65446EE2250F109ADAD2387C8FB0BA4EC503A4CCF7033A1B300D46CF7BA5D2301E90F57EAE39F0C2D11638D35FD93C9DE1DD25CB46D81FEA04F2F16423B7300F1FF82E6CB1BCE93C5CB498413D426553E83FE68AA54B1EED2D08140DC12A22208784F6090757D3704C609085B5AEA7C246046FE9034B7B1E91DB5B36401A4A6D1E1217B9FDF4788146C6420D1B9C4EC8777210570331FCC85BB909C31439A394B48A22BC146AEF64B3','domain':'sso.rajasthan.gov.in'})

    bro.get("https://sso.rajasthan.gov.in/dashboard")

    login_name = bro.find_element_by_xpath('''//*[@id="cpBody_txt_Data1"]''')

    password_field = bro.find_element_by_xpath(
        '''//*[@id="cpBody_txt_Data2"]''')

    login_button = bro.find_element_by_xpath(
        '''//*[@id="cpBody_btn_LDAPLogin"]''')

    login_name.send_keys(username)

    password_field.send_keys(password)

    
    winsound.Beep(1500, 1000)
    winsound.Beep(8392, 500)

    sys.stdout.write('Sir Waiting for 15 Sec Please Fill The Recaptcha')

    bro.find_element_by_xpath(
        '''//*[@id="cpBody_ssoCaptcha_imgCaptcha"]''').screenshot('Captch.png')

    time.sleep(15)

    login_button.click()

    try:

        # click on bill button
        bro.find_element_by_xpath('''//*[@id="billapp"]/a/span''').click()

        # click on bulb button
        bro.find_element_by_xpath(
            '''//*[@id="cpBody_dlBillPayment_lnkBills_1"]/div/div/div[2]/div[1]/img''').click()

    except Exception as e:

        print(e)
        try:
            alerts = bro.switch_to_active_element()
            time.sleep(2)
            if alerts.is_displayed():
                print('WARNING data not found:', alerts.text)
        except Exception as e:

            # password field is removed from here
            password_field = bro.find_element_by_xpath(
                '''//*[@id="cpBody_txt_Data2"]''')
            if password_field.is_displayed():
                password_field.send_keys(password)
                # check box
                bro.find_element_by_xpath(
                    '''//*[@id="cpBody_cbx_newsession"]''').click()
                # login button
            login_button = bro.find_element_by_xpath(
                '''//*[@id="cpBody_btn_LDAPLogin"]''')
            login_button.click()
            sys.stdout.write('Login Succesfull\n')
            return


def dashboard():
    print('welcome to dashboard')
    '''This is used for navigation in dashboard'''
    # pay bill button
    try:
        bro.find_element_by_xpath('''//*[@id="billapp"]/a/span''').click()
    # bulb button
        bro.find_element_by_xpath(
            '''//*[@id="cpBody_dlBillPayment_lnkBills_1"]/div/div/div[2]/div[1]/img''').click()
        return
    except Exception as e:
        print(e)
        time.sleep(2)
        dashboard()





def main_menu(row, K_no, Amount, provider='DISCOM'):
    '''This Function is used to fill required detail in the main billing area'''

    bar2 = Select(bro.find_element_by_xpath('''//*[@id="cpBody_ddl_List"]'''))
    # print(bar2)
    ''' here we need to make sure that we get the right provider make sure bar.select_by_name'''

    bar2.select_by_visible_text(provider)

    # below line is added because sometimes it loads
    while True:
        if bro.find_element_by_xpath('''//*[@id="UpdateProgress"]/div/div''').is_displayed():
            sys.stdout.write(
                f'Loading Please Wait..... Breath In Breath Out\n')
            time.sleep(1)
        else:
            break

    # Not Sure Why sleep here
    time.sleep(2)
    # K_no is send here
    bro.find_element_by_xpath('''//*[@id="cpBody_txt_Search"]''').send_keys(K_no)
    # 3 sec wait sometimes cause error
    time.sleep(2)
    # click submit button
    bro.find_element_by_xpath('''//*[@id="cpBody_btn_Submit"]''').click()

    try:
        # loader
        while True:
            if bro.find_element_by_xpath('''//*[@id="UpdateProgress"]/div/div''').is_displayed():
                print('Loading...')
            else:
                break
        time.sleep(3)
        # alert message if
        print(bro.switch_to_alert().text())
        winsound.Beep(9898, 1000)
    except Exception as e:
        

        try:
            # handeling laoder
            while True:
                if bro.find_element_by_xpath('''//*[@id="UpdateProgress"]/div/div''').is_displayed():
                    print('Loading Please Wait on line no 159')
                    time.sleep(1)
                    continue
                time.sleep(3)
                while True:
                    print('checking for alert in exception line no 183')
                   

                    '''Try: bro.switch_to_acative()'''
                    alert = bro.switch_to_active_element()
                    # print('alert could be on 189')
                    if alert.is_displayed():
                        print('we got an alert', alert.text)
                        
                        print('***** No data Found **** for k_no ',K_no,provider)
                       
                        bro.find_element_by_xpath(
                            '''//*[@id="7a68db42-b84a-4226-8ec6-063c87d481d2"]''').click()
                        
                        return
                    else:
                       
                        time.sleep(2)

        except Exception as e:
            print(e)
            try:
                while True:
                    # print('checking for alert in exception line no 207')
                    alerts = bro.switch_to_alert()
                    try:
                        print('*******ALERT:' 'No Data Found**** for user',K_no,provider)
                        # write_excel(row, K_no, provider,
                        #             'No Data Found', file_to_write)
                        # print('check for the ok button clikc')

                        bro.find_element_by_xpath(
                            '''//*[@id="555abe16-ea84-44aa-94bc-219e52ae4b9f"]''').click()
                        return
                    except Exception as e:
                        # print('NO ALERT FOUND on line on 218', e)
                        break
            except Exception as e:
                # print('line no 221', e)
                # gettting bill amount
                amount = bro.find_element_by_xpath(
                    '''//*[@id="billBox"]/div[1]/div[3]/div/table/tbody/tr[7]/td[4]/span''')
                print('****** The Bill Amount is : ', amount.text)
                command=input('Do you Want to pay now type Y for Yes and N for No').lower()
                if command=='n':
                    return 

                # bro.find_element_by_xpath(
                #     '''//*[@id="cpBody_btn_PayBill"]''').click()  # paybutton here down
                pay = bro.find_element_by_name("ctl00$cpBody$btn_PayBill")
                if pay.is_displayed():
                    pay.click()
                else:
                    print('pay not displayed')

            ''' Before Finding second pay button we need to conquer loading'''

        try:
            # find the pay button if failed return response 'Not Paid'
            # print('finding one pay')
            time.sleep(2)
            amount = bro.find_element_by_xpath(
                '''//*[@id="billBox"]/div[1]/div[3]/div/table/tbody/tr[7]/td[4]/span''')
            print('[ALERT] :=> The Bill Amount is : ', amount.text)
            if int(amount.text)== int(Amount):
                print('***** Bill Amount and excel amount is same')
            else:
                print('#####  Bill Amount and excel amount is not same')
                print('in final module process will get canceled from here')
            command=input('Do you Want to pay now --  type Y for Yes and N for No').lower()
            if command=='n':
                print('Payment canceled ')
                return             
            pay = bro.find_element_by_name("ctl00$cpBody$btn_PayBill")

            if pay.is_displayed():
                print('found pay2 on line no 254')
                pay.click()
            else:
                print('pay not displayed')
            time.sleep(2)
            print('Pending second pay button on line no 260')
            # pay2=bro.find_element_by_id("paymentModeTypeQrCodeId") no 1
            pay2 = bro.find_element_by_id("paymentModeTypeNetBId")
            if pay2.is_displayed():
                pay2.click()
                print('finding the back_to_sso on line no 265')
                back_to_sso = bro.find_element_by_xpath(
                    '''//*[@id="cpBody_btnBack"]''').click()
                print('going back to dashboard')
                dashboard()
            else:
                print('pay2 not displayed')
                pay2.click()

            # here we will directly write to excel file
            

            # write_excel(row, K_no, provider, 'PAID', file_to_write)
            return
        except Exception as e:
            print('WARNING :', e)
            print('exceptin on raise c in line 262')
            alert = bro.switch_to_active_element()['value']
            button_ok = button = alert.find_elements_by_tag_name('button')[1]
            button_ok.click()
            print(alert.text.split('\n')[0])
            print('[WARNING]','NO DATA FOUND')
            # write_excel(row, K_no, provider,
            #             alert.text.split('\n')[0], file_to_write)
            return

    # second_pay_button=bro.find_element_by_xpath('''//*[@id="paymentModeTypeNetBId"]''')
    pay2 = bro.find_element_by_id("paymentModeTypeQrCodeId")
    pay2.click()
    # back to sso
    back_to_sso = bro.find_element_by_xpath('''//*[@id="cpBody_btnBack"]''')
    dashboard()
    # print('click second pay')
    # second_pay_button.click()


def Extract_Excel(filename, all_data={}):
    ''' This function is used to get the discom,K No,Amount form Excel Sheet and
    returns dict version of all three first Dicom then K No then Amount'''
    data = pd.read_excel(filename)

    Discom = list(data.get('Discom'))

    K_no = list(data.get('K No'))
 
    Amount = list(data.get('Amount'))

    for k in range(len(Discom)):
        main_menu(k, Discom[k], Amount[k])

    print('Succesfull Extracted Data')

    return all_data
if __name__=='__main__':
    file_name='BILL CHECKED.xlsx'
    username='seshmoub'
    password='shubham@1234'
    login(username,password)
    dashboard()
    Extract_Excel(file_name)