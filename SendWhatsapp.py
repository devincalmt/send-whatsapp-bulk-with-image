# Program to send bulk messages through WhatsApp web from an excel sheet without saving contact numbers
# Author @inforkgodara

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from time import sleep
import pandas
import urllib.parse
import os
import platform

if platform.system() == 'Darwin':
    # MACOS Path
    chrome_default_path = os.getcwd() + '\\driver\\chromedriver'
else:
    # Windows Path
    chrome_default_path = os.getcwd() + '\\driver\\chromedriver.exe'

excel_data = pandas.read_excel('Data.xlsx', sheet_name='Final')

count = 0

chrome_options = Options()
chrome_options.add_argument('--user-data-dir={}'.format(os.getcwd()+'\\User_Data'))

driver = webdriver.Chrome(executable_path=chrome_default_path, options=chrome_options)
driver.get('https://web.whatsapp.com')
input("Press ENTER after login into Whatsapp Web and your chats are visiable.")

fail_list = []

for column in excel_data['Phone'].tolist():
    try:
        url = 'https://web.whatsapp.com/send?phone=' + str(int(excel_data['Phone'][count])) + '&text=' + urllib.parse.quote(excel_data['Text'][count])
        sent = False
        
        driver.get(url)

        check_fail = []

        images = excel_data['Image'][count].split(";")

        if(excel_data['Image'][count] != ''):
            try:
                attachment_btn = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.CLASS_NAME, '_2jitM')))
            except Exception as e2:
                print("Message Attachment: "+str(e2))
                # check_fail.append('Image '+str(idx+1)+' '+str(e))
            else:
                sleep(2)
                attachment_btn.click()

                try:
                    choose_image_btn = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, "//input[@accept='image/*,video/mp4,video/3gpp,video/quicktime']")))
                    choose_image_btn.send_keys(os.getcwd() + "\\Image\\" + excel_data['Image'][count])
                    sleep(3)
                except Exception as e3:
                    print("Message Image: "+str(e3))
                    # check_fail.append('Image '+str(idx+1)+' '+str(e))
                else:
                    try:
                        send_image_btn = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.CLASS_NAME, '_33pCO')))
                    except Exception as e5:
                        print("Message Image: "+str(e5))
                        # check_fail.append('Image '+str(idx+1)+' '+str(e))
                    else:
                        sleep(2)
                        send_image_btn.click()
                        sleep(3)
        else:
            try:
                send_message_btn = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.CLASS_NAME, '_1Ae7k')))
            except Exception as e1:
                check_fail.append('Message')
            else:
                sleep(2)
                send_message_btn.click()

        # if count(check_fail) == 0:
        #     print('Message sent successfully to: ' + str(excel_data['Phone'][count]))
        # else:
        #     fail_list.append(excel_data['Phone'][count])
        #     print('There are some problem with {} for {}'.format(" - ".join(check_fail), str(excel_data['Phone'][count])))

        sleep(5)

        count = count + 1
    except Exception as e4:
        fail_list.append(excel_data['Phone'][count])
        print('Failed to send message to ' + str(excel_data['Phone'][count]) + ' '+str(e4))

driver.quit()
print("The script executed successfully.")

if len(fail_list) > 0:
    print("List Fail to Send")
    for fail in fail_list:
        print(fail)
