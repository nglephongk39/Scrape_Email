########## Tìm ra Email

import re
from selenium import webdriver
import xlwings as xw
from time import sleep
wb = xw.Book(r'C:\Users\nglep\Downloads\New folder\Buy_Account_LMHT\DATA_T8.2021_Copy.xlsx')
sh1 = wb.sheets('XK')

start = 950
end = 1050
List_Company_Name_Old = sh1[f'AE{start}:AE{end}'].value


for index in range(len(List_Company_Name_Old)):
    try:

        # Here we instantiate the webdriver, so that we can use it for the project.
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument('headless')
        chrome_options.add_argument('window-size=1920x1080')
        chrome_options.add_argument("disable-gpu")
        driver = webdriver.Chrome('chromedriver', chrome_options=chrome_options)
        # driver = webdriver.Chrome('C:\Program Files (x86)\chromedriver.exe')
        driver.get(List_Company_Name_Old[index])
        sleep(3)

        # Get the page source code
        page_source = driver.page_source

        # Regex to find e-mails
        EMAIL_REGEX = r"""(?:[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*|"(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21\x23-\x5b\x5d-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])*")@(?:(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?|\[(?:(?:(2(5[0-5]|[0-4][0-9])|1[0-9][0-9]|[1-9]?[0-9]))\.){3}(?:(2(5[0-5]|[0-4][0-9])|1[0-9][0-9]|[1-9]?[0-9])|[a-z0-9-]*[a-z0-9]:(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21-\x5a\x53-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])+)\])"""

        # Create a list and add all the emails
        list_of_emails = []

        # Finds all the emails
        for re_match in re.finditer(EMAIL_REGEX, page_source):
            list_of_emails.append(re_match.group())

    #New code

    

        sh1[f'AF{index + start}'].value = str(list_of_emails[0])
    except:
        sh1[f'AF{index + start}'].value = ''
    
    print(f'Finish Row {index + start}')
    driver.close()
        
    


# # # Lists all the e-mails that we managed to scrape
# # for i, email in enumerate(list_of_emails):
# #     print(f'{i + 1}: {email}')

# Close the driver since we don't need it
# driver.close()


# # # 1. Cách viết 1 list vào 1 ô trong excel


