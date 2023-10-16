from csv import excel
import time
from selenium import webdriver 
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.support import expected_conditions as EC # for Ajax see stackoverflow bookmark
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
import pandas as pd
import xlwings as xw
import os
import win32com.client
import json
import string
import openpyxl

driver = webdriver.Firefox(options=Options())
# load website 

url = 'https://exhibitors.expandnorthstar.com/north-star-2023/Exhibitor'

# get the entire website content 
driver.get(url)
driver.maximize_window()

time.sleep(2)

# Create object of the Select class
select_country = Select(driver.find_element(By.XPATH, "//*[@id='country_id']"))
            
# Select the option with value "United Arab Emirates"
select_country.select_by_value("United Arab Emirates")

time.sleep(1)
search_box = driver.find_element(By.ID, 'keyword_search')

#Write what will be searched
search_box.send_keys('data')

#Submit the text
search_box.send_keys(Keys.RETURN)

#Now we need to click on each alphabet A-Z and put it in a data frame

time.sleep(1)
WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#alphabet_A'))).click()

time.sleep(1)

# #************************************************************
#Function to check if element exists by xpath
def check_exists_by_xpath(xpath):
    try:
        driver.find_element_by_xpath(xpath)
    except NoSuchElementException:
        return False
    return True

def check_exists_by_css(css):
    try:
        driver.find_element_by_css_selector(css)
    except NoSuchElementException:
        return False
    return True

df = pd.DataFrame(columns=(['Company_Name','Description','Website','LinkedIn']))

for alphabet in list(string.ascii_uppercase):
# for alphabet in ['K']:
    for i in range(1,100):
        time.sleep(1)
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#alphabet_"+ str(alphabet)))).click()
        time.sleep(0.5)
        if not check_exists_by_xpath('//*[@id="load_data"]/div[' + str(i) + ']/div/div/div[3]/div[1]/a'):
            break
        else:
            time.sleep(0.5)
            WebDriverWait(driver,20).until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="request"]')))
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="load_data"]/div[' + str(i) + ']/div/div/div[3]/div[1]/a'))).click()
            if check_exists_by_css(".linkdin_link"):
                linkedin = driver.find_element_by_css_selector(".linkdin_link").get_attribute('href')
            else:
                linkedin = ""

            if check_exists_by_css(".social_website > a:nth-child(1)"):
                website = driver.find_element_by_css_selector(".social_website > a:nth-child(1)").get_attribute('href')
            else:
                website = ""    

            elem_data = {
            'Company_Name': driver.find_element_by_css_selector("h4").text,
            'Description': driver.find_element_by_css_selector("p.group").text,
            'Website': website,
            'LinkedIn': linkedin
            }

            df_dictionary = pd.DataFrame([elem_data])
            df = pd.concat([df, df_dictionary], ignore_index=True)

            driver.back()



            # element = driver.find_element_by_xpath("//*[@id='country_id']")
            # driver.execute_script("return arguments[0].scrollIntoView(true);", element)



            time.sleep(1)
            search_box = driver.find_element(By.ID, 'keyword_search')
            # driver.execute_script("return arguments[0].scrollIntoView(true);", element)
            select_country = Select(driver.find_element(By.XPATH, "//*[@id='country_id']"))
            time.sleep(1)

            driver.find_element(By.XPATH, '//body').send_keys(Keys.CONTROL + Keys.HOME)
            # driver.execute_script("return arguments[0].scrollIntoView(true);", element)
            WebDriverWait(driver,20).until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="request"]')))
            select_country.select_by_value("United Arab Emirates")
            #Write what will be searched
            search_box.send_keys('data')

            #Submit the text
            search_box.send_keys(Keys.RETURN)
            time.sleep(2)
            WebDriverWait(driver,20).until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="request"]')))
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#alphabet_"+ str(alphabet)))).click()
    
    print(alphabet + " is done")

driver.close()
print(df)
#************************************************************************************************************
#Now for product sector = 'Big Data & Analytics'
driver = webdriver.Firefox(options=Options())
# load website 

url = 'https://exhibitors.expandnorthstar.com/north-star-2023/Exhibitor'

driver.get(url)
time.sleep(2)

WebDriverWait(driver,20).until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="request"]')))
WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div[1]/div/div[3]/span/div/button"))).click()
WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div[1]/div/div[3]/span/div/div/button[5]"))).click()
time.sleep(1)

select_country = Select(driver.find_element(By.XPATH, "//*[@id='country_id']"))
# Select the option with value "United Arab Emirates"
select_country.select_by_value("United Arab Emirates")

for alphabet in list(string.ascii_uppercase):
# for alphabet in ['G']:
    for i in range(1,100):
        time.sleep(2.5)
        WebDriverWait(driver,20).until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="request"]')))
        WebDriverWait(driver, 40).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#alphabet_"+ str(alphabet)))).click()
        time.sleep(1)
        if not check_exists_by_xpath('//*[@id="load_data"]/div[' + str(i) + ']/div/div/div[3]/div[1]/a'):
            break
        else:
            time.sleep(1)
            WebDriverWait(driver,20).until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="request"]')))
            WebDriverWait(driver, 40).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="load_data"]/div[' + str(i) + ']/div/div/div[3]/div[1]/a'))).click()
            if check_exists_by_css(".linkdin_link"):
                linkedin = driver.find_element_by_css_selector(".linkdin_link").get_attribute('href')
            else:
                linkedin = ""

            if check_exists_by_css(".social_website > a:nth-child(1)"):
                website = driver.find_element_by_css_selector(".social_website > a:nth-child(1)").get_attribute('href')
            else:
                website = ""    

            elem_data = {
            'Company_Name': driver.find_element_by_css_selector("h4").text,
            'Description': driver.find_element_by_css_selector("p.group").text,
            'Website': website,
            'LinkedIn': linkedin
            }

            df_dictionary = pd.DataFrame([elem_data])
            df = pd.concat([df, df_dictionary], ignore_index=True)

            driver.back()
            time.sleep(2)
            driver.find_element(By.XPATH, '//body').send_keys(Keys.CONTROL + Keys.HOME)
            WebDriverWait(driver,20).until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="request"]')))
            WebDriverWait(driver, 40).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div[1]/div/div[3]/span/div/button"))).click()
            WebDriverWait(driver, 40).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div/div[1]/div/div[3]/span/div/div/button[5]"))).click()
            time.sleep(2)
            select_country = Select(driver.find_element(By.XPATH, "//*[@id='country_id']"))
            time.sleep(2)
            driver.find_element(By.XPATH, '//body').send_keys(Keys.CONTROL + Keys.HOME)
            WebDriverWait(driver,20).until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="request"]')))
            select_country.select_by_value("United Arab Emirates")

            #Submit the text
            time.sleep(1.5)
            driver.find_element(By.XPATH, '//body').send_keys(Keys.CONTROL + Keys.HOME)
            WebDriverWait(driver,20).until(EC.invisibility_of_element_located((By.XPATH, '//*[@id="request"]')))
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#alphabet_"+ str(alphabet)))).click()

    print(alphabet + " is done")

driver.close()
print(df)
df.drop_duplicates(inplace=True)
print(df)
df.to_excel("C:/Users/geodx/Desktop/exhibit_list_all_data.xlsx")
