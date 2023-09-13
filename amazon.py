from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support.ui import Select
import undetected_chromedriver as uc
from time import sleep
from selenium.common.exceptions import TimeoutException 
from selenium import webdriver
# from openpyxl import load_workbook
import random
import re
import pandas as pd
# Load the Excel file with pandas
df = pd.read_excel('links.xlsx')
# Convert the column into a list
links = df['URL'].tolist()

print('Starting...')

user_agents = [
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36 Edg/110.0.1587.41',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/109.0',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36'
]
custom_ua = random.choice(user_agents)

options = webdriver.ChromeOptions()
options.add_argument(f'user-agent={custom_ua}')
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)
driver = uc.Chrome(chrome_options=options)
driver.implicitly_wait(4)
wait = WebDriverWait(driver, 30)
driver.maximize_window()
for website in links:
    driver.get(website)
    try:
        accept_cookies_button = WebDriverWait(driver,4).until(EC.presence_of_element_located((By.XPATH, '//input[@id="sp-cc-accept"]')))
        action=webdriver.ActionChains(driver)
        action.move_to_element(accept_cookies_button).click().perform()
    except:
        pass
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    see_review=wait.until(EC.presence_of_element_located((By.XPATH,'//a[@data-hook="see-all-reviews-link-foot"]')))
    see_review.click()
    dropdown= wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,'#a-autoid-5-announce')))
    dropdown.click()

    critical_comment=wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,'#star-count-dropdown_7'))).click()
    weight_regex = re.compile(r"[0-9]+lbs|[0-9]+pound|[0-9]+\slbs|[0-9]+\spound")
    height_regex = re.compile(r"[0-9]+’[0-9]+|[0-9]+,[0-9]+|[0-9]+’\s[0-9]+”")
    height=''
    weight=''
    sleep(5)
    driver.switch_to.default_content()
    comments=wait.until(EC.presence_of_all_elements_located((By.XPATH,'//div[@data-hook="review"]')))

    for comment in comments:
        wait=WebDriverWait(comment, 10)
        name=wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,'span.a-profile-name'))).text
        rating=wait.until(EC.presence_of_element_located((By.XPATH,'.//span[@class="a-icon-alt"]'))).text
        text=wait.until(EC.presence_of_element_located((By.XPATH,'.//span[@data-hook="review-body"]'))).text
        review=wait.until(EC.presence_of_element_located((By.XPATH,'.//a[@class="a-link-normal"]/i/span'))).text
        try:
            size_color=wait.until(EC.presence_of_element_located((By.XPATH,'.//a[@class="a-size-mini a-link-normal a-color-secondary"]'))).text
            size_t=size_color.split(':')[1]
            size = size_t.replace("Color", "")
            color=size_color.split(':')[2]
        except Exception as e:
            size=''
            color=''
        
        # sleep(3)
        # text=wait.until(EC.presence_of_element_located((By.XPATH,'//span[@data-hook="review-body"]'))).text
        weight_matches = re.findall(weight_regex, text)

        height_matches = re.findall(height_regex, text)

        if weight_matches:
            weight = weight_matches[0]
        else:
            weight = ''
    

        if height_matches:
            height = height_matches[0]
        else:
            height = ''
    
        web_data={
            'name':name,
            'rating':review,
            'size':size,
            'color':color,
            'comment':text,
            'height':height,
            'weight':weight
        }
        print(web_data)
sleep(2000)

