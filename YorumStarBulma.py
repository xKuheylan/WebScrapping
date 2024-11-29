from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
import time
import pandas as pd


service = Service(executable_path='./chromedriver.exe')
chromedriver_path = 'chromedriver.exe'
options = webdriver.ChromeOptions()
options.add_argument('--headless')
driver = webdriver.Chrome(service=service, options=options)

website_url = 'https://www.trendyol.com/technomen/akilli-saat-plus-kablosuz-kulaklik-ikili-siyah-set-ios-android-smartwatch-p-347921451/yorumlar?boutiqueId=61&merchantId=309016'
driver.get(website_url)

for _ in range(2500):
    body = driver.find_element(By.TAG_NAME, 'body')
    body.send_keys(Keys.PAGE_DOWN)
    time.sleep(3)
    print(_)

comments_elements = driver.find_elements(By.CSS_SELECTOR, 'div.comment')

stars_list = []

for index, comment_element in enumerate(comments_elements):
        
    width_100_count = comment_element.get_attribute("outerHTML").count("width: 100%")
    stars_list.append(width_100_count)   
    
    
df = pd.DataFrame({'Stars ': stars_list}) 
excel_file_path = 'TrendyolStars.xlsx'
df.to_excel(excel_file_path, index=False)

print("Done...")

driver.quit()
