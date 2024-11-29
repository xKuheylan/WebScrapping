from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains    
import pandas as pd
from collections import Counter
import re
import time
import keyboard
import os

# WebDriver'ı başlatma
service = Service(executable_path='./chromedriver.exe')
chromedriver_path = 'chromedriver.exe'
options = webdriver.ChromeOptions()
#options.add_argument('--headless') # --headless eklenirse tarayıcı görsel arayüzü olmadan çalışır
driver = webdriver.Chrome(service=service, options=options)

# İncelenmek istenen web sitesinin URL'sini belirtme
website_url = 'https://www.trendyol.com/xiaomi/android-xiaomi-mi-true-wireless-earphones-2-basic-bluetooth-kulaklik-uyumlu-p-45551953/yorumlar?boutiqueId=61&merchantId=106994'
driver.get(website_url)

# Sayfa içerisinde hareket etmek için "body" elementini bulma
body = driver.find_element(By.TAG_NAME, 'body')

for m in range(5000):
    body.send_keys(Keys.PAGE_DOWN)
    comments_elements = driver.find_elements(By.CSS_SELECTOR, 'div.comment-text > p') #yorum sayısı kontrol

    if m < 1000 and m % 50 == 0:
        for i in range(5):
            body.send_keys(Keys.PAGE_UP)
            os.system('cls')
            
    elif m % 50 == 0:
        # 1000'den sonra her 50 basışta PAGE_UP sayısını arttır
        for i in range(5):
            body.send_keys(Keys.PAGE_UP)
            os.system('cls')
            
    # Sayfa yüklenmesini beklemek için kısa bir süre bekletme
    time.sleep(2)
    
    # Sayaç sayısını ve anlık yorum sayısını ekrana yazdırma
    print(f'Sayaç sayısı: {m}')
    print(f"Anlık yorum sayısı: {len(comments_elements)}")
    
    #Döngu bitmeden çıkmak için q ya basarak çıkmak
    if keyboard.is_pressed('q'):
        print("Çıkış yapıldı!")
        break
    
# Yorumları içeren liste oluşturma
comments_elements = driver.find_elements(By.CSS_SELECTOR, 'div.comment-text > p')  
comments = [comment.text for comment in comments_elements]

# DataFrame oluşturma ve Excel'e yazdırma SADECE YORUMLARI
df = pd.DataFrame({'Comments': comments})
excel_file_path = 'TrendyolJustComments.xlsx'
df.to_excel(excel_file_path, index=False)

#Comments listesindeki yorumlara birer space koyarak texte aktarma
all_comments_text = ' '.join(comments)
#Yorumları kelime kelime ayırma
all_words = re.findall(r'\b\w+\b', all_comments_text.lower())
#Kelimeleri sayma
word_counts = Counter(all_words)   

# En çok geçen kelimeleri seçme
num_top_keywords = 200
top_keywords = [word for word, count in word_counts.most_common(num_top_keywords)]

# DataFrame oluşturma
keywords_df = pd.DataFrame({'Keywords': top_keywords})

# Yorumları içeren DataFrame oluşturma
comments_data = {'Comments': comments}
comments_df = pd.DataFrame(comments_data)

# Her bir anahtar kelimenin geçip geçmediğini kontrol etme
for word in top_keywords:
    comments_df[word] = comments_df['Comments'].str.lower().str.contains(word).astype(int)

# DataFrame oluşturma ve Excel'e yazdırma
analyzed_excel_file_path = 'TrendyolCommentsAndAttributes.xlsx'
comments_df.to_excel(analyzed_excel_file_path, index=False)

#Yıldızlar için liste oluşturma
comments_elements = driver.find_elements(By.CSS_SELECTOR, 'div.comment')
stars_list = []

for index, comment_element in enumerate(comments_elements):
    
    #Yıldızları width degeri ile ayırt edebilmek    
    width_100_count = comment_element.get_attribute("outerHTML").count("width: 100%")
    stars_list.append(f'{width_100_count-5}-Yildizli Yorum') # 5 adet width degeri yıldızla ilgili degil 
    
    
# DataFrame oluşturma ve Excel'e yazdırma
df = pd.DataFrame({'Stars ': stars_list}) 
excel_file_path = 'TrendyolStars.xlsx'
df.to_excel(excel_file_path, index=False)

print("Tamamlandı...")

driver.quit()
