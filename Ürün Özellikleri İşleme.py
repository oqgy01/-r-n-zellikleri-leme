import requests
from bs4 import BeautifulSoup
url = "https://docs.google.com/spreadsheets/d/1AP9EFAOthh5gsHjBCDHoUMhpef4MSxYg6wBN0ndTcnA/edit#gid=0"
response = requests.get(url)
html_content = response.content
soup = BeautifulSoup(html_content, "html.parser")
first_cell = soup.find("td", {"class": "s2"}).text.strip()
if first_cell != "Aktif":
    exit()
first_cell = soup.find("td", {"class": "s1"}).text.strip()
print(first_cell)

import pandas as pd
import re
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime, timedelta
from io import BytesIO
import os
import numpy as np
import shutil
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import time
from selenium.webdriver.chrome.service import Service
from concurrent.futures import ThreadPoolExecutor
from tqdm import tqdm
from selenium.common.exceptions import TimeoutException, WebDriverException
import xml.etree.ElementTree as ET
import warnings
from colorama import init, Fore, Style
import openpyxl
warnings.filterwarnings("ignore")
import tkinter as tk
from tkinter import simpledialog
import chromedriver_autoinstaller
pd.options.mode.chained_assignment = None




google_sheet_url = "https://docs.google.com/spreadsheets/d/1ECaRelQHEfEarkQHcapdjd9o1I_Ut2MvjnTYca8BHQ0/gviz/tq?tqx=out:csv"

try:
    google_df = pd.read_csv(google_sheet_url)
    google_excel_file = "E-Tablo Verileri.xlsx"
    
    # 2. ve 16. sütunlar dışındaki sütunları al
    columns_to_keep = [col for col_idx, col in enumerate(google_df.columns) if col_idx in [1, 15]] 
    
    # Veri çerçevesini istenilen sütunlarla sınırla
    google_df = google_df[columns_to_keep]
    
    # "Formül" sütununu 1. sütun olarak ayarla
    google_df = google_df.rename(columns={google_df.columns[1]: 'Formül'})
    
    # 1. sütundaki boş hücreleri içeren satırları sil
    google_df = google_df.dropna(subset=[google_df.columns[1]])
    
    # 0. sütundaki verilere başına "m1." ekle
    google_df.iloc[:, 0] = 'm1.' + google_df.iloc[:, 0].astype(str)
    
    # 0. sütundaki hücrelerde sondan 1. karakter "." değilse sonuna bir "." ekle
    google_df.iloc[:, 0] = google_df.iloc[:, 0].apply(lambda x: x if x[-1] == '.' else x + '.')

    # 0. sütunun adını "ModelKodu" olarak değiştir
    google_df.rename(columns={google_df.columns[0]: 'ModelKodu'}, inplace=True)
    
    # 1. sütunun adını "Aciklama" olarak değiştir
    google_df.rename(columns={google_df.columns[1]: 'Aciklama'}, inplace=True)

    # Verileri Excel dosyasına kaydet
    google_df.to_excel(google_excel_file, index=False)
except Exception as e:
    print("Bir hata oluştu:", e)

# Linkler
links = [
    "https://task.haydigiy.com/FaprikaXls/RADSBM/1/",
    "https://task.haydigiy.com/FaprikaXls/RADSBM/2/",
    "https://task.haydigiy.com/FaprikaXls/RADSBM/3/"
]

# Excel dosyalarını indirip birleştirme
dfs = []
for link in links:
    response = requests.get(link)
    if response.status_code == 200:
        # BytesIO kullanarak indirilen veriyi DataFrame'e dönüştürme
        df = pd.read_excel(BytesIO(response.content))
        dfs.append(df)

# Tüm DataFrame'leri birleştirme
merged_df = pd.concat(dfs, ignore_index=True)

# Sonuç DataFrame'ini tek bir Excel dosyasına yazma
merged_df.to_excel("Ürün Özellikleri.xlsx", index=False)

# Önce Excel dosyasını okuyalım
excel_file = "Ürün Özellikleri.xlsx"
df_merged = pd.read_excel(excel_file)

# 'Aciklama' sütunundaki dolu olan satırları filtrele
df_merged = df_merged[df_merged['Aciklama'].isna()]

# Güncellenmiş DataFrame'i aynı Excel dosyasının üstüne yaz
df_merged.to_excel('Ürün Özellikleri.xlsx', index=False)

# "sonuc_excel" dosyasını oku
sonuc_excel_file = "Ürün Özellikleri.xlsx"
sonuc_df = pd.read_excel(sonuc_excel_file)

# "E-Tablo Verileri" dosyasını oku
e_tablo_excel_file = "E-Tablo Verileri.xlsx"
e_tablo_df = pd.read_excel(e_tablo_excel_file)

# "ModelKodu" sütununda eşleşen satırları bul ve "Aciklama" sütununu kopyala
for index, row in sonuc_df.iterrows():
    model_kodu = row["ModelKodu"]
    e_tablo_row = e_tablo_df[e_tablo_df["ModelKodu"] == model_kodu]
    if not e_tablo_row.empty:
        sonuc_df.at[index, "Aciklama"] = e_tablo_row.iloc[0]["Aciklama"]

# Sonucu kaydet
sonuc_df.to_excel("Ürün Özellikleri.xlsx", index=False)

# Önce Excel dosyasını oku
excel_file = "Ürün Özellikleri.xlsx"
df = pd.read_excel(excel_file)

# "Aciklama" sütununda dolu olan satırları sil
filtered_df = df.dropna(subset=["Aciklama"])

# Kaydet
filtered_excel_file = "Ürün Özellikleri.xlsx"
filtered_df.to_excel(filtered_excel_file, index=False)

# Dosyayı sil
os.remove("E-Tablo Verileri.xlsx")


# Ürün YÜkleme Alanı
#Selenium Giriş Yapma
options = webdriver.ChromeOptions()

chromedriver_autoinstaller.install()
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('--log-level=1') 
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])  
driver = webdriver.Chrome(options=chrome_options)

login_url = "https://task.haydigiy.com/kullanici-giris/?ReturnUrl=%2Fadmin"
driver.get(login_url)

email_input = driver.find_element("id", "EmailOrPhone")
email_input.send_keys("mustafa_kod@haydigiy.com")

password_input = driver.find_element("id", "Password")
password_input.send_keys("123456")
password_input.send_keys(Keys.RETURN)

#region Excelle Ürün Yükleme Alanı
desired_url = "https://task.haydigiy.com/admin/importproductxls/edit/24"
driver.get(desired_url)

# Yükle Butonunu Bul
file_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[name="qqfile"]')))

# CalismaAlani Excel dosyasının mevcut çalışma dizininde olduğunu varsay
file_path = os.path.join(os.getcwd(), "Ürün Özellikleri.xlsx")

# Dosyayı seç
file_input.send_keys(file_path)


# "İşlemler" düğmesine tıkla
operations_button = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'btn-success')))
operations_button.click()

# Dosya yükleme işlemi bittikten sonra çalıştır butonuna tıkla
execute_button = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'import-product-xls-execute-confirm')))
execute_button.click()

# 10 saniye bekle
time.sleep(10)

def wait_for_element_and_click(driver, by, value, timeout=10):
    try:
        element = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))
        element.click()
        return True
    except (TimeoutException, WebDriverException) as e:
        print(f"Hata: {e}")
        return False

def wait_for_page_load(driver):
    while True:
        if driver.title:  # Tarayıcı başlığı varsa, sayfa yüklenmiş demektir
            break
        time.sleep(2)

# "Evet" butonunu tıkla
if wait_for_element_and_click(driver, By.ID, 'import-product-xls-execute'):
    # Yüklenmeyi Bekle
    wait_for_page_load(driver)

driver.quit()

# Dosya Adı
dosya_adlari = ['Ürün Özellikleri.xlsx']

# Dosyaları sil
for dosya_adi in dosya_adlari:
    try:
        os.remove(dosya_adi)
        
    except FileNotFoundError:
        print(f"{dosya_adi} bulunamadı.")
    except Exception as e:
        print(f"{dosya_adi} silinirken bir hata oluştu: {str(e)}")

print(Fore.GREEN + "Ürün Özellikleri E-Tablodan Çekilip Siteye Göre İşlendi")
#endregion
