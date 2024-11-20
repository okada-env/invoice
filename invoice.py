import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import re
import unicodedata
from Levenshtein import distance
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import difflib
from datetime import datetime  # 日付を取得するためのモジュールをインポート


def select_excel_file():
    root = Tk()
    root.withdraw()
    file_path = askopenfilename(title="エクセルファイルを選択", filetypes=[("Excel files", "*.xlsx *.xls")])
    return file_path


excel_file = select_excel_file()
if not excel_file:
    print("ファイルが選択されていません。終了します。")
    exit()

url = "https://www.invoice-kohyo.nta.go.jp/"
df = pd.read_excel(excel_file)

invoice_numbers = []
for index, value in enumerate(df['明細情報:フリー１(インボイス番号)']):
    if str(value) in ['N9999999999999', 'T9999999999999']:
        df.at[index, '企業名'] = "インボイス無し"
    else:
        number_only = re.sub(r'\D', '', str(value))
        invoice_numbers.append((index, number_only))

chrome_options = Options()
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
# Headlessオプションを無効化してブラウザを表示
# chrome_options.add_argument("--headless")  # 無効化または削除

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
driver.get(url)
time.sleep(2)

# 「sumSearchOn」をクリックし、10件登録可能に
try:
    add_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "sumSearchOn"))
    )
    add_button.click()
    time.sleep(2)
except Exception as e:
    print(f"初期設定エラー: {e}")
    driver.quit()
    exit()

for i in range(0, len(invoice_numbers), 10):
    batch_numbers = invoice_numbers[i:i+10]

    for j, (index, number) in enumerate(batch_numbers):
        # idが"regNo" + 変数番号(1〜10)で指定される
        search_box_id = f"regNo{j+1}"
        search_box = driver.find_element(By.ID, search_box_id)
        search_box.clear()
        search_box.send_keys(number)
    
    # 検索ボタンをクリック
    search_button = driver.find_element(By.CSS_SELECTOR, 'button[type="submit"]')
    search_button.click()
    time.sleep(2)  # 検索結果を確認するために待機時間を追加

    # 検索結果を取得
    for j, (index, number) in enumerate(batch_numbers):
        try:
            company_name_xpath = f'//*[@id="appForm"]/div[2]/div[1]/div[3]/table/tbody/tr[{j+1}]/td[1]'
            company_name_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, company_name_xpath))
            )
            
            full_text = company_name_element.text.strip()
            company_name = full_text.split("\n")[-1].strip()

            # 取得した企業名をデータフレームに反映
            df.at[index, '企業名'] = company_name if company_name else "企業名未取得"
        except Exception as e:
            print(f"企業名取得エラー: {e} (インデックス: {index})")
            df.at[index, '企業名'] = "取得エラー"

driver.quit()

# 現在の日付を取得してファイル名に埋め込む
current_date = datetime.now().strftime("%Y-%m-%d")
output_file = f"/Users/kazukiokada/Desktop/経費申請インボイス確認_{current_date}.xlsx"
df.to_excel(output_file, index=False)
print(f"検索結果が保存されました: {output_file}")

wb = load_workbook(output_file)
ws = wb.active
yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")


def remove_invisible_characters(text):
    return ''.join(c for c in text if not unicodedata.category(c).startswith('C'))


def normalize_text(text):
    text = unicodedata.normalize('NFKC', text)
    text = remove_invisible_characters(text)
    text = text.lower()
    text = re.sub(r'[\s\r\n]+', '', text)
    text = text.replace('㈱', '株式会社')
    text = re.sub(r'(株式会社|有限会社|合同会社|グループ|一般社団法人|一般財団法人|公益社団法人|公益財団法人|医療法人|学校法人|宗教法人|社会福祉法人|農業協同組合|漁業協同組合|生活協同組合|労働組合|特定非営利活動法人|独立行政法人|地方独立行政法人|特殊法人|特定目的会社|外国法人|㈱｜\（株\）|\(株\))', '', text)
    return text


def has_partial_match(g_name, h_name, length=4, threshold=3):
    if g_name == h_name:
        return True

    for i in range(len(g_name) - length + 1):
        substring = g_name[i:i + length]
        dist = distance(substring, h_name)
        if dist <= threshold:
            return True
    return False


for index, row in df.iterrows():
    invoice_number = str(row['明細情報:フリー１(インボイス番号)'])
    g_column_name = normalize_text(str(row['明細情報:フリー２(支払先)']))
    h_column_name = normalize_text(str(row['企業名']))

    if invoice_number in ['N9999999999999', 'T9999999999999']:
        continue

    if not has_partial_match(g_column_name, h_column_name, length=4, threshold=2):
        for col in range(1, 9):
            ws.cell(row=index + 2, column=col).fill = yellow_fill

wb.save(output_file)
print(f"ハイライト付きのファイルが保存されました: {output_file}")
