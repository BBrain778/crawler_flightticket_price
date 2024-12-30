from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import openpyxl
from datetime import datetime  # 用於獲取當前日期
import re  # 用於正則表達式處理

# 設定 WebDriver
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=options)

try:
    # 長榮數據長榮數據
    url = "https://www.google.com/travel/flights/booking?tfs=CBwQAho_EgoyMDI1LTA2LTA4Ih8KA1RQRRIKMjAyNS0wNi0wOBoDQ0VCKgJCUjIDMjgxagcIARIDVFBFcgcIARIDQ0VCGj8SCjIwMjUtMDctMDMiHwoDQ0VCEgoyMDI1LTA3LTAzGgNUUEUqAkJSMgMyODJqBwgBEgNDRUJyBwgBEgNUUEVAAUgBcAGCAQsI____________AZgBAQ&tfu=CmxDalJJYzBsSFUxZHphMmhVY0hOQlJuZEdiMEZDUnkwdExTMHRMUzB0ZEdKaWJXc3hNMEZCUVVGQlIyUjRTRVk0VDBrNE4yMUJFZ2RDVWpJNE1pTXlHZ29JbjBRUUFCb0RWRmRFT0J4dzlNOEISAggAIgMKATE&hl=zh-TW"
    driver.get(url)

    # 等待目標元素加載
    wait = WebDriverWait(driver, 10)  # 最多等待 10 秒
    target_element = wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, 'span[data-gs]')
        )
    )
    
    price_text1 = target_element.get_attribute('aria-label')  # 確保抓取 aria-label 屬性中的值
    print(f"長榮:抓取的價格（完整）是: {price_text1}")
    
    price_number1 = int(re.search(r'\d+', price_text1).group())  # 使用正則表達式提取第一個數字
    print(f"長榮:提取的價格（數字部分）是: {price_number1}")
    
    
    # 華航數據
    url = "https://www.google.com/travel/flights/booking?tfs=CBwQAhpEEgoyMDI1LTA2LTA4Ih8KA1RQRRIKMjAyNS0wNi0wOBoDQ0VCKgJDSTIDNzA1agwIAhIIL20vMGZ0a3hyBwgBEgNDRUIaRBIKMjAyNS0wNy0wMyIfCgNDRUISCjIwMjUtMDctMDMaA1RQRSoCQ0kyAzcwNmoHCAESA0NFQnIMCAISCC9tLzBmdGt4QAFIAXABggELCP___________wGYAQE&tfu=CmxDalJJY2pOSmJIaDRRVEUxVmtWQlIwOWhNVUZDUnkwdExTMHRMUzB0ZEdKaVltc3pNRUZCUVVGQlIyUjRUa3BWUVZWdk0wdEJFZ2REU1Rjd05pTXlHZ29JbXpvUUFCb0RWRmRFT0J4d3JMRUISAggAIgMKATE&hl=zh-TW"
    driver.get(url)

    # 等待目標元素加載
    wait = WebDriverWait(driver, 10)  # 最多等待 10 秒
    target_element = wait.until(
        EC.presence_of_element_located(
            (By.CSS_SELECTOR, 'span[data-gs]')
        )
    )
    
    price_text2 = target_element.get_attribute('aria-label')  # 確保抓取 aria-label 屬性中的值
    print(f"華航:抓取的價格（完整）是: {price_text2}")
    
    price_number2 = int(re.search(r'\d+', price_text2).group())  # 使用正則表達式提取第一個數字
    print(f"華航:提取的價格（數字部分）是: {price_number2}")



    # 獲取當前日期，格式為「12月28日」
    current_date = datetime.now().strftime("%m月%d日")
    print(f"抓取時間: {current_date}")

    # 將抓取到的數據保存到 Excel
    #excel_file = "price_data.xlsx"  # Excel 檔案名稱
    excel_file = r"D:\code\機票票價\price_data.xlsx" # Excel 檔案名稱

    try:
        # 嘗試加載已存在的檔案
        workbook = openpyxl.load_workbook(excel_file)
        sheet = workbook.active
    except FileNotFoundError:
        # 如果檔案不存在，則新建一個
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Flight Prices"
        # 添加標題行
        sheet.append(["日期", "長榮價格", "華航價格"])

    # 寫入當前日期與價格
    sheet.append([current_date, price_number1, price_number2])

    # 保存 Excel 檔案
    workbook.save(excel_file)
    print(f"數據已保存到 {excel_file}")

except TimeoutException:
    print("目標元素未在指定時間內加載完成。")
except Exception as e:
    print(f"發生錯誤: {e}")
finally:
    # 關閉瀏覽器
    driver.quit()
