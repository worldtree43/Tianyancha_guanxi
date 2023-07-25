from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time
import pandas as pd
import random
import os

# 从excel中获取公司名称
def get_companies(file_path, sheet_name, column_name):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    return df[column_name].tolist()

def tianyancha_relation_screenshot(companies, max_retry = 3):
    # 配置ChromeOptions，指定浏览器驱动路径
    chrome_options = Options()
    chrome_options.add_argument("--disable-gpu")  # 禁用GPU加速，解决部分环境下的问题
    chrome_options.add_argument("--window-size=1920x1080")  # 设置窗口大小

    # 创建一个Chrome浏览器实例，并最大化窗口
    driver = webdriver.Chrome(options=chrome_options)
    driver.maximize_window()

    # 打开天眼查关系页面
    url = 'https://www.tianyancha.com/relation'
    driver.get(url)

    # 设置一个等待时间，留出一定时间手动登录
    time.sleep(30)

    # 设置一个显式等待时间，等待搜索按钮可点击，最多等待10秒
    wait = WebDriverWait(driver, 10)

    for i in range(len(companies)):
        for j in range(i+1, len(companies)):
            retry_count = 0
            while retry_count < max_retry:
                try:
                    search_input_from = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="fromKey"]')))
                    actions = ActionChains(driver)
                    actions.move_to_element(search_input_from)
                    actions.click()
                    actions.send_keys(companies[i])
                    actions.perform()

                    search_input_to = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="toKey"]')))
                    actions = ActionChains(driver)
                    actions.move_to_element(search_input_to)
                    actions.click()
                    actions.send_keys(companies[j])
                    actions.perform()

                    # 等待搜索按钮出现并点击
                    search_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="web-content"]/section/div/div/div[1]/div[1]/span[2]')))
                    driver.execute_script("arguments[0].click();", search_button)

                    # 等待查询结果加载完成，这里设置为10秒，根据页面加载速度适当调整
                    time.sleep(random.uniform(7, 10))

                    # 等待图片加载完成点击下载按钮
                    download_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="web-content"]/section/div/div/div[2]/div/div[3]/div/div[3]/div')))
                    download_button.click()
                    time.sleep(random.uniform(2, 5))

                    close_button_from = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="web-content"]/section/div/div/div[1]/div[1]/div[1]/img')))
                    close_button_from.click()
                    close_button_to = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="web-content"]/section/div/div/div[1]/div[1]/div[3]/img')))
                    close_button_to.click()

                    break

                except Exception as e:
                    # 处理加载失败情况，点击搜索按钮再次尝试查询
                    print(f"加载失败，尝试重新查询，错误信息：{e}")
                    retry_count += 1
                    driver.refresh()
                    time.sleep(3)

            if retry_count >= max_retry:
                print(f"查询失败：{companies[i]} 和 {companies[j]}")
    # 关闭浏览器
    driver.quit()

def find_missing_files(folder_path, companies):
    missing_files = []
    for i in range(len(companies)):
        for j in range(i + 1, len(companies)):
            file_name = f"查关系图谱-{companies[i]}&{companies[j]}-天眼查.png"
            file_path = os.path.join(folder_path, file_name)
            if not os.path.exists(file_path):
                missing_files.append(file_name)
    return missing_files

if __name__ == "__main__":
    excel_file = '/Users/zhonghaitian/Desktop/公司名称.xlsx'
    sheet_name = 'Sheet1'
    company_column = '公司列表'
    companies = get_companies(excel_file, sheet_name, company_column)
    tianyancha_relation_screenshot(companies, max_retry=3)
    downloads_folder = "/Users/zhonghaitian/Downloads"
    missing_files = find_missing_files(downloads_folder, companies)

    if len(missing_files) == 0:
        print("所有文件已保存。")
    else:
        print("以下文件未保存：")
        for file_name in missing_files:
            print(file_name)