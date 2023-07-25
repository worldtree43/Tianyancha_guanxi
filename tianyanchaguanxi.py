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

excel_file_path = '/Users/zhonghaitian/Desktop/公司名称.xlsx'
df = pd.read_excel(excel_file_path)
companies = df['公司列表'].tolist()

# 配置ChromeOptions，指定浏览器驱动路径，并最大化窗口
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

        # 等待查询结果加载完成，这里设置为15秒，根据页面加载速度适当调整
        time.sleep(random.uniform(7, 10))

        # 等待图片加载完成点击下载按钮
        download_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="web-content"]/section/div/div/div[2]/div/div[3]/div/div[3]/div')))
        download_button.click()
        time.sleep(random.uniform(7, 10))

        # 因为该网站清除搜索内容的xpath地址不定，采取主动刷新页面方式清除搜索内容
        driver.refresh()
        time.sleep(random.uniform(8, 10))
# 关闭浏览器
driver.quit()

print("查询结果的截图已保存")