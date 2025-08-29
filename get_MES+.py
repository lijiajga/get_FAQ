from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time


# 配置 ChromeOptions
chrome_options = Options()
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": r"D:\download\MES+指导书",  # 替换为实际的下载路径
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

# 初始化 WebDriver
driver = webdriver.Chrome(options=chrome_options)

driver.maximize_window()

# 打开目标网站
driver.get('https://smes.xfusion.com/smes/cpwrpt/#/common/InstructionFiling')  # 替换为实际的网站地址

# 等待页面加载完成
wait = WebDriverWait(driver, 20)

try:
    # 假设需要登录，先进行登录操作
    # 输入用户名和密码
    username = wait.until(EC.presence_of_element_located((By.NAME, 'username')))
    password = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="password-item"]/div/div/input')))
    username.send_keys('p50011748')
    password.send_keys('Pan787127839!')
    
    time.sleep(10)
    # 点击登录按钮
    login_button = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div[4]/div/div[2]/div/form/div[3]/button')))
    login_button.click()
    
    # 进入下载页面
    download_page = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="pane-1"]/div/div[1]/div[3]/div/button[1]')))
    download_page.click()
    time.sleep(5)


    total_files = 100
    files_per_page = 10
    total_pages = (total_files + files_per_page - 1) // files_per_page  # 计算总页数
    
    current_page = 1
    downloaded_count = 0
    download_links = []
    for i in range(1,11):
        download_links.append(f'//*[@id="pane-1"]/div/div[3]/div/div[2]/div[3]/table/tbody/tr[{i}]/td[10]/div/a')
    # print(download_links)
    while current_page <= total_pages and downloaded_count < total_files:
        # 等待当前页面加载完成
        time.sleep(10)
        # 遍历下载链接
        for link in download_links:
            if downloaded_count >= total_files:
                break
            try:
                # 点击下载链接
                download_file = wait.until(EC.element_to_be_clickable((By.XPATH,link)))
                download_file.click()
                time.sleep(10)
                downloaded_count += 1
                print(f'已下载第 {downloaded_count} 个文件')
                # 等待下载完成（根据实际情况调整）
                time.sleep(10)
                
                # # 如果下载后打开了新窗口，切换回主窗口
                # if len(driver.window_handles) > 1:
                #     driver.switch_to.window(driver.window_handles[0])
            except Exception as e:
                print(f'下载第 {downloaded_count} 个文件时出错：{e}')
                continue
        
        # 点击下一页按钮
        if current_page < total_pages:
            next_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="pane-1"]/div/div[3]/div/div[3]/button[2]')))
            next_button.click()
            current_page += 1
            time.sleep(10)
        else:
            break

finally:
    # 关闭浏览器
    driver.quit()
