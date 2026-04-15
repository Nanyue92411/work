import subprocess
import time
import requests
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import ddddocr
from datetime import datetime, timedelta
import os
import glob
import shutil
from selenium.webdriver.support.ui import WebDriverWait
from selenium import webdriver

#========================基础配置==============================

URL = "https://corp.bank.ecitic.com/cotb/electronic/login.html"#网址
DOWNLOAD_folder = r"C:\Users\DELL\Desktop" #下载目录
TARGET_folder = r"D:\银行流水\2026年4月"  #目标目录

# ========================结束==================================

def get_captcha_text(driver):
    """
    从验证码元素获取图片字节流并识别
    """
    # 等待验证码图片元素可见且 src 属性不为空
    captcha_img = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "verifyCode"))
    )
    # 确保图片加载完成（Selenium 无法直接判断，建议短暂 sleep 或等待属性）
    time.sleep(0.5)  # 给图片一点加载时间

    img_url = captcha_img.get_attribute("src")
    print(f"验证码图片URL: {img_url}")

    # 获取当前浏览器的 cookies 用于下载
    cookies = {c["name"]: c["value"] for c in driver.get_cookies()}

    # 下载图片
    response = requests.get(img_url, cookies=cookies, timeout=10)
    if response.status_code != 200:
        raise Exception(f"下载验证码图片失败，状态码：{response.status_code}")

    # 使用 ddddocr 识别
    ocr = ddddocr.DdddOcr(show_ad=False)
    result = ocr.classification(response.content)
    print(f"识别结果：{result}")
    return result

# ===========================网页操作=================================

# 关闭所有 Chrome 进程 防止其他进程干扰 （可选，会丢失已有会话）
subprocess.run('taskkill /f /im chrome.exe', shell=True, capture_output=True)
time.sleep(2)
# 等待浏览器完全打开
chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
user_data_dir = r"D:\selenium_profile"
cmd = f'"{chrome_path}" --remote-debugging-port=9222 --user-data-dir="{user_data_dir}"'
subprocess.Popen(cmd, shell=True)
# time.sleep(3)

# 连接到调试浏览器
chrome_options = Options()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
driver = webdriver.Chrome()
# 打开登录页面
driver.get(URL)
print("当前页面标题:", driver.title)
time.sleep(3)
# 自动填写用户名
Username = driver.find_element(By.NAME, "userCode")
Username.send_keys("账号")
# 等待用户输入密码
time.sleep(15)

# ==============================验证码处理==========================================
MAX_TRIES = 15

for attempt in range(1, MAX_TRIES + 1):
    print(f"第 {attempt} 次尝试登录...")

    # 1. 获取并识别验证码
    captcha_text = get_captcha_text(driver)

    # 2. 输入验证码
    captcha_input = driver.find_element(By.ID, "verifyCodeInput")
    captcha_input.clear()
    captcha_input.send_keys(captcha_text)

    # 3. 点击登录按钮
    login_btn = driver.find_element(By.XPATH, "//button[@data-i18n-key='i0008']")
    login_btn.click()

    # 4. 判断登录结果
    try:
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, "//a[@class='first-menu']"))
        )
        print("✅ 登录成功！")
        break
    except:
        # 登录失败，检查是否有“关闭”按钮（验证码错误弹窗）
        try:
            close_btn = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.XPATH, "//button[text()='关闭']"))
            )
            close_btn.click()
            print("验证码错误，弹窗已关闭，准备重试...")
            # 关键：刷新验证码图片（可点击图片触发刷新）
            captcha_img = driver.find_element(By.ID, "verifyCode")
            captcha_img.click()
            time.sleep(1)
        except:
            print("未找到关闭按钮，可能账号密码错误或其他原因，停止重试")
            break
else:
    print(f"达到最大尝试次数 {MAX_TRIES}，登录失败。")
    driver.quit()
    exit()


# ========================账户余额及交易明细查询=============
# 点击账户明细
menu = driver.find_element(By.XPATH, "(//a[@class='first-menu'])[2]")
menu.click()
# 选择账号
# 1. 点击只读输入框，触发下拉面板
input_box = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.ID, "bsnAcc"))
)
input_box.click()
# 2. 等待下拉选项出现 选择网银账户 option = driver.find_element(By.XPATH, "//li[@data-value='123456789']")
option = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//li[contains(text(), '请输入网银账户')]"))
)
option.click()

# 处理日期选择器 选择昨日银行流水
# 第一步：算出目标日期
today     = datetime.now()
yesterday = today - timedelta(days=1)
# 日历的 data-ymd 格式是 2026-4-13（月和日没有补零）
# 所以格式化时用 %-m 和 %-d（Linux），Windows 用 %#m 和 %#d
date_str = f"{yesterday.year}-{yesterday.month}-{yesterday.day}"
print(date_str)  # 2026-4-13

# 第二步：点击输入框弹出日历
start_input = driver.find_element(By.NAME, "startDate")
start_input.click()

# 第三步：处理开始日期
wait = WebDriverWait(driver, 10)
wait.until(EC.visibility_of_element_located((By.ID, "jedatebox")))

# 第四步：直接找到目标日期的 li 并点击
target = driver.find_element(By.XPATH, f"//li[@data-ymd='{date_str}']")
target.click()

# 第五步：点确定
start_input = driver.find_element(By.NAME, "endDate")
start_input.click()

# 处理截止日期
wait = WebDriverWait(driver, 10)
wait.until(EC.visibility_of_element_located((By.ID, "jedatebox")))

# 直接找到目标日期的 li 并点击
target = driver.find_element(By.XPATH, f"//li[@data-ymd='{date_str}']")
target.click()

# 点击查询按钮
qry = driver.find_element(By.XPATH, "//button[@ng-click='qry()']")
qry.click()
time.sleep(3)

# 点击下载
download = driver.find_element(By.XPATH, "//a[text()='全部下载']")
download.click()
time.sleep(3)

# 选择Excel下载
download_select = driver.find_element(By.XPATH, "//a[text()='Excel下载']")
download_select.click()
time.sleep(3)

# 确保目标文件夹存在
os.makedirs(TARGET_folder, exist_ok=True)
# 获取原文件路径
old_path = glob.glob(os.path.join(DOWNLOAD_folder, "COBP*.xlsx"))[0]
# 复制到目标文件夹并重命名为日期.xlsx
new_path = os.path.join(TARGET_folder, f"{date_str}.xlsx")
shutil.move(old_path, new_path)


# =================================================================================================