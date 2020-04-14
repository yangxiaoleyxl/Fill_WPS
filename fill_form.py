# 每日健康统计完整版
# 由于必须登录才能填写，因此通过QQ验证身份
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time 
import numpy 
import win32api,win32con 

# 输入文本
def input_text(wait, item, content):
    answer = wait.until(EC.element_to_be_clickable((By.ID, item))) 
    answer.send_keys(content)
    return None 

# 单选题
def click_label(wait, item):
    answer = wait.until(EC.element_to_be_clickable((By.ID, item))) 
    answer.click()
    return None 

# 点击日历控件
def calender(wait): 
    answer = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.ant-calendar-picker'))) 
    time.sleep(3)
#     answer = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'ant-calendar-picker')))
    answer.click()
    answer = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.ant-calendar-today-btn ')))
    time.sleep(2) 
#     answer = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'ant-calendar-today-btn ')))
    answer.click()
    return None 

# 分时段单选题
def time_label(wait):
    localtime = time.localtime(time.time())
    if localtime.tm_hour < 9:
        t = 0
        print("填写时间为：0700-0900")
    elif localtime.tm_hour < 12:
        t = 1
        print("填写时间为：1100-1200")
    else:
        t = 2
        print("填写时间为：1800-2000") 
    time.sleep(1)
    answer = wait.until(EC.element_to_be_clickable((By.ID, 'select_label_wrap_5_' + str(t))))
    answer.click() 
    return None 

# 确认提交   
def submit_confirm(wait):
    time.sleep(4)
    commit = wait.until(EC.element_to_be_clickable((By.ID, 'submit_button')))
    commit.click() 
    time.sleep(1)
    yes = wait.until(EC.element_to_be_clickable((By.ID, 'bind_phone_modal_confirm_button')))
    yes.click() 
    return None 

options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-automation'])
options.add_experimental_option("prefs", {"profile.managed_default_content_settings.images": 2})

browser = webdriver.Chrome(options=options) 
wait = WebDriverWait(browser, 10)
# url = r'https://f.wps.cn/form-write/XXXX/' # WPS填表页面网址，通过Chrome打开 
browser.get(url)  

answer = wait.until(EC.element_to_be_clickable((By.ID, 'top_bar_login_button' )))
answer.click() 
time.sleep(2)  # 避免网页加载过慢出错 

answer = wait.until(EC.element_to_be_clickable((By.ID, 'qq')))
answer.click()
time.sleep(2)

answer = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, 'dialog-footer-ok')))
answer.click()
time.sleep(2)

# 此处网页中包含一个iframe结构，必须要在该结构下才能找到元素
browser.switch_to.frame('ptlogin_iframe') 
ele = browser.find_element_by_id('switcher_plogin') 
ActionChains(browser).move_to_element(ele).perform() 

answer = wait.until((EC.element_to_be_clickable((By.ID, 'switcher_plogin')))) 
answer.click() 
time.sleep(2) 

answer = wait.until((EC.element_to_be_clickable((By.ID, 'u')))) 
answer.send_keys('XXXXXXXXXX@qq.com') # 利用QQ邮箱登录
time.sleep(1)

answer = wait.until((EC.element_to_be_clickable((By.ID, 'p')))) 
answer.send_keys('XXXXX') # QQy邮箱密码
time.sleep(1)

answer = wait.until((EC.element_to_be_clickable((By.ID, 'login_button')))) 
answer.click() 
time.sleep(1)

answer = wait.until((EC.element_to_be_clickable((By.ID, 'write_again_button')))) 
answer.click() 
time.sleep(1)

try:
    time.sleep(3)
    input_text(wait, 'input_0', 'XXX')  # 填写姓名
    click_label(wait, f'select_label_wrap_1_0')  
    click_label(wait, f'select_label_wrap_2_0') 
    click_label(wait, f'select_label_wrap_3_0')
    calender(wait)
    time_label(wait)
    input_text(wait, 'input_6', '36.4') # 温度
# click_label(wait, f'select_label_wrap_6_0') 
    click_label(wait, f'select_label_wrap_7_0') 
    click_label(wait, f'select_label_wrap_8_0')  
    click_label(wait, f'select_label_wrap_9_0') 
    click_label(wait, f'select_label_wrap_10_1')  
    click_label(wait, f'select_label_wrap_11_0')
    click_label(wait, f'select_label_wrap_12_0')
    click_label(wait, f'select_label_wrap_13_0')
    submit_confirm(wait) 
    
    win32api.MessageBox(0, "填写完成", "提醒", win32con.MB_ICONWARNING )  # 弹窗提醒
except:
    win32api.MessageBox(0, "未提交!", "警告！", win32con.MB_ICONWARNING ) 