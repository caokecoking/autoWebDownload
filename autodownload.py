from selenium import webdriver
import os
import path
from pathlib import Path
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from PIL import Image
import time
from yundama import YDMHttp
import csv

### yundama ------------------------------------- 
ydmUsername    = 'seaxw'            # 用户名
ydmPassword    = 't46900_'          # 密码                     
appid       = 1                     # 开发者相关 功能使用和用户无关                
appkey      = '22cc5376925e9387a23cf797cb9ba745'    # 开发者相关 功能使用和用户无关   
filename    = 'capcha.png'          # 验证码截图              
# 验证码类型，# 例：1004表示4位字母数字，不同类型收费不同。请准确填写，否则影响识别率。在此查询所有类型 http://www.yundama.com/price.html
codetype    = 1004
# 超时时间，秒
timeout     = 60                                    
# 检查
if (ydmUsername == ''):
    print('请设置好相关参数再测试')
else:
    # 初始化
    yundama = YDMHttp(ydmUsername, ydmPassword, appid, appkey)
    # 登陆云打码
    uid = yundama.login()
    print('云打码登录成功 uid: %s' % uid)
    # # 查询余额
    # balance = yundama.balance()
    # print('balance: %s' % balance)

    # # 开始识别，图片路径，验证码类型ID，超时时间（秒），识别结果
    # cid, result = yundama.decode(filename, codetype, timeout)
    # print('cid: %s, result: %s' % (cid, result))
### yundama ------------------------------------- 

def get_snap(driver):  # 对目标网页进行截屏。这里截的是全屏
    driver.save_screenshot('full_snap.png')
    page_snap_obj=Image.open('full_snap.png')
    return page_snap_obj
 
 
def get_image(driver): # 对验证码所在位置进行定位，然后截取验证码图片
    xp = '//*[@id="UpdatePanel1"]/ul/li[4]/img'
    img = driver.find_element_by_xpath(xp)
    time.sleep(2)
    location = img.location
    size = img.size
    left = location['x']
    top = location['y']
    right = left + size['width']
    bottom = top + size['height']
    page_snap_obj = get_snap(driver)
    image_obj = page_snap_obj.crop((left, top, right, bottom))
    # image_obj.show()
    image_obj.save("capcha.png")
    return image_obj  # 得到的就是验证码


def is_download_finished(temp_folder):
    firefox_temp_file = sorted(Path(temp_folder).glob('*.part'))
    chrome_temp_file = sorted(Path(temp_folder).glob('*.crdownload'))
    downloaded_files = sorted(Path(temp_folder).glob('*.*'))
    if (len(firefox_temp_file) == 0) and \
       (len(chrome_temp_file) == 0) and \
       (len(downloaded_files) >= 1):
        return True
    else:
        return False


def hr01(url,pageUsername,pagePassword,folder):

    PROJECT_ROOT = os.path.abspath(os.path.dirname(__file__))
    DRIVER_BIN = os.path.join(PROJECT_ROOT, "chromedriver")
    chrome_options = webdriver.ChromeOptions()
    dlto = os.path.join(PROJECT_ROOT,folder)
    prefs = {'download.default_directory' : dlto}
    chrome_options.add_experimental_option('prefs', prefs)
    # driver = webdriver.Chrome(chrome_options=chrome_options)
    driver = webdriver.Chrome(executable_path = DRIVER_BIN, chrome_options=chrome_options)
    driver.get(url)

    username = driver.find_element_by_id("txtname")
    password = driver.find_element_by_id("txtpwd")
    vCode = driver.find_element_by_id("txtcheckcode")

    get_image(driver)
    cid, resultCode = yundama.decode(filename, codetype, timeout)
    print(resultCode)

    username.send_keys(pageUsername)
    password.send_keys(pagePassword)
    vCode.send_keys(resultCode)

    driver.find_element_by_name("ImgSubmit").click()
    time.sleep(3)
    driver.find_element_by_link_text("流向查询").click()
    time.sleep(3)
    driver.find_element_by_link_text("库存查询").click()
    time.sleep(3)
    kcFrame = driver.find_element_by_xpath("//iframe[contains(@src,'storelist')]")
    driver.switch_to.frame(kcFrame)
    driver.find_element_by_xpath("/html/body/div[1]/div[2]/table/tbody/tr[1]/td[2]/span[1]/input[1]").clear()
    time.sleep(1)
    driver.find_element_by_xpath("/html/body/div[1]/div[2]/table/tbody/tr[1]/td[2]/span[1]/input[1]").send_keys("2019-06-08")
    time.sleep(1)
    driver.find_element_by_xpath('/html/body/div[1]/div[2]/table/tbody/tr[1]/td[2]/span/span/span').click()
    time.sleep(1)
    driver.find_element_by_link_text("查询").click()
    time.sleep(1)
    driver.find_element_by_link_text("导出").click()
    limit = 10
    finished = False
    while limit > 0:
        limit = limit -1
        time.sleep(1)
        if is_download_finished(folder):
            limit = 0
            finished = True
        elif limit == 0 and  finished == False:
            print('download failed or take too long!')

    input('Press ENTER to close the automated browser')

    driver.quit()



with open('wdt.csv',newline='') as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:
        folder = row['code']
        workerID = row['cat']
        url = row['url']
        pageUsername = row['useracc']
        pagePassword = row['passwd']

        # 检测目标文件夹是否存在
        if not os.path.exists(folder):
            os.mkdir(folder)

        # 删除目标文件夹中的firefox和chrome下载专用临时文件
        apath = os.path.abspath(os.path.dirname(__file__))
        bpath = os.path.join(apath, folder)
        fl = os.listdir(bpath)
        for item in fl:
            if item.endswith('.part','.crdownload'):
                os.remove(os.path.join(bpath,item))

        globals()[workerID](url,pageUsername,pagePassword,folder)
        
