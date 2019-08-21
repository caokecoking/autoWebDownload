import datetime
from selenium import webdriver
import os
# import path
from pathlib import Path
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC, ui
from selenium.webdriver.support.wait import WebDriverWait
from PIL import Image
import time
from yundama import YDMHttp
import csv

# 当天日期、昨天日期
today = datetime.datetime.now()
yesterday = (today - datetime.timedelta(days=1)).strftime("%Y-%m-%d")
# 获取当月第一天日期
first_day = datetime.datetime(today.year, today.month, 1).strftime("%Y-%m-%d")
print(yesterday)
# ------------------------yundama -------------------------------------
ydmUsername = 'seaxw'  # 用户名
ydmPassword = 't46900_'  # 密码
appid = 1  # 开发者相关 功能使用和用户无关
appkey = '22cc5376925e9387a23cf797cb9ba745'  # 开发者相关 功能使用和用户无关
filename = 'cache.png'  # 验证码截图
# 验证码类型，# 例：1004表示4位字母数字，不同类型收费不同。请准确填写，否则影响识别率。在此查询所有类型 http://www.yundama.com/price.html
codetype = 1004
# 超时时间，秒
timeout = 60

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
    page_snap_obj = Image.open(r'full_snap.png')
    return page_snap_obj


def get_image(driver, xp, a, b, c, d):  # 对验证码所在位置进行定位，然后截取验证码图片
    # driver.refresh()  # 刷新页面
    # driver.maximize_window()  # 浏览器最大化
    # driver.set_window_size(1920,1080)
    # width_driver = driver.get_window_size()['width']
    # xp = "//*[@id='UpdatePanel1']/ul/li[4]/img"
    img = driver.find_element_by_xpath(xp)
    time.sleep(2)
    location = img.location
    size = img.size
    left = location['x'] + a
    top = location['y'] + b
    right = left + size['width'] + c
    bottom = top + size['height'] + d
    print((left, top, right, bottom))
    page_snap_obj = get_snap(driver)
    # driver.save_screenshot('full_snap.png')
    # page_snap_obj = Image.open('full_snap.png')
    image_obj = page_snap_obj.crop((left, top, right, bottom))
    # image_obj.show()
    image_obj.save("cache.png")
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


#  ---------------------------------------------主要方法-------------------------------------
# 获取driver
def getDriver(folder):
    PROJECT_ROOT = os.path.abspath(os.path.dirname(__file__))  # 项目根目录 -- 定义web driver时使用该目录，由于chromedriver.exe存放于该项目下
    DRIVER_BIN = os.path.join(PROJECT_ROOT, "chromedriver.exe")
    options = webdriver.ChromeOptions()
    dlto = os.path.join(PROJECT_ROOT, folder)
    prefs = {'download.default_directory': dlto}
    options.add_experimental_option('prefs', prefs)
    driver = webdriver.Chrome(executable_path=DRIVER_BIN, chrome_options=options)
    return driver


# 营销华润 -- 已完成（库存、销售）
def hr01(url, pageUsername, pagePassword, folder):
    # 获取驱动
    driver = getDriver(folder)
    driver.get(url)
    # 定位用户名、密码、验证码输入框、验证码图片
    username = driver.find_element_by_id("txtname")
    password = driver.find_element_by_id("txtpwd")
    vCode = driver.find_element_by_id("txtcheckcode")
    xp = "//*[@id='UpdatePanel1']/ul/li[4]/img"
    # 打码
    get_image(driver, xp, 179, 80, 10, 7)
    cid, resultCode = yundama.decode(filename, codetype, timeout)
    print(resultCode)
    # 用户名、密码、验证码赋值
    username.send_keys(pageUsername)
    password.send_keys(pagePassword)
    vCode.send_keys(resultCode)
    # 登录系统
    driver.find_element_by_name("ImgSubmit").click()
    time.sleep(3)
    driver.find_element_by_link_text("流向查询").click()
    time.sleep(3)
    # 库存流向查询、销售流向查询
    getSelect(driver, "库存查询", 'storelist', first_day, yesterday)
    driver.switch_to.default_content()
    getSelect(driver, "销售查询", 'salelist', first_day, yesterday)

    input('Press ENTER to close the automated browser')
    driver.quit()


# 流向查询（库存、销售）
def getSelect(driver, select_name, src_name, start_time, end_time):
    driver.find_element_by_link_text(select_name).click()
    time.sleep(3)
    kcFrame = ''
    if src_name == 'storelist':
        kcFrame = driver.find_element_by_xpath("//iframe[contains(@src,'storelist')]")
    #     DFS/work/salelist.html
    if src_name == 'salelist':
        kcFrame = driver.find_element_by_xpath("//iframe[contains(@src,'salelist')]")
    driver.switch_to.frame(kcFrame)
    # /html/body/div[1]/div[2]/table/tbody/tr[1]/td[2]/span/input[1]
    if src_name == 'storelist':
        driver.find_element_by_xpath("/html/body/div[1]/div[2]/table/tbody/tr[1]/td[2]/span[1]/input[1]").clear()
        time.sleep(1)
        driver.find_element_by_xpath("/html/body/div[1]/div[2]/table/tbody/tr[1]/td[2]/span[1]/input[1]").send_keys(
            end_time)
        time.sleep(1)
    if src_name == 'salelist':
        driver.find_element_by_xpath('/html/body/div[1]/div[2]/table/tbody/tr[1]/td[2]/span[1]/input[1]').clear()
        time.sleep(1)
        driver.find_element_by_xpath('/html/body/div[1]/div[2]/table/tbody/tr[1]/td[2]/span[1]/input[1]').send_keys(
            start_time)
        time.sleep(1)
        driver.find_element_by_xpath('/html/body/div[1]/div[2]/table/tbody/tr[1]/td[2]/span[2]/input[1]').clear()
        time.sleep(1)
        driver.find_element_by_xpath('/html/body/div[1]/div[2]/table/tbody/tr[1]/td[2]/span[2]/input[1]').send_keys(
            end_time)
        time.sleep(1)
        driver.find_element_by_xpath('//*[@id="batchNo"]').click()
    time.sleep(1)
    # //*[@id="search"]/span/span
    driver.find_element_by_link_text("查询").click()
    time.sleep(1)
    driver.find_element_by_link_text("导出").click()
    limit = 10
    finished = False
    while limit > 0:
        limit = limit - 1
        time.sleep(1)
        if is_download_finished(folder):
            limit = 0
            finished = True
        elif limit == 0 and finished == False:
            print('download failed or take too long!')


# 老百姓贝林网载获取 -- 异常，部分按钮点击失败
def hr02(url, pageUsername, pagePassword, folder):
    # 获取驱动器
    driver = getDriver(folder)
    driver.get(url)
    # 找到用户名、密码
    username = driver.find_element_by_name("userId")
    password = driver.find_element_by_name("password")
    time.sleep(2)
    # 给用户名、密码赋值
    username.send_keys(pageUsername)
    password.send_keys(pagePassword)
    # 登录系统
    driver.find_element_by_xpath("//input[@value='登 录']").click()
    time.sleep(3)
    # 找到销售流向
    driver.find_element_by_link_text("销售管理").click()
    time.sleep(3)
    driver.find_element_by_link_text("配送/物流中心出库/返仓").click()
    time.sleep(3)

    # 跳转iframe
    mainframe = driver.find_element_by_id("mainframe")  # //*[@id="mainframe"]
    driver.switch_to.frame(mainframe)
    time.sleep(1)

    # 定位起始时间 input
    driver.find_element_by_xpath('//*[@id="form"]/span[1]/span[1]/input').send_keys(first_day)
    time.sleep(1)

    # 定位结束时间 input
    driver.find_element_by_xpath("//*[@id='form']/span[2]/span[1]/input").send_keys(yesterday)
    time.sleep(3)

    # 定位商品窗口
    actions = ActionChains(driver)
    goods_span = driver.find_element_by_xpath('//*[@id="input_goods_id"]/span/span/span[2]/span')
    actions.move_to_element(to_element=goods_span).click().perform()
    # document.querySelector("#input_goods_id\\$value")  document.querySelector("#input_goods_id\\$text")
    # js = 'document.querySelector("#input_goods_id$value") = "1074091";'

    # js = """function getQueryJson(){
    #         data=form.getData(true,true);
    #         data.param.BEGINDATE="2019-06-18";
    #         data.param.ENDDATE="2019-06-19";
    #         data.param.GOODSID="1074091";
    #         data.param.DEPTID="611";
    #         return data
    #     }"""
    # driver.execute_script(js)
    time.sleep(1)

    # 切换iframe //*[@id="mini-17"]/div/div[2]/div[2]/iframe /srm/sale/MerchandiseQuery.jsp?entity=org.gocom.components.coframe.org.dataset.OrgOrganization&_winid=w2419&_t=343997
    # wait = ui.WebDriverWait(driver, 10)
    # check_box_iframe = wait.until(lambda driver: driver.find_element_by_xpath("//iframe[contains(@src,'/srm/sale/MerchandiseQuery')]"))
    check_box_iframe = driver.find_element_by_xpath('//div[@id="mini-17"]/div/div[2]/div[2]/iframe')
    driver.switch_to.frame(check_box_iframe)
    time.sleep(2)
    # 选定“全选”框，点击   '//*[@id="mini-17$headerCell2$2"]/div/div[1]/input'
    box_element = driver.find_element_by_xpath("//*[@id='mini-17checkall']")
    actions.move_to_element(box_element).click().perform()
    time.sleep(1)

    # 选定“确认”按钮，点击
    submit_button_element = driver.find_element_by_xpath('/html/body/a[1]/span')
    actions.move_to_element(submit_button_element).click().perform()
    # driver.find_element_by_link_text("确认").click()
    time.sleep(1)
    # 选定部门
    dept_element = driver.find_element_by_xpath('//*[@id="input_dept$text"]')
    actions.move_to_element(dept_element).click().perform()
    time.sleep(1)
    dept_element_select = driver.find_element_by_xpath('//*[@id="mini-10$1"]/td[2]')
    actions.move_to_element(dept_element_select).click().perform()
    time.sleep(1)
    # 选定查询按钮并点击
    select_button = driver.find_element_by_xpath('//*[@id="form"]/a[1]/span')
    actions.move_to_element(select_button).click().perform()
    # driver.find_element_by_link_text("查询").click()
    time.sleep(1)
    # 选定导出按钮并点击
    export_button = driver.find_element_by_xpath('//*[@id="form"]/a[3]/span')
    actions.move_to_element(export_button).click().perform()
    # driver.find_element_by_link_text("导出").click()
    limit = 10
    finished = False
    while limit > 0:
        limit = limit - 1
        time.sleep(1)
        if is_download_finished(folder):
            limit = 0
            finished = True
        elif limit == 0 and finished == False:
            print('download failed or take too long!')

    input('Press ENTER to close the automated browser')

    driver.quit()


# 重庆 -- 页面跳转失败，点击操作异常
def hr03(url, pageUsername, pagePassword, folder):
    driver = getDriver(folder)
    driver.get(url)
    # 找到登录按钮，点击跳转到登录页面
    driver.find_element_by_id("login_logout").click()
    time.sleep(2)

    # 找到用户名、密码的输入框//*[@id="myform"]/div[1]/span
    username = driver.find_element_by_xpath('//*[@id="sitename"]')
    password = driver.find_element_by_xpath('//*[@id="myform"]/div[2]/span/input')
    vCode = driver.find_element_by_id("txtcheckCode")
    xp = '//*[@id="checkCode"]'
    time.sleep(2)
    # 打码
    get_image(driver, xp, 75, 120, 7, 7)
    cid, resultCode = yundama.decode(filename, codetype, timeout)
    print(resultCode)

    # 配置用户名、密码
    username.send_keys(pageUsername)
    password.send_keys(pagePassword)
    vCode.send_keys(resultCode)
    # 找到submit按钮 //*[@id="myform"]/div[5]/input[3]
    driver.find_element_by_xpath('//*[@id="myform"]/div[5]/input[3]').click()
    time.sleep(3)
    # 找到“上游客户服务”按钮 并点击 //*[@id="layoutContainers"]/div/div[2]/div/div[5]/section/div/div/div/ul/a[6]/li/h4
    element = driver.find_element_by_xpath(
        '//*[@id="layoutContainers"]/div/div[2]/div/div[5]/section/div/div/div/ul/a[6]/li/span')
    actions = ActionChains(driver)
    # ------------------------------问题位置---------------------------
    actions.move_to_element(to_element=element).click(
        '//*[@id="layoutContainers"]/div/div[2]/div/div[5]/section/div/div/div/ul/a[6]').perform()
    time.sleep(3)
    # //*[@id="ctl00_ContentPlaceHolder1_db__cbfgs_I"]
    limit = 10
    finished = False
    while limit > 0:
        limit = limit - 1
        time.sleep(1)
        if is_download_finished(folder):
            limit = 0
            finished = True
        elif limit == 0 and finished == False:
            print('download failed or take too long!')

    input('Press ENTER to close the automated browser')
    driver.quit()


# 营销广东康美 -- 已完成
def hr04(url, pageUsername, pagePassword, folder):
    # 获取驱动器
    driver = getDriver(folder)
    driver.get(url)
    # 找到登录按钮
    driver.find_element_by_xpath('//*[@id="Map"]/area[5]').click()
    time.sleep(2)
    # 定位用户名、密码、验证码输入框、验证码图片
    username = driver.find_element_by_id("txtUserName")
    password = driver.find_element_by_id("txtPwd")
    vCode = driver.find_element_by_id("txtCode")
    xp = '//*[@id="imgValidate"]'
    # 打码 参数，驱动器、路径、截图位置（left，top，right，bottom）
    get_image(driver, xp, 103, 73, 10, 7)
    cid, resultCode = yundama.decode(filename, codetype, timeout)
    print(resultCode)
    # 输入用户名、密码、验证码
    username.send_keys(pageUsername)
    password.send_keys(pagePassword)
    vCode.send_keys(resultCode)
    # 登录系统
    driver.find_element_by_name("ImageButton1").click()
    time.sleep(3)
    # 找到查询报表位置
    driver.find_element_by_xpath('//*[@id="left_cxbb"]').click()
    time.sleep(3)
    # 定位销售流向位置
    driver.find_element_by_xpath('//*[@id="left_li_xslxcx"]').click()
    time.sleep(3)
    # 设定起始时间
    driver.find_element_by_xpath('//*[@id="txtBeginTime"]').clear()
    driver.find_element_by_xpath('//*[@id="txtBeginTime"]').send_keys(first_day)
    # 设定结束时间
    driver.find_element_by_xpath('//*[@id="txtEndTime"]').clear()
    driver.find_element_by_xpath('//*[@id="txtEndTime"]').send_keys(yesterday)
    # 定位查询按钮
    driver.find_element_by_xpath('//*[@id="Button2"]').click()
    time.sleep(1)
    # 定位导出按钮
    driver.find_element_by_xpath('//*[@id="btnSearch"]').click()
    time.sleep(1)
    actions = ActionChains(driver)
    try:
        actions.send_keys(Keys.ENTER).perform()
    except:
        print("页面弹框，属正常现象，程序继续运行。")
    time.sleep(1)
    limit = 10
    finished = False
    while limit > 0:
        limit = limit - 1
        time.sleep(1)
        if is_download_finished(folder):
            limit = 0
            finished = True
        elif limit == 0 and finished == False:
            print('download failed or take too long!')
    input('Press ENTER to close the automated browser')
    driver.quit()


# 营销广东康美 -- 已完成
def hr05(url, pageUsername, pagePassword, folder):
    # 获取驱动器
    driver = getDriver(folder)
    driver.get(url)
    # 定位用户名、密码、验证码输入框、验证码图片
    username = driver.find_element_by_id("txtUserName")
    password = driver.find_element_by_id("txtPassword")
    # 输入用户名、密码、验证码
    username.clear()
    username.send_keys(pageUsername)
    password.send_keys(pagePassword)
    # 登录系统
    driver.find_element_by_link_text('登录').click()
    time.sleep(3)
    # 找到流向查询位置
    driver.find_element_by_link_text('流向查询').click()
    time.sleep(3)
    salelist = driver.find_element_by_xpath('//iframe[contains(@src,"DataFlowDetails")]')
    driver.switch_to.frame(salelist)
    # 设定起始时间
    driver.find_element_by_xpath('//*[@id="txtBeginDate"]').clear()
    driver.find_element_by_xpath('//*[@id="txtBeginDate"]').send_keys(first_day)
    # 设定结束时间
    driver.find_element_by_xpath('//*[@id="txtEndDate"]').clear()
    driver.find_element_by_xpath('//*[@id="txtEndDate"]').send_keys(yesterday)
    # 定位查询按钮
    driver.find_element_by_link_text('查询').click()
    time.sleep(1)
    # 定位导出按钮
    driver.find_element_by_link_text('导出').click()
    time.sleep(1)
    limit = 10
    finished = False
    while limit > 0:
        limit = limit - 1
        time.sleep(1)
        if is_download_finished(folder):
            limit = 0
            finished = True
        elif limit == 0 and finished == False:
            print('download failed or take too long!')
    input('Press ENTER to close the automated browser')
    driver.quit()

with open('wdt.csv', newline='', encoding='utf-8') as csvfile:
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
            if item.endswith('.part' or '.crdownload'):
                os.remove(os.path.join(bpath, item))

        globals()[workerID](url, pageUsername, pagePassword, folder)
