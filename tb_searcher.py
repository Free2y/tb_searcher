import os
import re
import time
import urllib
import xlsxwriter
import pymongo
from PIL import Image
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from pyquery import PyQuery as pq
import pandas as pd

from tools.freezy_selenium_spider import FreezySeleniumSpider

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
output_name = ''
USERNAME = ''
PASSWORD = ''
KEYWORD = 'default'
MONGO_TABLE = 'default'

chrome_options = Options()
chrome_options.add_argument('--headless')
chrome_options.add_argument('--disable-gpu')
chromedriver_path = r'D:\Miniconda3\chromedriver.exe'
browser = FreezySeleniumSpider(path=chromedriver_path, options=chrome_options)
wait = WebDriverWait(browser, 15)
db_client = pymongo.MongoClient('localhost', 27017)
db = db_client['search_taobao']
if MONGO_TABLE in db.list_collection_names():
    db.drop_collection(db[MONGO_TABLE])
table = db[MONGO_TABLE]
TOTAL_SUM = 0
USE_SALE_DESC = False
USE_TMALL = False


def search_page():
    print('正在搜索')
    try:
        browser.get('https://www.taobao.com')
        input = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#q"))
        )
        submit = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, '#J_TSearchForm > div.search-button > button')))
        input.send_keys(KEYWORD)
        submit.click()
        check_login()
        if USE_TMALL:
            sort_tmall_a = wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '#main a#tabFilterMall')))
            sort_tmall_a.click()
            time.sleep(1)
        if USE_SALE_DESC:
            sort_sale_a = wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '#mainsrp-sortbar a[data-value="sale-desc"]')))
            sort_sale_a.click()
            time.sleep(1)
        total = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '#mainsrp-pager > div > div > div > div.total')))
        print(total.text)
        get_products()
        return total.text

    except TimeoutException:
        return search_page()


def next_page(page_number):
    print('翻页中', page_number)
    try:
        input = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '#mainsrp-pager > div > div > div > div.form > input'))
        )
        submit = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, '#mainsrp-pager > div > div > div > div.form > span.btn.J_Submit')))
        input.clear()
        input.send_keys(page_number)
        submit.click()
        check_login()
        wait.until(EC.text_to_be_present_in_element(
            (By.CSS_SELECTOR, '#mainsrp-pager > div > div > div > ul > li.item.active > span'), str(page_number)))
        get_products()
    except TimeoutException:
        next_page(page_number)


def get_products():
    global TOTAL_SUM
    print('开始加载商品')
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#mainsrp-itemlist .items .item')))
    html = browser.page_source
    print('开始解析商品')
    doc = pq(html)
    items = doc('#mainsrp-itemlist .items .item').items()
    index = 0
    for item in items:
        index = index + 1
        image_url = item.find('.J_ItemPic.img').attr('src')
        if not image_url.startswith('http'):
            image_url = item.find('.J_ItemPic.img').attr('data-ks-lazyload')
        image_file = os.path.join(BASE_DIR, output_name, str(TOTAL_SUM + index) + '.png')
        download_image(image_url, image_file)
        link_url = item.find('.J_ClickStat').attr('href')
        if not link_url.startswith('https:'):
            link_url = 'https:' + link_url
        product = {
            'price': item.find('.price').text(),
            'deal': item.find('.deal-cnt').text(),
            'title': item.find('.title').text(),
            'shop': item.find('.shop').text(),
            'location': item.find('.location').text(),
            'image_url': image_url,
            'link_url': link_url,
            'image_file': image_file
        }

        save_to_mongo(product)
    TOTAL_SUM = TOTAL_SUM + index
    print('已经爬取' + str(TOTAL_SUM) + '件商品')


def check_login():
    if browser.current_url.find('login.taobao.com') != -1:
        input_username = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '#fm-login-id'))
        )
        input_password = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '#fm-login-password'))
        )
        submit = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, '.fm-button.fm-submit.password-login')))
        input_username.send_keys(USERNAME)
        input_password.send_keys(PASSWORD)
        submit.click()


def download_image(url, filename):
    try:
        suff_word = '_180x180.jpg'
        loc = url.find(suff_word)
        if loc != -1:
            url = url[:loc]
        urllib.request.urlretrieve(url, filename=filename)
        img = Image.open(filename)
        img.load()
        img.save(filename)
    except IOError as e:
        print("IOE ERROR:", e)
    except Exception as e:
        print("Exception:", e)


def save_to_mongo(result):
    try:
        if table.insert_one(result):
            print('存储成功', result)
    except Exception as e:
        print('存储失败', e)


def mongoexport_to_csv(book):
    print('开始导出数据到CSV')
    csv_name = os.path.join(BASE_DIR, output_name, output_name + '.csv')
    data = pd.DataFrame(list(table.find()))
    data.to_csv(csv_name, encoding='utf-8', index=False)
    print('导出数据到' + csv_name + '完成')
    print('开始执行数据格式化Excel')
    format_excel(csv_name, book)
    print('成功生成Xlsx文档')


def format_excel(csv_name, book):
    df = pd.read_csv(csv_name)
    # print(df.columns)
    my_format = book.add_format({
        'bold': True,  # 字体加粗
        'align': 'center',  # 水平位置设置：居中
        'valign': 'vcenter',  # 垂直位置设置，居中
        'font_size': 11,  # '字体大小设置'
        'text_wrap': 1
    })
    sheet = book.add_worksheet("formated")
    # 定义一下两列的name,再把要匹配的昵称填充进去。
    sheet.write("A1", "商品id", my_format)
    sheet.write("B1", "商品价格", my_format)
    sheet.write("C1", "商品销量", my_format)
    sheet.write("D1", "商品描述", my_format)
    sheet.write("E1", "商品所属店铺", my_format)
    sheet.write("F1", "商品所在地", my_format)
    sheet.write("G1", "商品图片", my_format)
    sheet.write_column(1, 0, df._id.values.tolist(), my_format)  # 昵称放在第一列
    sheet.write_column(1, 1, df.price.values.tolist(), my_format)
    sheet.write_column(1, 2, df.deal.values.tolist(), my_format)
    sheet.write_column(1, 3, df.title.values.tolist(), my_format)
    sheet.write_column(1, 4, df.shop.values.tolist(), my_format)
    sheet.write_column(1, 5, df.location.values.tolist(), my_format)
    sheet.write_column(1, 6, df.link_url.values.tolist())
    image_width = 160
    image_height = 180
    cell_width = 20
    cell_height = 180
    sheet.set_column('A:G', cell_width)
    for i in range(len(df.image_file.values.tolist())):
        image_path = df.image_file.values.tolist()[i]
        if os.path.exists(image_path):
            sheet.set_row(i + 1, cell_height)  # 设置行高
            width_scale = image_width / (
                Image.open(image_path).size[0]
            )
            height_scale = image_height / (
                Image.open(image_path).size[1]
            )
            sheet.insert_image(
                "G{}".format(i + 2),
                image_path,
                {"x_scale": width_scale, "y_scale": height_scale, "x_offset": 15, "y_offset": 20},
            )  # 设置一下x_offset和y_offset让图片尽量居中


def main(max_page):
    try:
        global output_name
        output_name = KEYWORD + '_' + time.strftime("%Y-%m-%d-%H-%M-%S", time.localtime())
        output_path = os.path.join(BASE_DIR, output_name)
        if not os.path.exists(output_path):
            os.makedirs(output_path)
        total = search_page()
        total = int(re.compile('(\d+)').search(total).group(1))
        for i in range(2, min(total + 1, max_page + 1)):
            time.sleep(3)
            next_page(i)
        book = xlsxwriter.Workbook(os.path.join(BASE_DIR, output_name, output_name + '.xlsx'),
                                   {'nan_inf_to_errors': True})
        mongoexport_to_csv(book)
    except Exception as e:
        print('爬取错误', e)
    finally:
        browser.close()
        db_client.close()
        book.close()
        print('本次任务结束:总共爬取了' + str(TOTAL_SUM) + '件' + KEYWORD + '类型商品')


if __name__ == '__main__':
    key_word = ''
    num = ''
    username = ''
    password = ''
    table_name = ''
    while True:
        if key_word == '':
            key_word = input('请输入要搜索的关键词:')
        if key_word == '':
            continue
        KEYWORD = key_word
        if num == '':
            num = input('请输出要爬取的最大页数(默认10):')
            try:
                if num == '':
                    num = 10
                num = int(num)
            except Exception:
                num = ''
                continue
        if username == '':
            username = input('请输入您的用户名:')
            if username != '':
                USERNAME = username
            else:
                username = USERNAME
        if password == '':
            password = input('请输入您的密码:')
            if password != '':
                PASSWORD = password
            else:
                password = PASSWORD

        use_sale_desc = input('是否使用销量排序Y/N(默认综合排序):')
        if use_sale_desc == 'Y':
            USE_SALE_DESC = True
        elif use_sale_desc in ['N', '']:
            USE_SALE_DESC = False
        else:
            continue

        use_tmall_desc = input('是否使用天猫搜索Y/N(默认所有宝贝):')
        if use_tmall_desc == 'Y':
            USE_TMALL = True
        elif use_tmall_desc in ['N', '']:
            USE_TMALL = False
        else:
            continue
        # if table_name == '':
        #     table_name = input('请输入您想要保存的表名(非中文):')
        #     if table_name != '':
        #         if table_name in db.list_collection_names():
        #             db.drop_collection(db[table_name])
        #         table = db[table_name]
        #     else:
        #         continue
        main(num)
        break
