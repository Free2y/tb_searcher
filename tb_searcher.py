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
import getpass

from tools.freezy_selenium_spider import FreezySeleniumSpider


class TBSearcher:

    def __init__(self, chromedriver_path, output_basedir, output_name, username, password, keyword, user_sale_sort,
                 user_tmall, mongo_table='default'):
        self.BASE_DIR = output_basedir
        self.output_name = output_name
        self.USERNAME = username
        self.PASSWORD = password
        self.KEYWORD = keyword
        self.MONGO_TABLE = mongo_table

        chrome_options = Options()
        chrome_options.add_argument('--headless')
        chrome_options.add_argument('--disable-gpu')
        self.chromedriver_path = chromedriver_path
        self.browser = FreezySeleniumSpider(path=self.chromedriver_path, options=chrome_options)
        self.wait = WebDriverWait(self.browser, 15)
        self.db_client = pymongo.MongoClient('localhost', 27017)
        self.db = self.db_client['search_taobao']
        if self.MONGO_TABLE in self.db.list_collection_names():
            self.db.drop_collection(self.db[self.MONGO_TABLE])
        self.table = self.db[self.MONGO_TABLE]
        self.TOTAL_SUM = 0
        self.USE_SALE_DESC = user_sale_sort
        self.USE_TMALL = user_tmall

    def search_page(self):
        print('正在搜索')
        try:
            self.browser.get('https://www.taobao.com')
            input = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#q"))
            )
            submit = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, '#J_TSearchForm > div.search-button > button')))
            input.send_keys(self.KEYWORD)
            submit.click()
            self.check_login()
            if self.USE_TMALL:
                sort_tmall_a = self.wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, '#main a#tabFilterMall')))
                sort_tmall_a.click()
                time.sleep(1)
            if self.USE_SALE_DESC:
                sort_sale_a = self.wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, '#mainsrp-sortbar a[data-value="sale-desc"]')))
                sort_sale_a.click()
                time.sleep(1)
            total = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '#mainsrp-pager > div > div > div > div.total')))
            print(total.text)
            self.get_products()
            return total.text

        except TimeoutException:
            return self.search_page()

    def next_page(self, page_number):
        print('翻页中', page_number)
        try:
            input = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '#mainsrp-pager > div > div > div > div.form > input'))
            )
            submit = self.wait.until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, '#mainsrp-pager > div > div > div > div.form > span.btn.J_Submit')))
            input.clear()
            input.send_keys(page_number)
            submit.click()
            self.check_login()
            self.wait.until(EC.text_to_be_present_in_element(
                (By.CSS_SELECTOR, '#mainsrp-pager > div > div > div > ul > li.item.active > span'), str(page_number)))
            self.get_products()
        except TimeoutException:
            self.next_page(page_number)

    def get_products(self):

        print('开始加载商品')
        self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#mainsrp-itemlist .items .item')))
        html = self.browser.page_source
        print('开始解析商品')
        doc = pq(html)
        items = doc('#mainsrp-itemlist .items .item').items()
        index = 0
        for item in items:
            index = index + 1
            image_url = item.find('.J_ItemPic.img').attr('src')
            if not image_url.startswith('http'):
                image_url = item.find('.J_ItemPic.img').attr('data-ks-lazyload')
            image_file = os.path.join(self.BASE_DIR, self.output_name, str(self.TOTAL_SUM + index) + '.png')
            self.download_image(image_url, image_file)
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

            self.save_to_mongo(product)
        self.TOTAL_SUM = self.TOTAL_SUM + index
        print('已经爬取' + str(self.TOTAL_SUM) + '件商品')

    def check_login(self):
        if self.browser.current_url.find('login.taobao.com') != -1:
            input_username = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '#fm-login-id'))
            )
            input_password = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '#fm-login-password'))
            )
            submit = self.wait.until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, '.fm-button.fm-submit.password-login')))
            input_username.send_keys(self.USERNAME)
            input_password.send_keys(self.PASSWORD)
            submit.click()

    def download_image(self, url, filename):
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

    def save_to_mongo(self, result):
        try:
            if self.table.insert_one(result):
                print('存储成功', result)
        except Exception as e:
            print('存储失败', e)

    def mongoexport_to_csv(self, book):
        print('开始导出数据到CSV')
        csv_name = os.path.join(self.BASE_DIR, self.output_name, self.output_name + '.csv')
        data = pd.DataFrame(list(self.table.find()))
        data.to_csv(csv_name, encoding='utf-8', index=False)
        print('导出数据到' + csv_name + '完成')
        print('开始执行数据格式化Excel')
        self.format_excel(csv_name, book)
        print('成功生成Xlsx文档')

    def format_excel(self, csv_name, book):
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

    def start_search(self, max_page):
        try:
            output_path = os.path.join(self.BASE_DIR, self.output_name)
            if not os.path.exists(output_path):
                os.makedirs(output_path)
            total = self.search_page()
            total = int(re.compile('(\d+)').search(total).group(1))
            for i in range(2, min(total + 1, max_page + 1)):
                time.sleep(3)
                self.next_page(i)
            book = xlsxwriter.Workbook(os.path.join(self.BASE_DIR, self.output_name, self.output_name + '.xlsx'),
                                       {'nan_inf_to_errors': True})
            self.mongoexport_to_csv(book)
        except Exception as e:
            print('爬取错误', e)
        finally:
            self.browser.close()
            self.db_client.close()
            book.close()
            print('本次任务结束:总共爬取了' + str(self.TOTAL_SUM) + '件' + self.KEYWORD + '类型商品')


if __name__ == '__main__':

    while True:
        chromedriver_path = input('请输入chormedriver可执行程序路径:')
        if chromedriver_path != '' and os.path.exists(chromedriver_path):
            pass
        elif chromedriver_path == '':
            chromedriver_path = r'D:\Miniconda3\chromedriver.exe'
        else:
            print('请确认可执行程序路径是否正确')
            continue
        break

    while True:
        key_word = input('请输入要搜索的关键词:')
        if key_word == '':
            continue
        break

    while True:
        username = input('请输入您的用户名:')
        if username == '':
            continue
        break

    password = getpass.getpass('请输入您的密码:')
    # password = input('请输入您的密码:')

    while True:
        num = input('请输出要爬取的最大页数(默认10):')
        try:
            if num == '':
                num = 10
            num = int(num)
        except Exception:
            print('请正确输入数字')
            continue
        break

    use_sale_desc = input('是否使用销量排序Y/N(默认综合排序):')
    if use_sale_desc == 'Y':
        use_sale_desc_flag = True
    else:
        use_sale_desc_flag = False

    use_tmall = input('是否使用天猫搜索Y/N(默认所有宝贝):')
    if use_tmall == 'Y':
        use_tmall_flag = True
    else:
        use_tmall_flag = False

    while True:
        output_basedir = input('请选择输出根目录:')
        if not os.path.isdir(output_basedir):
            print('请输入正确的目录')
            continue
        break
    output_name = input('请选择输出目录名称:')
    if output_name == '':
        output_name = key_word + '_' + time.strftime("%Y-%m-%d-%H-%M-%S", time.localtime())

    print('您使用的chromedriver.exe路径为:', chromedriver_path)
    print('您使用的搜索关键词为:', key_word)
    if use_sale_desc_flag:
        sort_info = '您将使用销量排序'
    else:
        sort_info = '您将使用默认综合排序'
    if use_tmall_flag:
        tmall_info = '您将仅搜索天猫店铺'
    else:
        tmall_info = '您将搜索全部店铺'
    print(tmall_info)
    print(sort_info)
    print('您将一共爬取' + str(num) + '页数据')
    print('您的资料将存在' + os.path.join(output_basedir, output_name) + '  目录下')

    tbs = TBSearcher(chromedriver_path, output_basedir, output_name, username, password, key_word, use_sale_desc_flag,
                     use_tmall_flag)

    tbs.start_search(num)
