# coding:utf-8

"""
@describe: 基于selenium版本进一步封装 只针对于谷歌浏览器
"""
import os
from selenium.webdriver.chrome.webdriver import WebDriver

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
STEALTH_JS_FILE = 'stealth.min.js'


class FreezySeleniumSpider(WebDriver):
    """基于selenium进一步封装"""

    def __init__(self, path, options=None, crack_file=os.path.join(BASE_DIR, STEALTH_JS_FILE), *args, **kwargs):
        """
        初始化
        :param path: str selenium驱动路径
        :param params: list driver 附加参数
        :param args: tuple
        :param kwargs:
        """
        self.__path = path
        self.__crack_file = crack_file
        self.__options = options
        super(FreezySeleniumSpider, self).__init__(executable_path=self.__path, options=self.__options, *args, **kwargs)

        self.crack_by_js(self.__crack_file)

    def crack_by_js(self, crack_file):
        with open(crack_file) as f:
            js = f.read()
        self.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
            "source": js
        })
