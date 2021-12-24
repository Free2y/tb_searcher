import os.path
import time

from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QFileDialog

from start_ui import Ui_Dialog
from tb_searcher import TBSearcher


class MyWindow(QtWidgets.QWidget, Ui_Dialog):
    def __init__(self):
        super(MyWindow, self).__init__()
        self.setupUi(self)
        self.addDriverButton.clicked.connect(self.read_file)
        self.selectDirButton.clicked.connect(self.write_folder)
        self.pushButtonStart.clicked.connect(self.process_start)

    def read_file(self):
        filename, filetype = QFileDialog.getOpenFileName(self, '选取文件', 'C:/', 'All Files(*);;Exe Files(*.exe)')
        self.lineEditDrvier.setText(filename)

    def write_folder(self):
        foldername = QFileDialog.getExistingDirectory(self, '选取文件件', 'C:/')
        self.lineEditDir.setText(foldername)

    def process_start(self):
        try:
            can_start = True
            chromedriver_path = self.lineEditDrvier.text()
            if chromedriver_path == '':
                can_start = False
                print('请选择chromedriver所在路径')
            output_basedir = self.lineEditDir.text()
            if output_basedir == '':
                can_start = False
                print('请选择保存宝贝的目录')
            output_name = self.lineEditOutname.text()
            username = self.lineEditUsername.text()
            if username == '':
                can_start = False
                print('请输入用户名')
            password = self.lineEditPassword.text()
            if password == '':
                can_start = False
                print('请输入密码')
            key_word = self.lineEditKeyword.text()
            if key_word == '':
                can_start = False
                print('请输入需要搜索的宝贝关键词')

            use_sale_desc_flag = False if self.checkBoxSaleSort.checkState() == 0 else True
            use_tmall_flag = False if self.checkBoxTmall.checkState() == 0 else True
            browser_hide_flag = True if self.checkBoxHide.checkState() == 0 else False

            num = self.lineEditPageNum.text()
            try:
                if num == '':
                    num = 10
                num = int(num)
            except Exception:
                can_start = False
                print('请输入数字')

            if output_name == '':
                output_name = key_word + '_' + time.strftime("%Y-%m-%d-%H-%M-%S", time.localtime())

            if can_start:
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
                tbs = TBSearcher(chromedriver_path, output_basedir, output_name, username, password, key_word,
                                 use_sale_desc_flag,
                                 use_tmall_flag, browser_hide_flag)

                tbs.start_search(num)
        except Exception as e:
            print(e)


if __name__ == '__main__':
    import sys

    app = QtWidgets.QApplication(sys.argv)
    ui = MyWindow()
    ui.show()
    sys.exit(app.exec_())
