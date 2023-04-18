# coding: utf-8
# +-------------------------------------------------------------------
# | Project     : zk-sceea
# | Author      : 今夕何夕
# | QQ/Email    : 184566608<qingyueheji@qq.com>
# | Time        : 2020-07-07 01:20
# | Describe    : zk-sceea
# +-------------------------------------------------------------------
import os

import resource_rc
from copy import deepcopy
from PyQt5.QtCore import Qt
from libs.Notification import NotificationWindow
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog, QStatusBar, QToolBar, QAction, QTableWidgetItem
from libs.excel import ExcelRead, ExcelWrite
from designer.ui_main_window import Ui_MainWindow
from libs.worker import WorkerEnterCourse


class Window(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super(Window, self).__init__()
        self.setupUi(self)
        self.statusBar = QStatusBar(self)
        self.toolbar = QToolBar(self)
        self.initUi()
        self.showMaximized()
        self._pos = {}
        self._data = []

    def initUi(self):
        self.setStatusBar(self.statusBar)

        self.toolbar_action("open_excel", "svg/export.svg", "导入数据")
        self.toolbar_action("excel_template", "svg/download.svg", "下载数据模板")
        self.toolbar_action("execute", "svg/run.svg", "开始")
        self.toolbar_action("stop", "svg/stop.svg", "停止")
        self.toolbar_action("clear", "svg/clear.svg", "清空表")

        self.findChild(QAction, "action_clicked_execute").setDisabled(True)
        self.findChild(QAction, "action_clicked_stop").setDisabled(True)
        self.findChild(QAction, "action_clicked_clear").setDisabled(True)
        # 图形对象点击触发自定义槽函数
        self.toolbar.actionTriggered[QAction].connect(self.tool_btn_pressed)

        self.tableWidget.setColumnWidth(0, 80)
        self.tableWidget.setColumnWidth(1, 140)
        self.tableWidget.setColumnWidth(2, 80)
        self.tableWidget.setColumnWidth(3, 420)
        self.tableWidget.setColumnWidth(4, 100)
        self.tableWidget.setColumnWidth(5, 100)
        self.tableWidget.setColumnWidth(6, 80)

        self.tableWidget.hide()
        self.resize(1165, 673)

    def action_clicked_clear(self, action):
        """
        清除数据表格
        :param action:
        :return:
        """
        self._pos.clear()
        self._data.clear()
        self.tableWidget.clear()
        self.tableWidget.setRowCount(0)
        self.tableWidget.setHorizontalHeaderLabels(["姓名", "身份证", "类型", "科目", "城市", "区域", "指定区域", "状态"])
        action.setDisabled(True)
        self.tableWidget.hide()
        self.findChild(QAction, "action_clicked_clear").setDisabled(True)
        self.findChild(QAction, "action_clicked_execute").setDisabled(True)
        self.findChild(QAction, "action_clicked_open_excel").setDisabled(False)
        self.findChild(QAction, "action_clicked_excel_template").setDisabled(False)

    def action_clicked_stop(self, action=None):
        """
        停止线程
        :param action:
        :return:
        """
        for id_card, _ in self._pos.items():
            worker = self.__getattribute__(f"_worker_{id_card}")
            worker.stop = True
        self.findChild(QAction, "action_clicked_stop").setDisabled(True)
        self.findChild(QAction, "action_clicked_clear").setDisabled(False)
        self.findChild(QAction, "action_clicked_execute").setDisabled(False)

    def action_clicked_execute(self, action):
        """
        执行
        :param action:
        :return:
        """
        for k, v in self._pos.items():
            v['status'].setText("开始处理...")
        data_list = deepcopy(self._data)
        for item in data_list:
            id_card = item.get("id_card")
            item["pwd"] = id_card[-6:]
            setattr(self, f"_worker_{id_card}", WorkerEnterCourse(**item))
            work = getattr(self, f"_worker_{id_card}")
            work.msg.connect(self.loop_row_log)
            work.start()
        action.setDisabled(True)
        self.findChild(QAction, "action_clicked_stop").setDisabled(False)
        self.findChild(QAction, "action_clicked_clear").setDisabled(True)

    def loop_row_log(self, id_card, log, status):
        """
        往表格写入日志
        :param status:
        :param id_card:
        :param log:
        :return:
        """
        row = self._pos.get(id_card)
        if not row: return
        row.get("status").setText(log)
        # work: WorkerEnterCourse = getattr(self, f"_worker_{id_card}")
        # if work.isFinished() and status and not log.startswith("网络请求失败"):
        #     work.stop = False
        #     work.start()

    def toolbar_action(self, func, icon, text, b=False):
        """
        添加工具栏
        :param func:
        :param icon:
        :param text:
        :param b:
        :return:
        """
        action = QAction(QIcon(f":/{icon}"), text, self)
        action.setDisabled(b)
        action.setStatusTip(text)
        action.setToolTip(text)
        action.setObjectName(f'action_clicked_{func}')
        self.toolbar.addAction(action)
        self.addToolBar(Qt.TopToolBarArea, self.toolbar)

    def tool_btn_pressed(self, action: QAction):
        """
        工具栏点击事件
        :param action:
        :return:
        """
        try:
            clicked = action.objectName()
            self.__getattribute__(clicked)(action)
        except:
            pass

    def insert_row(self, row, col, txt):
        """
        表格行插入
        :param row:
        :param col:
        :param txt:
        :return:
        """
        item = QTableWidgetItem(txt)
        self.tableWidget.setItem(row, col, item)
        return item

    def action_clicked_open_excel(self, action: QAction):
        """
        打开excel
        :param action:
        :return:
        """
        try:
            temp = os.path.join(os.environ["TEMP"], "zk-sceea-cache-read")
            path = os.path.join(os.environ['USERPROFILE'], "Desktop")
            if os.path.exists(temp):
                with open(temp, 'r', encoding="utf8") as f:
                    path = f.read()
            excel_file, _ = QFileDialog.getOpenFileName(self, "选取Excel数据集文件", path, "Excel工作簿文件 (*.xlsx)")
            if not excel_file or os.path.splitext(excel_file)[-1] != ".xlsx":
                return
            with open(temp, 'w', encoding="utf8") as f:
                f.write(os.path.dirname(excel_file))

            field = {'default': ['name', 'id_card', 'ctype', 'cname', 'addr1', 'addr2', 'area']}
            excel = ExcelRead(excel_file, field=field)
            _, _, temp_data = excel.data
            data_dict = {item.get('id_card').strip(): item for item in temp_data}
            data_dict.pop('', None)
            data = [item for _, item in data_dict.items()]
            self._data.extend(data)
            count = self.tableWidget.rowCount()
            self.tableWidget.setRowCount(count + len(data))
            for i, item in enumerate(data, count):
                self._pos[item['id_card']] = {
                    'name':    self.insert_row(i, 0, item['name']),
                    'id_card': self.insert_row(i, 1, item['id_card']),
                    'ctype':   self.insert_row(i, 2, item['ctype']),
                    'cname':   self.insert_row(i, 3, item['cname']),
                    'addr1':   self.insert_row(i, 4, item['addr1']),
                    'addr2':   self.insert_row(i, 5, item['addr2']),
                    'area':    self.insert_row(i, 6, item['area']),
                    'status':  self.insert_row(i, 7, "待处理")
                }
            if len(data) > 0:
                action.setDisabled(True)
                self.findChild(QAction, "action_clicked_excel_template").setDisabled(True)
                self.findChild(QAction, "action_clicked_clear").setDisabled(False)
                self.findChild(QAction, "action_clicked_execute").setDisabled(False)
            self.tableWidget.show()
        except:
            NotificationWindow.warning("数据格式错误警告", "Excel数据不符合规则,应该如下:\n\n姓名|身份证号码|报考类型|报考科目|市|区|指定区域\n")

    def action_clicked_excel_template(self, action):
        """导出模板"""
        try:
            temp = os.path.join(os.environ["TEMP"], "zk-sceea-cache-save")
            path = os.path.join(os.environ['USERPROFILE'], "Desktop")
            if os.path.exists(temp):
                with open(temp, 'r', encoding="utf8") as f:
                    path = f.read()
            fname, _ = QFileDialog.getSaveFileName(self, '选取Excel数据集文件', path, "Excel工作簿文件 (*.xlsx)")
            if not fname:
                return
            with open(temp, 'w', encoding="utf8") as f:
                f.write(os.path.dirname(fname))

            header = {"抢位名单": {
                '考生姓名':    '必填\n说明:\n考试考生姓名,这个可以填写错误,但是不能不填(必填项)',
                '身份证号':    '必填\n说明:\n考试身份证号码,不可填写错误(必填项)',
                '报考类型':    '选填\n说明:\n报考类型(社会型/应用型)(必填项)',
                '报考科目':    '选填\n说明:\n要报考的科目,使用"、"分隔(必填项)',
                '报考地点(市)': '选填\n说明:\n市级(成都市),必须填写完整',
                '报考科目(区)': '选填\n说明:\n区级(青羊区),必须填写完整',
                '指定区域':    '选填\n说明:\n是或否,默认否',
            }}

            write = ExcelWrite(filename=fname, header=header)  # 实例化
            write.create_sheet("抢位名单", [
                {
                    'name':    '张三',
                    'id_card': '51122250040464385X',
                    'ctype':   '应用型',
                    'cname':   '信息系统设计与分析、计算机信息检索、Windows及应用、英语(二)',
                    'addr1':   '成都市',
                    'addr2':   '青羊区',
                    'is_area': '否',
                }
            ])
            write.save()

            def open_path():
                os.system(f"explorer {os.path.dirname(fname)}".replace('/', '\\'))

            NotificationWindow.success("文件保存提示", "文件保存成功，点击打开文件路径！", callback=open_path)
        except:
            NotificationWindow.warning("未知错误", "文件保存失败，请重试！")


if __name__ == '__main__':
    import sys

    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon(":/icon.ico"))
    QApplication.setStyle("Fusion")
    window = Window()
    window.setWindowTitle("四川省考试院自动报考抢位系统-送你一个小星星 V2.0")
    window.show()
    sys.exit(app.exec_())
