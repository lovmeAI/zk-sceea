import json
import time
import traceback

import requests
import cv2
import numpy as np
from PyQt5.QtCore import QThread, pyqtSignal
from lxml import etree

from libs.utils import bk_done

requests.packages.urllib3.disable_warnings()

headers = {
    "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
    "User-Agent":   "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36"
}

samples = np.loadtxt('generalsamples.data', np.float32)
responses = np.loadtxt('generalresponses.data', np.float32)
responses = responses.reshape((responses.size, 1))
model = cv2.ml.KNearest_create()
model.train(samples, cv2.ml.ROW_SAMPLE, responses)


def city(name):
    with open("city.json", "r", encoding="UTF-8") as f:
        data = json.load(f)
        return data.get(name)


class WorkerEnterCourse(QThread):
    msg = pyqtSignal(str, str, bool)  # 值变化信号

    def __init__(self, name, id_card, ctype, pwd, cname, addr1, addr2, area):
        """
        线程抢位置
        :param name: 姓名
        :param id_card: 身份证号码
        :param ctype: 考试类型
        :param pwd: 密码(默认为身份证后六位)
        :param cname: 报考科目名字，不是科目代码
        :param addr1: 地区1(市级，如成都市，绵阳市等)
        :param addr2: 地区2(区/县，如成华区，简阳市等)
        :param area: 是都报考指定区域，是则只报考指定的(区/县)，否则报考指定(市)级下的任意一个区域
        """
        super(WorkerEnterCourse, self).__init__()
        self._username = name
        self._name = id_card
        self._pwd = pwd
        # 社会型 0，应用型 1
        self._ctype = '1' if ctype == '应用型' else '0'
        self._app_type = ctype
        self._addr1 = addr1
        self._addr2 = addr2
        self._course_list = cname.split('、')
        self._login_count = 0
        self._get_addr_count = 0
        self._get_bk_count = 0
        self._times = {}
        self._yx_course = []
        self._cookies = None
        self._enter = None
        self._area = True if area == '是' else False
        self.stop = False

    def run(self) -> None:
        # 获取 cookie
        self._login_cookies()
        if not self._cookies:
            return
        # 获取页面中一些要提交报考的参数
        try:
            self._enter = self._get_view()
            if not self._enter:
                return
            self._get_course()
            self._get_query_address()
        except:
            print(traceback.format_exc())

    def _post(self, url, data):
        try:
            return requests.post(
                url=url,
                data=data,
                headers=headers,
                verify=False,
                cookies=self._cookies
            )
        except requests.exceptions.ConnectionError:
            self._emit_msg("网络请求失败，请检查网络......")
        except:
            print(traceback.format_exc())

    def _get(self, url):
        try:
            return requests.get(url=url, headers=headers, verify=False, cookies=self._cookies)
        except requests.exceptions.ConnectionError:
            self._emit_msg("网络请求失败，请检查网络......")
        except:
            print(traceback.format_exc())

    def _get_view(self, zkz=None):
        """
        获取页面上的属性
        :return:
        """
        e = None
        xxbm = None
        zkzh = None
        jf_log = ""
        while True:
            if self.stop:
                self._emit_msg("停止获取页面上的属性......")
                break
            if not zkz:
                view = self._get("https://zk.sceea.cn/RegExam/switchPage?resourceId=view")
            else:
                view = self._get(f"https://zk.sceea.cn/RegExam/switchPage?resourceId=view&zkz={zkz}")
            if not view:
                return
            e = etree.HTML(view.text)
            zkzh = e.xpath(f"//div[@class='divbox']/input[@kslb='{self._ctype}']/@value")
            if len(zkzh) <= 0:
                self._emit_msg("准考证号获取失败......")
                continue
            kslb = e.xpath(f"//div[@class='divbox'][1]/input[@checked]/@kslb")
            if len(kslb) > 0 and kslb[0] != self._ctype:
                zkz = zkzh[0]
                self._emit_msg(f"切换应用类型为{self._app_type}......")
                continue
            else:
                zkz = zkzh[0]
            if zkz:
                jf_view = self._get(f"https://zk.sceea.cn/RegExam/search?resourceId=view&zkz={zkz}")
                e1 = etree.HTML(jf_view.text)
                if '已缴费' in [item.replace('\r\n', '') for item in e1.xpath("//table[@id='tableYXCourse']//td[9]/text()")]:
                    jf_log = "(已缴费)"
                    break
            self._emit_msg("获取准考证号码")
            xxbm = e.xpath(f"//div[@class='divbox']/input[@kslb='{self._ctype}']/@xxbm")
            if len(xxbm) <= 0:
                self._emit_msg("准考证号码获取异常")
                continue
            else:
                break
        if self.stop:
            return
        if len(jf_log) > 0:
            self._emit_msg(bk_done(e, jf_log, self._addr1), False)
            return
        self._yx_course = e.xpath("//table[@id='tableYXCourse']/tbody/tr/td/input/@tp")
        if len(self._yx_course) >= len(self._course_list):
            self._emit_msg(bk_done(e, jf_log, self._addr1), False)
            return

        if xxbm[0] == 'null':
            xxbm = [None]

        self._emit_msg("获取页面内隐藏的其他参数")
        enter = {
            "xx_bm": xxbm[0],
            "sfzh":  self._name,
            "zybm":  e.xpath(f"//div[@class='divbox']/input[@kslb='{self._ctype}']/@zybm")[0],
            "kslb":  e.xpath(f"//div[@class='divbox']/input[@kslb='{self._ctype}']/@kslb")[0],
            "zkzh":  zkzh[0]
        }
        if len(self._yx_course) > 0:
            enter['mainIds'] = e.xpath("//table[@id='tablePlace']//input/@value")[0]
            enter['qxname'] = e.xpath("//table[@id='tablePlace']//input/@mc")[0]
        return enter

    def _is_rest(self, addr):
        try:
            is_temp = []
            for i in list(self._times.keys()):
                REST = addr.get(f"REST_{i}")
                if int(REST) <= 0:
                    is_temp.append(False)
                else:
                    is_temp.append(True)
            return any(is_temp)
        except:
            return False

    def _get_query_address(self, loop=True):
        while True:
            if self.stop:
                self._emit_msg("停止获取考试地址......")
                break
            if len(self._yx_course) >= len(self._course_list):
                break
            if 0 < len(self._yx_course) < len(self._course_list):
                self._submit(True)
                continue
            params = {"bkzt": "", "qxbm": "", "sfzh": self._name, "dsz": city(self._addr1), "stuType": self._ctype, "stuScope": None, "times": ','.join(list(self._times.keys()))}
            if loop and not self._area:
                self._emit_msg("获取报考城市和区域")
            address_list = self._post("https://zk.sceea.cn/RegExam/switchPage?resourceId=searchPlace", params)
            if address_list.text == '0':
                self._login_cookies("验证码已过期，正在重新登录......")
                continue
            if not address_list:
                break
            address_dict = {}
            try:
                for item in address_list.json().get('data'):
                    address_dict[item.get("QX_MC")] = item
                    if self._area and item.get("QX_MC") == self._addr2:
                        break
            except:
                continue
            self._get_addr_count += 1
            if loop:
                if self._get_addr_count == 1:
                    self._emit_msg(f"第{self._get_addr_count}次获取指定(首选)考区")
                else:
                    self._emit_msg(f"第{self._get_addr_count}次获取指定(首选)考区，{self._addr2}暂无座位，继续查询！！！")
                addr2 = address_dict.get(self._addr2)
                if not self._is_rest(addr2):
                    loop = self._area
                    continue
                else:
                    self._enter['qxname'] = addr2.get("QX_MC")
                    self._enter['mainIds'] = addr2.get("QX_BM")
                    self._submit()
            else:
                self._emit_msg(f"第{self._get_addr_count}次获取{self._addr1}剩下的考区")
                loop_flag = False
                for k, rest in address_dict.items():
                    if not self._is_rest(rest):
                        loop_flag = True
                        continue
                    else:
                        loop_flag = False
                        self._enter['qxname'] = rest.get("QX_MC")
                        self._enter['mainIds'] = rest.get("QX_BM")
                        self._submit()
                        break
                if loop_flag:
                    loop = self._area
                else:
                    break

    def _submit(self, isbk=False):
        if self.stop:
            self._emit_msg("停止报考......")
            return
        if not isbk:
            self._emit_msg("提交报考科目，请稍等......")
        self._enter.pop("zybm", None)
        data = self._post("https://zk.sceea.cn/RegExam/switchPage?resourceId=reg", self._enter)
        if data.text == '0':
            self._login_cookies("验证码已过期，正在重新登录......")
            return
        e = etree.HTML(self._get(f"https://zk.sceea.cn/RegExam/switchPage?resourceId=view&zkz={self._enter.get('zkzh')}").text)
        location = e.xpath("//table[@id='tablePlace']//td[3]/text()")
        if len(location) > 0:
            location = location[0]
        _yx_course_list = e.xpath("//table[@id='tableYXCourse']/tbody/tr/td[3]/text()")[::2]
        courses = '、'.join(_yx_course_list)
        if data.text.find("容量不足") == -1:
            self._emit_msg(f"报考科目：{courses}，报考位置({self._addr1}/{location})", False)
        else:
            if isbk:
                self._get_bk_count += 1
                self._emit_msg(f"循环第{self._get_bk_count}次,还差{len(self._course_list) - len(_yx_course_list)}门课程没座位，报考科目：{courses}，报考位置({self._addr1}/{location})")
            # self._get_query_address(self._area)

    def _distinguish(self, im):
        """
        验证码识别
        :param im:
        :return:
        """
        denois = cv2.fastNlMeansDenoisingColored(im, None, 10, 10, 7, 21)
        gray = cv2.cvtColor(denois, cv2.COLOR_BGR2GRAY)
        blur = cv2.GaussianBlur(gray, (5, 5), 0)
        thresh = cv2.adaptiveThreshold(blur, 255, 1, 1, 11, 2)
        image, contours, hierarchy = cv2.findContours(thresh, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
        code = []
        for cnt in contours:
            if cv2.contourArea(cnt) > 50:
                [x, y, w, h] = cv2.boundingRect(cnt)
                if h > 20:
                    cv2.rectangle(im, (x, y), (x + w, y + h), (0, 255, 0), 2)
                    roi = thresh[y:y + h, x:x + w]
                    roismall = cv2.resize(roi, (10, 10))
                    roismall = roismall.reshape((1, 100))
                    roismall = np.float32(roismall)
                    retval, results, neigh_resp, dists = model.findNearest(roismall, k=1)
                    code.append(str(int((results[0][0]))))
        code.reverse()
        return ''.join(code)

    def _get_code(self):
        """
        获取验证码
        :return:
        """
        file = self._get(url="https://zk.sceea.cn/RegExam/login/AuthImageServlet")
        if not file:
            return None, None
        code = self._distinguish(cv2.imdecode(np.frombuffer(file.content, np.uint8), 1))
        if len(code) != 4:
            return self._get_code()
        self._cookies = file.cookies
        return code

    def _login_cookies(self, log='正在登录系统......'):
        """
        获取 Cookie
        :param log:
        :return:
        """
        while True:
            if self.stop:
                self._emit_msg("停止登录系统......")
                break
            code = self._get_code()
            if code is None:
                break
            if self._login_count > 0:
                log = "登录失败，正在重试......"
            self._emit_msg(log)
            self._login_count += 1
            response = self._post("https://zk.sceea.cn/RegExam/elogin?resourceId=login", {"name": self._name, "code": code, "pwd": self._pwd})
            if not response:
                return
            if response.text == "1":
                self._emit_msg("登录成功......")
                self._cookies = response.cookies
                break
            elif response.text == "2":
                self._emit_msg(f"验证码错误，第{self._login_count}次尝试重新登录！")
                continue
            elif response.text == "8":
                self._emit_msg("服务请求过多，正在抢位！")
                continue
            elif response.text == "7":
                self._emit_msg("当前在线人数过多，正在抢位！")
                continue
            elif response.text == "5":
                self._emit_msg("您没有提交证件照片，不能报考，请及时到注册区县市州，助学点或高校办理！")
                continue
            elif response.text == "9":
                self._emit_msg("您的准考证类型不在可报考时间，请根据报考通知时间段进行登录！")
                continue
            elif response.text == "88":
                self._emit_msg("证件号密码错误或当前不是该类型考生报考时段！")
                continue
            else:
                self._emit_msg("证件号或密码错误，请检查后输入！")

    def _get_course(self):
        """
        获取报考课程列表
        :return:
        """
        if self.stop:
            self._emit_msg("停止获取报考课程列表......")
            return
        if not self._cookies:
            return
        if len(self._yx_course) >= len(self._course_list):
            return
        self._emit_msg("获取课程科目......")
        response = self._post("https://zk.sceea.cn/RegExam/switchPage?resourceId=SearchCourseByZY", {"zybm": self._enter.get("zybm")})
        data = response.json().get("data")
        if not data:
            self._emit_msg("没有获取到课程科目......")
            return
        data_enters = []
        self._emit_msg("获取要报考科目......")
        course_list = {item.get("KC_MC"): item for item in data}
        for item in self._course_list:
            course = course_list.get(item)
            if not course: continue
            SJ_BM = course.pop("SJ_BM")
            if SJ_BM in self._yx_course: continue
            course.pop('KC_BKFY', None)
            course.pop('KSSJ', None)
            course.pop('ZY_MC', None)
            course["zy_bm"] = course.pop("ZY_BM")
            course["kc_bm"] = course.pop("KC_BM")
            self._times[SJ_BM] = course.pop("KC_MC")
            data_enters.append(course)
        self._emit_msg("序列化报考科目数据......")
        self._enter['courseJson'] = json.dumps(data_enters, ensure_ascii=False)

    def _emit_msg(self, log, status=True):
        self.msg[str, str, bool].emit(self._name, log, status)
