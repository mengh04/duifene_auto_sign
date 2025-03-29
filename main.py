import configparser
import os.path
import random
import re
import time
from datetime import datetime

import pyautogui
import pyperclip
import requests
import urllib3
from bs4 import BeautifulSoup
from pywinauto import findwindows
from pywinauto.application import Application
from wxauto import WeChat

import tkinter as tk
from tkinter import messagebox, ttk


class Course:
    id = '0'
    class_id = '0'
    flag = True
    check_list = []
    class_list = []

def get_pid(process_name) -> int | None:
    handles = findwindows.find_elements()
    for handle in handles:
        if handle.class_name == process_name:
            return handle.process_id
    return None

def copy_pdf_link():
    msg = 'https://open.weixin.qq.com/connect/oauth2/authorize?appid=wx1b5650884f657981&redirect_uri=https://www.duifene.com/_FileManage/PdfView.aspx?file=https%3A%2F%2Ffs.duifene.com%2Fres%2Fr2%2Fu6106199%2F%E5%AF%B9%E5%88%86%E6%98%93%E7%99%BB%E5%BD%95_876c9d439ca68ead389c.pdf&response_type=code&scope=snsapi_userinfo&connect_redirect=1#wechat_redirect'
    pyperclip.copy(msg)

def auto_send_link():
    try:
        wx = WeChat()
        who = config['SETTING']['sending_object']
        copy_pdf_link()
        wx.SendMsg(pyperclip.paste(), who)
        text_box.insert(tk.END, f"\n链接已发送至：{who}\n")
    except Exception as e:
        messagebox.showerror("发送失败", f"请检查微信是否登录且'{who}'存在\n错误信息：{str(e)}")

def autoLogin():
    wx = WeChat()
    process1 = 'WeChatMainWndForPC'
    process2 = 'Chrome_WidgetWin_0'
    copy_pdf_link()

    id1 = get_pid(process1)
    app = Application(backend='uia').connect(process=id1)
    win_main_Dialog = app.window(class_name='WeChatMainWndForPC')
    # win_main_Dialog.draw_outline(colour='red')
    # win_main_Dialog.print_control_identifiers(depth=None, filename=None)
    text_control = win_main_Dialog.child_window(title="搜索", control_type="Edit")
    text_control.wait('visible')
    text_control.click_input()
    pyautogui.keyDown('ctrl')
    pyautogui.keyDown('v')
    pyautogui.keyUp('v')
    pyautogui.keyUp('ctrl')
    time.sleep(1)
    text_control = win_main_Dialog.child_window(title="文章、公众号、视频号等", control_type="Text")
    text_control.wait('visible')
    text_control.click_input()

    id2 = get_pid(process2)
    app = Application(backend='uia').connect(process=id2)
    win_main_Dialog = app.window(class_name=process2)
    access_web_text = win_main_Dialog.child_window(title="访问网页", control_type="Text")
    access_web_text.wait('visible')
    access_web_text.click_input()

    tmp_text = win_main_Dialog.child_window(title="PDF文件预览", control_type="TabItem")
    tmp_text.wait('visible')
    # win_main_Dialog.print_control_identifiers(depth=None, filename=None)
    menu_item = win_main_Dialog.child_window(title="更多", control_type="MenuItem")
    menu_item.wait('visible')
    menu_item.click_input()
    menu_item = win_main_Dialog.child_window(title="复制链接", control_type="MenuItem")
    menu_item.wait('visible')
    menu_item.click_input()

    for _ in range(2):
        pyautogui.keyDown('alt')
        pyautogui.keyDown('f4')
        pyautogui.keyUp('f4')
        pyautogui.keyUp('alt')

    link_entry.insert(0, pyperclip.paste())
    login_link()


def on_combo_change(event):
    className = combo_var.get()
    for i in Course.class_list:
        if i["CourseName"] == className:
            Course.id = i["CourseID"]
            Course.class_id = i["TClassID"]
            print(f"课程已选择：{Course.id} ({Course.class_id})")


def login_link():
    try:
        link = link_entry.get()
        if not link.startswith("http"):
            raise ValueError("链接格式不正确")

        code = re.search(r"(?<=code=)\S{32}", link)
        if code is None:
            raise ValueError("链接中未找到授权码")

        x.cookies.clear()
        _r = x.get(url=host + f"/P.aspx?authtype=1&code={code.group(0)}&state=1")
        if _r.status_code == 200:
            get_class_list()
        else:
            raise ConnectionError(f"登录请求失败，状态码：{_r.status_code}")

    except Exception as e:
        messagebox.showerror("登录失败", str(e))
        text_box.insert(tk.END, f"\n[登录错误] {str(e)}")


def get_user_id():
    _r = x.get(url=host + "/_UserCenter/MB/index.aspx")
    if _r.status_code == 200:
        soup = BeautifulSoup(_r.text, "html.parser")
        stu_id = soup.find(id="hidUID").get("value")
        return stu_id


def sign(sign_code):
    # 签到码
    if len(sign_code) == 4:
        headers = {
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "Referer": "https://www.duifene.com/_CheckIn/MB/CheckInStudent.aspx?moduleid=16&pasd="
        }
        params = f"action=studentcheckin&studentid={get_user_id()}&checkincode={sign_code}"
        _r = x.post(
            url=host + "/_CheckIn/CheckIn.ashx", data=params, headers=headers)
        if _r.status_code == 200:
            msg = _r.json()["msgbox"]
            text_box.insert(tk.END, f"\t{msg}\n\n")
            if msg == "签到成功！":
                return True
    # 二维码
    else:
        _r = x.get(url=host + "/_CheckIn/MB/QrCodeCheckOK.aspx?state=" + sign_code)
        if _r.status_code == 200:
            soup = BeautifulSoup(_r.text, "html.parser")
            msg = soup.find(id="DivOK").get_text()
            if "签到成功" in msg:
                text_box.insert(tk.END, f"\t{msg}\n\n")
            else:
                text_box.insert(tk.END, f"\t非微信链接登录，二维码无法签到\n\n")
            return True


def sign_location(longitude, latitude):
    longitude = str(round(float(longitude) + random.uniform(-0.000089, 0.000089), 8))
    latitude = str(round(float(latitude) + random.uniform(-0.000089, 0.000089), 8))

    headers = {
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "Referer": "https://www.duifene.com/_CheckIn/MB/CheckInStudent.aspx?moduleid=16&pasd="
    }
    params = f"action=signin&sid={get_user_id()}&longitude={longitude}&latitude={latitude}"
    _r = x.post(
        url=host + "/_CheckIn/CheckInRoomHandler.ashx", data=params, headers=headers)
    if _r.status_code == 200:
        msg = _r.json()["msgbox"]
        text_box.insert(tk.END, f"\t{msg}\n\n")
        if msg == "签到成功！":
            return True

def get_arrival_count(ciid):

    ajax_url = "https://www.duifene.com/_CheckIn/MBCount.ashx"  # 替换成实际URL
    params = {
        "action": "getcheckintotalbyciid",
        "ciid": ciid,
        "t": "cking"
    }

    response = x.get(ajax_url, params=params)
    data = response.json()
    arrival_count = data["TotalNumber"] - data["AbsenceNumber"]
    total_count = data["TotalNumber"]
    # print(int(signed_percent.get()) / 100)
    if (arrival_count / total_count) >= (int(signed_percent.get()) / 100):
        return True, arrival_count
    else:
        return False, arrival_count


def watching_sign():
    is_login()

    line_count = int(text_box.index('end-1c').split('.')[0])
    text_box.delete(f"{line_count}.0", f"{line_count}.end")
    current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    text_box.insert(tk.END, f"持续监控：{current_time}")  # 插入当前时间
    text_box.see(tk.END)  # 滚动到最后一行

    _r = x.get(url=host + f"/_CheckIn/MB/TeachCheckIn.aspx?classid={Course.class_id}&temps=0&checktype=1&isrefresh=0"
                          f"&timeinterval=0&roomid=0&match=")
    if _r.status_code == 200:
        if "HFChecktype" in _r.text:
            status = False
            soup = BeautifulSoup(_r.text, "html.parser")
            HFSeconds = soup.find(id="HFSeconds").get("value")
            HFChecktype = soup.find(id="HFChecktype").get("value")
            HFCheckInID = soup.find(id="HFCheckInID").get("value")
            HFClassID = soup.find(id="HFClassID").get("value")
            signable, signed_count = get_arrival_count(HFCheckInID)
            # print(signable, signed_count)
            if Course.class_id in HFClassID:
                if HFCheckInID not in Course.check_list:
                    # 数字签到
                    if HFChecktype == '1':
                        sign_code = soup.find(id="HFCheckCodeKey").get("value")
                        # if sign_code is not None and int(HFSeconds) <= int(seconds_entry.get()):
                        if signable:
                            text_box.insert(tk.END, f"\n\n{current_time} 签到ID：{HFCheckInID} 开始签到\t签到码：{sign_code}")
                            status = sign(sign_code)
                        else:
                            text_box.insert(tk.END, f"\t签到码签到\t倒计时：{HFSeconds}秒\t已签到人数：{signed_count}人\t签到码：{sign_code}")
                    # 二维码签到
                    elif HFChecktype == '2':
                        # if HFCheckInID is not None and int(HFSeconds) <= int(seconds_entry.get()):
                        if signable:
                            text_box.insert(tk.END, f"\n\n{current_time} 签到ID：{HFCheckInID} 开始签到\t二维码签到")
                            status = sign(HFCheckInID)
                        else:
                            text_box.insert(tk.END, f"\t二维码签到\t倒计时：{HFSeconds}秒\t已签到人数：{signed_count}人")
                    # 定位签到
                    elif HFChecktype == '3':
                        HFRoomLongitude = soup.find(id="HFRoomLongitude").get("value")
                        HFRoomLatitude = soup.find(id="HFRoomLatitude").get("value")
                        # if HFRoomLongitude is not None and HFRoomLatitude is not None and int(HFSeconds) <= int(seconds_entry.get()):
                        if HFRoomLongitude and HFRoomLatitude and signable:
                            text_box.insert(tk.END, f"\n\n{current_time} 签到ID：{HFCheckInID} 开始签到\t定位签到")
                            status = sign_location(HFRoomLongitude, HFRoomLatitude)
                        else:
                            text_box.insert(tk.END, f"\t定位签到\t倒计时：{HFSeconds}秒\t已签到人数：{signed_count}人")
                    if status:
                        Course.check_list.append(HFCheckInID)
            else:
                text_box.insert(tk.END, f"\t 检测到非本班签到")
    if Course.flag:
        root.after(1000, watching_sign)


def go_sign():
    if combo.get() is None or combo.get() == '':
        messagebox.showerror("错误提示", "请先登录")
        return


    print(f"开始监听课程ID：{Course.id} ({Course.class_id})")

    headers = {
        "Referer": "https://www.duifene.com/_UserCenter/MB/index.aspx"
    }
    _r = x.get(url=host + "/_UserCenter/MB/Module.aspx?data=" + Course.id, headers=headers)
    if _r.status_code == 200:
        if Course.id in _r.text:
            text_box.delete("1.0", "end")
            soup = BeautifulSoup(_r.text, "html.parser")
            CourseName = soup.find(id="CourseName").text
            text_box.insert(tk.END, f"正在监听【{CourseName}】的签到活动\n\n")
            watching_sign()


def get_class_list():
    headers = {
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "Referer": "https://www.duifene.com/_UserCenter/PC/CenterStudent.aspx"
    }
    params = "action=getstudentcourse&classtypeid=2"
    _r = x.post(url=host + "/_UserCenter/CourseInfo.ashx", data=params, headers=headers)
    if _r.status_code == 200:
        _json = _r.json()
        if _json is not None:
            try:
                msg = _json["msgbox"]
                messagebox.showerror("", f"{msg} 请重新登录。")
                x.cookies.clear()
            except Exception as e:
                class_name_list = []
                for i in _json:
                    class_name_list.append(i["CourseName"])
                combo['values'] = tuple(class_name_list)
                combo.set(class_name_list[0])
                Course.id = _json[0]['CourseID']
                Course.class_id = _json[0]["TClassID"]
                Course.class_list = _json


def is_login():
    headers = {
        "Referer": "https://www.duifene.com/_UserCenter/PC/CenterStudent.aspx",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"
    }
    _r = x.get(host + "/AppCode/LoginInfo.ashx", data="Action=checklogin", headers=headers)
    if _r.status_code == 200:
        if _r.json()["msg"] == "1":
            return True
        else:
            messagebox.showwarning("登录状态失效", "请重新登录账号")
            x.cookies.clear()
            Course.flag = False
            return False

def save_setting():
    config['SETTING'] = {
        "signed_percent": signed_percent.get(),
        "sending_object": sending_entry.get()
    }
    try:
        with open(filename, 'w', encoding='utf-8') as configfile:  # 添加编码
            config.write(configfile)
        messagebox.showinfo("提示", "设置保存成功")
    except Exception as e:
        messagebox.showerror("保存失败", f"错误信息：{str(e)}")

def read_setting(filename):
    if not os.path.exists(filename):
        config['SETTING'] = {'signed_percent': '50', 'sending_object': '文件传输助手'}
        try:
            with open(filename, 'w', encoding='utf-8') as configfile:  # 添加编码
                config.write(configfile)
        except Exception as e:
            messagebox.showerror("创建配置失败", f"错误信息：{str(e)}")
    try:
        config.read(filename, encoding='utf-8')  # 添加编码
    except Exception as e:
        messagebox.showerror("读取配置失败", f"错误信息：{str(e)}")


if __name__ == '__main__':
    # 初始化配置
    host = "https://www.duifene.com"
    UA = 'Mozilla/5.0 (iPhone; CPU iPhone OS 16_6 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Mobile/15E148 MicroMessenger/8.0.40(0x1800282a) NetType/WIFI Language/zh_CN '
    urllib3.disable_warnings()
    x = requests.Session()
    x.headers['User-Agent'] = UA
    x.verify = False
    config = configparser.ConfigParser()
    filename = 'duifenyi.ini'
    read_setting(filename)

    # 创建主窗口
    root = tk.Tk()
    root.title("对分易自动签到 v2.2")
    root.geometry("1000x680")  # 调整窗口尺寸
    root.resizable(True, True)

    # 配置网格布局
    root.grid_rowconfigure(1, weight=1)
    root.grid_columnconfigure(0, weight=3)
    root.grid_columnconfigure(1, weight=1)

    # 样式配置
    style = ttk.Style()
    style.theme_use('clam')
    font_style = ('微软雅黑', 10)
    style.configure('TButton', font=font_style, padding=5)
    style.configure('Accent.TButton', font=('微软雅黑', 12), foreground='white', background='#2196F3')
    style.map('Accent.TButton', background=[('active', '#1976D2')])

    # 选项卡控件
    tab_control = ttk.Notebook(root)
    tab1 = ttk.Frame(tab_control)
    tab2 = ttk.Frame(tab_control)
    tab_control.add(tab1, text="微信登录")
    tab_control.add(tab2, text="设置")
    tab_control.grid(row=0, column=0, padx=10, pady=10, sticky='nsew')

    # 登录页内容 -------------------------------------------------
    login_frame = ttk.Frame(tab1)
    ttk.Label(login_frame, text="登录链接", font=('微软雅黑', 12)).pack(pady=5)

    link_entry = ttk.Entry(login_frame, width=40)
    link_entry.pack(pady=5, ipady=3)

    btn_frame = ttk.Frame(login_frame)
    # 修改按钮绑定：添加command参数
    ttk.Button(btn_frame, text="自动登录", command=autoLogin, width=10, padding=2).grid(row=0, column=0, padx=2)
    ttk.Button(btn_frame, text="手动登录", command=login_link, width=10, padding=2).grid(row=0, column=1, padx=2)
    ttk.Button(btn_frame, text="发送链接", command=auto_send_link, width=10, padding=2).grid(row=0, column=2, padx=2)
    btn_frame.pack(pady=(3, 5))

    login_frame.pack(pady=15, fill=tk.X)

    # 设置页内容 -------------------------------------------------
    setting_frame = ttk.Frame(tab2)

    # 帮助说明
    help_text = """设置说明：
    1.签到百分比：当签到人数达到总人数设定比例时自动签到（推荐50-70）
    2.发送对象：微信中的准确联系人名称（建议先发送到'文件传输助手'测试）"""
    ttk.Label(setting_frame, text=help_text, font=font_style, justify=tk.LEFT).pack(pady=10, anchor=tk.W)

    # 设置项
    form_frame = ttk.Frame(setting_frame)
    ttk.Label(form_frame, text="签到百分比:", width=10).grid(row=0, column=0, sticky='e', padx=5)
    signed_percent = ttk.Entry(form_frame, width=8)
    signed_percent.insert(0, config['SETTING']['signed_percent'])
    signed_percent.grid(row=0, column=1, sticky='w', pady=5)

    ttk.Label(form_frame, text="发送对象:", width=10).grid(row=1, column=0, sticky='e', padx=5)
    sending_entry = ttk.Entry(form_frame, width=20)
    sending_entry.insert(0, config['SETTING']['sending_object'])
    sending_entry.grid(row=1, column=1, sticky='w', pady=5)

    ttk.Button(form_frame, text="保存设置", command=save_setting, width=18).grid(row=2, columnspan=2, pady=15)
    form_frame.pack(pady=10)
    setting_frame.pack(fill=tk.BOTH, expand=True)

    # 右侧控制面板 -----------------------------------------------
    control_frame = ttk.Frame(root)
    control_frame.grid(row=0, column=1, padx=10, pady=10, sticky='nsew')

    # 课程选择
    ttk.Label(control_frame, text="当前课程", font=('微软雅黑', 12)).pack(pady=8)
    combo_var = tk.StringVar()
    combo = ttk.Combobox(control_frame, textvariable=combo_var, state="readonly", width=18)
    combo.pack(pady=5, ipady=2)
    combo.bind("<<ComboboxSelected>>", on_combo_change)  # 这里是关键，绑定了选择课程的事件

    # 监听按钮
    ttk.Button(control_frame, text="开始监听", command=go_sign, style='Accent.TButton', width=16).pack(pady=15)

    # 日志输出 ---------------------------------------------------
    log_frame = ttk.Frame(root)
    text_box = tk.Text(log_frame, wrap=tk.WORD, font=('微软雅黑', 9),undo=True, maxundo=100)
    scroll = ttk.Scrollbar(log_frame, command=text_box.yview)
    text_box.configure(yscrollcommand=scroll.set)

    text_box.grid(row=0, column=0, sticky='nsew')
    scroll.grid(row=0, column=1, sticky='ns')
    log_frame.grid_rowconfigure(0, weight=1)
    log_frame.grid_columnconfigure(0, weight=1)

    log_frame.grid(row=1, column=0, columnspan=2, padx=10, pady=(0, 10), sticky='nsew')

    root.mainloop()