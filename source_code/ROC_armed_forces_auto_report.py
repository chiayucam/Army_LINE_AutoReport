import time
import tkinter as tk
import threading
import win32gui
import re
from tkinter import ttk
from tkinter.messagebox import showinfo
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException, NoAlertPresentException, TimeoutException
from threading import Timer


class App:
    def __init__(self, window):
        self.window = window
        self.window.title('國軍LINE回報排程 v1.1')
        self.window.protocol("WM_DELETE_WINDOW", self.on_exit)
        self.window.geometry('780x440')
        self.window.configure(background='white')
        self.window.resizable(False, False)

        #################
        ## login frame ##
        #################
        self.login_frame = tk.Frame(self.window, height=150, width=300, padx=10, borderwidth=1, highlightbackground="black", highlightcolor="black", highlightthickness=1)
        self.login_frame.grid(column=0, row=0, padx=10, sticky="NW")
        self.login_frame.pack_propagate(0)

        self.login_label = tk.Label(self.login_frame, text="LINE電腦版登入", font="bold, 11")
        self.login_label.pack()

        self.account_frame = tk.Frame(self.login_frame)
        self.account_frame.pack(side="top", pady=5)
        self.account_label = tk.Label(self.account_frame, text="帳號:")
        self.account_label.pack(side="left")
        self.account_entry = tk.Entry(self.account_frame)
        self.account_entry.pack(side="left")

        self.password_frame = tk.Frame(self.login_frame)
        self.password_frame.pack(side="top")
        self.password_label = tk.Label(self.password_frame, text="密碼:")
        self.password_label.pack(side="left")
        self.password_entry = tk.Entry(self.password_frame, show="\u2022")
        self.password_entry.pack(side="left")

        self.login_btn = tk.Button(self.login_frame, text="登入", command=self.login_btn_press)
        self.login_btn.pack(pady=5)

        self.login_status_label = tk.Label(self.login_frame, text="尚未登入")
        self.login_status_label.pack()

        self.config_code_label = tk.Label(self.login_frame, text="")
        self.config_code_label.pack()

        ###################
        ## message frame ##
        ###################
        self.message_frame = tk.Frame(self.window, height=270, width=300, borderwidth=1, highlightbackground="black", highlightcolor="black", highlightthickness=1)
        self.message_frame.grid(column=0, row=1, padx=10, pady=10, sticky="NW")
        self.message_frame.pack_propagate(0)

        self.message_label = tk.Label(self.message_frame, text="傳送訊息時間與內容", font="bold, 11")
        self.message_label.pack()

        self.chatroom_frame = tk.Frame(self.message_frame)
        self.chatroom_frame.pack(side="top")
        self.chatroom_label = tk.Label(self.chatroom_frame, text="群組名稱:")
        self.chatroom_label.pack(side="left")
        self.chatroom_entry = tk.Entry(self.chatroom_frame, width=30)
        self.chatroom_entry.pack(side="left")
        self.chatroom_entry.config(state="disabled")

        self.datetime_frame = tk.Frame(self.message_frame)
        self.datetime_frame.pack(side="top", pady=5)
        self.datetime_label = tk.Label(self.datetime_frame, text="傳送時間:")
        self.datetime_label.pack(side="left")
        self.datetime_entry = tk.Entry(self.datetime_frame, width=30)
        self.datetime_entry.insert(0, datetime.now().strftime('%m/%d, %H:%M:%S'))
        self.datetime_entry.pack(side="left")
        self.datetime_entry.config(state="disabled")
        
        self.headertype_frame = tk.Frame(self.message_frame)
        self.headertype_frame.pack(side="top", pady=5)
        self.headertype_label = tk.Label(self.headertype_frame, text="標頭格式:")
        self.headertype_label.pack(side="left")
        self.headertype_combobox = ttk.Combobox(self.headertype_frame, state="disabled", values=["無標頭", "M/D hhmm 回報"], width=27)
        self.headertype_combobox.current(0)
        self.headertype_combobox.pack(side="left")
        
        self.content_frame = tk.Frame(self.message_frame)
        self.content_frame.pack(side="top", pady=5)
        self.content_label = tk.Label(self.content_frame, text="傳送內容:")
        self.content_label.pack(side="left", anchor="nw")
        self.content_text = tk.Text(self.content_frame, height=7, width=30, font="TkDefaultFont")
        self.content_text.pack(side="left")
        self.content_text.config(state="disabled")

        self.insert_btn = tk.Button(self.message_frame, text="插入排程", command=self.insert_btn_press)
        self.insert_btn.pack(pady=5)
        self.insert_btn.config(state="disabled")

        ####################
        ## schedule frame ##
        ####################
        self.schedule_frame = tk.Frame(self.window, height=430, width=450, borderwidth=1, highlightbackground="black", highlightcolor="black", highlightthickness=1)
        self.schedule_frame.grid(column=1, row=0, rowspan=2, sticky="NW", padx=1)
        self.schedule_frame.pack_propagate(0)

        self.schedule_label = tk.Label(self.schedule_frame, text="預定傳送排程", font="bold, 11")
        self.schedule_label.pack()

        self.schedule_tree = ttk.Treeview(self.schedule_frame, columns=['1','2','3'], show='headings', height=16, selectmode="browse")
        self.schedule_tree.column('1', width=110, anchor="w")
        self.schedule_tree.column('2', width=200, anchor="w")
        self.schedule_tree.column('3', width=110, anchor="w")
        self.schedule_tree.heading('1', text="傳送日期與時間")
        self.schedule_tree.heading('2', text="傳送內容")
        self.schedule_tree.heading('3', text="群組")
        self.schedule_tree.pack(pady=5)

        self.delete_btn = tk.Button(self.schedule_frame, text="刪除排程", command=self.delete_btn_press)
        self.delete_btn.pack(pady=10)
        self.delete_btn.config(state="disabled")

    def login_btn_press(self):
        login(ID=self.account_entry.get(), password=self.password_entry.get())
        try:
            driver.implicitly_wait(5)
            driver.find_element_by_class_name("mdCMN01Code")
            self.login_status_label.config(text="登入成功", fg="green")
            self.login_frame.config(highlightbackground="green", highlightcolor="green")
            self.config_code_label.config(text="請於手機版LINE輸入驗證碼:"+get_config_code())
            self.login_btn.config(state="disabled")
            self.account_entry.config(state="disabled")
            self.password_entry.config(state="disabled")
            self.chatroom_entry.config(state="normal")
            self.datetime_entry.config(state="normal")
            self.headertype_combobox.config(state="readonly")
            self.content_text.config(state="normal")
            self.insert_btn.config(state="normal")
            self.delete_btn.config(state="normal")
        except (NoSuchElementException, TimeoutException) as error:
            incorrect_text = driver.find_element_by_id("login_incorrect").get_attribute("innerText")
            self.login_status_label.config(text=incorrect_text, fg="red")
            self.login_frame.config(highlightbackground="red", highlightcolor="red")

    def insert_btn_press(self):
        if check_time_format_valid(self.datetime_entry.get()):
            insert_position = get_insert_position(self.schedule_tree, self.datetime_entry.get())
            self.schedule_tree.insert("", insert_position, values=(self.datetime_entry.get(), self.content_text.get("1.0", "end"), self.chatroom_entry.get()))
            datetime_str, report_message, chatroom_name = self.schedule_tree.item(self.schedule_tree.get_children()[0])["values"]
            report_message = report_message.rstrip()
            report_datetime = datetime.strptime(str(datetime.now().year)+"/"+datetime_str, '%Y/%m/%d, %H:%M:%S')
            headertype = self.headertype_combobox.get()
            update_timer(report_datetime, report_message, chatroom_name, headertype)
        else:
            time_format_incorrect_popup()

    def delete_btn_press(self):
        self.schedule_tree.delete(self.schedule_tree.selection())
        try:
            datetime_str, report_message, chatroom_name = self.schedule_tree.item(self.schedule_tree.get_children()[0])["values"]
            report_message = report_message.rstrip()
            report_datetime = datetime.strptime(str(datetime.now().year)+"/"+datetime_str, '%Y/%m/%d, %H:%M:%S')
            headertype = self.headertype_combobox.get()
            update_timer(report_datetime, report_message, chatroom_name, headertype)
        except IndexError:
            kill_timer()

    def on_exit(self):
        self.window.destroy()
        driver.quit()

        

def init_webdriver():
    extension_path = "chrome-extension://ophjlpahpchlmihnnnihgmmeilfjmjjc/index.html"
    chrome_options = Options()
    chrome_options.add_extension('extension_2_4_1_0.crx')
    driver = webdriver.Chrome('./chromedriver.exe', options=chrome_options)
    # driver.set_window_position(-10000, 0)
    driver.get(extension_path)
    return driver

def login(ID, password):
    line_login_email = driver.find_element_by_id("line_login_email")
    line_login_pwd = driver.find_element_by_id("line_login_pwd")
    line_login_email.clear()
    line_login_email.send_keys(ID)
    line_login_pwd.clear()
    line_login_pwd.send_keys(password)
    driver.find_element_by_id("login_btn").click()

def refresh():
    driver.get(driver.current_url)
    try:
        driver.switch_to.alert.accept()
    except NoAlertPresentException:
        pass
    WebDriverWait(driver, 5, poll_frequency=0.1, ignored_exceptions=(TimeoutException)).until(EC.presence_of_element_located((By.ID, "line_login_pwd")))

def get_config_code():
    return driver.find_element_by_class_name("mdCMN01Code").get_attribute("innerText")

def select_chatroom(chatroom_name):
    WebDriverWait(driver, 5, poll_frequency=0.1, ignored_exceptions=(TimeoutException)).until(EC.presence_of_element_located((By.CLASS_NAME, "mdLFT07Groups")))  
    try:
        driver.find_element_by_xpath("//button[@title='群組']").click()
        chat_room_search = driver.find_element_by_xpath("//input[@placeholder='搜尋群組名稱']")
        chat_room_search.clear()
        chat_room_search.send_keys(chatroom_name)
        driver.find_element_by_xpath(f"//ul[@id='joined_group_list_body']//li[@title='{chatroom_name}']").click()
    except NoSuchElementException:
        chatroom_name_incorrect_popup(chatroom_name)

def chatroom_name_incorrect_popup(chatroom_name):
        showinfo("chatroom error", "找不到這個群組: "+chatroom_name)

def send_message(message_list):
    chat_room_input = driver.find_element_by_id("_chat_room_input")
    chat_room_input.click()
    for message in message_list:
        for part in list(filter(None, message.split("\n"))):
            chat_room_input.send_keys(part)
            chat_room_input.send_keys(Keys.SHIFT+Keys.ENTER)
    chat_room_input.send_keys(Keys.ENTER)

def reformat_message(message):
    message_list = re.split("(?=\s\d{3}\s)", message)
    return message_list

def insert_message(message_list, report_message):
    report_message = "\n"+report_message
    try:
        insert_index = next(idx for idx, message in enumerate(message_list[1:]) if message[:4] > report_message[:4])+1
        message_list.insert(insert_index, report_message)
        return message_list
    except StopIteration:
        message_list.insert(len(message_list), report_message)
        return message_list
    
def get_insert_position(schedule_tree, datetime_str):
    tree_items = schedule_tree.get_children()
    tree_size = len(tree_items)
    if tree_size:
        scheduled_datetime_list = [schedule_tree.item(tree_item)["values"][0] for tree_item in tree_items]
        try:
            insert_index = next(idx for idx, scheduled_datetime_str in enumerate(scheduled_datetime_list) if datetime.strptime(scheduled_datetime_str, '%m/%d, %H:%M:%S') > datetime.strptime(datetime_str, '%m/%d, %H:%M:%S'))
            return insert_index
        except StopIteration:
            return "end"
    else:
        return "end"

def get_prev_message(check_num, headertype):
    WebDriverWait(driver, 5, poll_frequency=1, ignored_exceptions=(TimeoutException)).until(EC.presence_of_element_located((By.ID, "_chat_room_input")))
    prev_messages = driver.find_elements_by_class_name("mdRGT07MsgTextInner")[-check_num:]
    report_title, keywords = get_report_title(datetime.now(), headertype)
    for prev_message in prev_messages[::-1]:
        prev_message_text = prev_message.get_attribute("innerText")
        if all(keyword in prev_message_text for keyword in keywords):
            return prev_message_text
    return None

def ceil_datetime(dt, delta):
    q, r = divmod(dt - datetime.min, delta)
    return (datetime.min + (q + 1)*delta) if r else dt

def get_report_title(report_datetime, headertype):
    report_title = "{d.month}/{d.day} {d.hour:02}{d.minute:02}".format(d=ceil_datetime(report_datetime, timedelta(hours=1)))+" 回報"
    keywords = report_title.split(" ")[:2]
    return report_title, keywords

def execute_timer(report_datetime, report_message, chatroom_name, headertype):
    refresh()
    login(ID=app.account_entry.get(), password=app.password_entry.get())
    select_chatroom(chatroom_name)
    prev_message = get_prev_message(5, headertype)
    if prev_message:
        send_message(insert_message(reformat_message(prev_message), report_message))
    else:
        send_message([get_report_title(report_datetime, headertype), report_message])
    app.schedule_tree.delete(app.schedule_tree.get_children()[0])
    try:
        datetime_str, report_message, chatroom_name = app.schedule_tree.item(app.schedule_tree.get_children()[0])["values"]
        report_message = report_message.rstrip()
        report_datetime = datetime.strptime(str(datetime.now().year)+"/"+datetime_str, '%Y/%m/%d, %H:%M:%S')
        update_timer(report_datetime, report_message, chatroom_name, headertype)
    except IndexError:
        kill_timer()

def update_timer(report_datetime, report_message, chatroom_name, headertype):
    for thread in threading.enumerate():
        if thread.name == "ROC_report_timer":
            thread.cancel()
    t = Timer(time_diff_seconds(report_datetime), execute_timer, (report_datetime, report_message, chatroom_name, headertype))
    t.daemon = True
    t.name = "ROC_report_timer"
    t.start()

def kill_timer():
    for thread in threading.enumerate():
        if thread.name == "ROC_report_timer":
            thread.cancel()

def check_time_format_valid(datetime_str):
    try:
        datetime.strptime(datetime_str, '%m/%d, %H:%M:%S')
        return True
    except ValueError:
        return False

def time_diff_seconds(report_datetime):
    try:
        time_diff = report_datetime-datetime.now()
        return int(time_diff.total_seconds())
    except ValueError:
        time_format_incorrect_popup()

def time_format_incorrect_popup():
    showinfo("datetime error", "日期或時間格式錯誤")

def enumWindowFunc(hwnd, windowList):
    text = win32gui.GetWindowText(hwnd)
    className = win32gui.GetClassName(hwnd)
    if 'chromedriver' in text.lower() or 'chromedriver' in className.lower():
        win32gui.ShowWindow(hwnd, False)

if __name__ == "__main__":
    driver = init_webdriver()
    win32gui.EnumWindows(enumWindowFunc, [])
    window = tk.Tk()
    app = App(window)
    window.mainloop()