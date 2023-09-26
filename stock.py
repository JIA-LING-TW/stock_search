# 要下載的指令
# pip install bs4
# pip install yfinance
# pip install twstock
# https://www.lfd.uci.edu/~gohlke/pythonlibs/#ta-lib (38 32)
# pip install TA_Lib-0.4.24-cp38-cp38-win32.whl
# pip install mplfinance

import atexit
import sys

import tkinter as tk
from tkinter import ttk
# 抓網址用
import requests
from bs4 import BeautifulSoup
# 股票資料
import twstock
import yfinance as yf
# 時間
import datetime
# 多執行緒
import threading
import time
# 繪製圖表用
import pandas as pd

import mplfinance as mpf
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


def cleanup_function():
    # .close()
    plt.close()
    win1.destroy()


global now_switch


now_switch = 0
####### 定義###########
# 收藏按鈕


def update_table():
    try:
        error_msg.set("")
        stock = twstock.realtime.get(input_field.get())
        # print(stock) #資料抓取
        # 抓當前價格  如果收盤，已收盤價為主
        if stock['realtime']['latest_trade_price'] == "-":
            nowprice = float(stock['realtime']['best_bid_price'][0])
        else:
            nowprice = float(stock['realtime']['latest_trade_price'])
        # 放入價格
        if input_field.get() not in total_list_stock:
            total_list_stock.append(input_field.get())
            # 寫入檔案
            path = 'collect.txt'
            f = open(path, 'w')
            for i in total_list_stock:
                f.writelines(i+'\n')
            f.close()
            # collet_table.insert(tk.END, "{:　<6} {:　<5} {:>9.2f} {:>20}".format(stock['info']['code'], stock['info']['name'], nowprice,stock['realtime']['accumulate_trade_volume']))
            collet_table.insert('', tk.END, values=(
                stock['info']['code'], stock['info']['name'], nowprice, stock['realtime']['accumulate_trade_volume']))
        else:
            error_msg.set("已收藏此股票")
    except:
        error_msg.set("查無此股票")


def show_now_button():
    btn1.grid(row=1, column=0, pady=10, padx=20)
    btn2.grid(row=1, column=1, padx=20)
    btn3.grid(row=1, column=2, padx=20)


def touch(event):  # 當按下收藏中的股票時
    global show_switch
    global it
    global col2
    show_switch += 1
    if show_switch == 1:
        show_now_button()
        show_switch += 1
    n = event.widget
    it = n.selection()[0]
    col2 = n.item(it, "values")[0]
    if col2 != "代碼":
        data = twstock.realtime.get(col2)
        get_data(data)
# 搜尋


def search():
    global show_switch
    show_switch += 1
    if show_switch == 1:
        show_now_button()
        show_switch += 1
    try:
        error_msg.set("")
        data = twstock.realtime.get(input_field.get())
        get_data(data)
    except:
        error_msg.set("查無此股票")
# 獲得分時的資訊


def get_data(data):
    stockname.set(data['info']['name'])
    stockname1.set(data['info']['code'])
    if data['realtime']['latest_trade_price'] == "-":
        stockname2.set(float(data['realtime']['best_bid_price'][0]))
    else:
        stockname2.set(float(data['realtime']['latest_trade_price']))
    # 取得前一天日期
    today = datetime.date.today()
    day = today.day-1
    yesterday = today.replace(day=day)

    # 抓取前天收盤價
    try:
        df = yf.download(f"{stockname1.get()}.TW", yesterday)
    except:
        df = yf.download(f"{stockname1.get()}.TWO", yesterday)
    yclose = float(df.values[0][4])

    # 抓取當天收盤價或最後價格
    if data['realtime']['latest_trade_price'] != "-":
        nowprice = float(data['realtime']['latest_trade_price'])
    else:
        nowprice = float(data['realtime']['best_bid_price'][0])

    # 計算漲跌
    agio = nowprice-yclose
    price = format(abs(agio)/yclose * 100, '.2f')
    if agio > 0:
        stockname3.set("▲"+format(agio, '.2f'))
        stockname4.set("("+price+"%)")
    elif agio == 0:
        updown = format(agio, '.2f')
        stockname3.set(updown)
        stockname4.set("("+price+"%)")
    else:
        stockname3.set("▼"+format(abs(agio), '.2f'))
        stockname4.set("("+price+"%)")
    now()
    now_switch = 1

# 重新整理


def stock_re():
    stock_re_button()
    # while 1:
    #     time.sleep(30)
    #     stock_re_button()
    #     if now_switch :
    #         now()


def stock_re_button():
    x = collet_table.get_children()
    for i in x:
        collet_table.delete(i)

    # collet_table.delete(0,tk.END)
    for i in total_list_stock:
        stock = twstock.realtime.get(i)
        # print(stock)
        if stock['realtime']['latest_trade_price'] == "-":
            nowprice = float(stock['realtime']['best_bid_price'][0])
        else:
            nowprice = float(stock['realtime']['latest_trade_price'])
        # collet_table.insert(tk.END, "{:　<6} {:　<5} {:>9.2f} {:>20}".format(stock['info']['code'], stock['info']['name'], float(nowprice),int(stock['realtime']['accumulate_trade_volume'])))
            collet_table.insert('', tk.END, values=(
                stock['info']['code'], stock['info']['name'], nowprice, stock['realtime']['accumulate_trade_volume']))
# 刪除按鈕


def delbox():
    global total_list_stock
    # 刪除
    f = open('collect.txt')
    total_list_stock = list(f.read().split())
    total_list_stock.remove(col2)
    f.close()
    # 寫入
    w = open('collect.txt', 'w')
    for i in total_list_stock:
        w.writelines(i+'\n')
    w.close()
    collet_table.delete(it)
# 3個按鈕方法


def now():
    plt.close()
    for widget in frame2.winfo_children():
        widget.destroy()
    columns = ["開盤", "最高價", "最低價", "成交量", "當天收盤", "漲跌"]
    table = ttk.Treeview(frame2, height=2, show="headings", columns=columns)
    table.column("開盤", width=80, anchor="center")
    table.column("最高價", width=80, anchor="center")
    table.column("最低價", width=80, anchor="center")
    table.column("成交量", width=80, anchor="center")
    table.column("當天收盤", width=80, anchor="center")
    table.column("漲跌", width=100, anchor="center")

    table.heading("開盤", text="開盤")
    table.heading("最高價", text="最高價")
    table.heading("最低價", text="最低價")
    table.heading("成交量", text="成交量")
    table.heading("當天收盤", text="當天收盤")
    table.heading("漲跌", text="漲跌")

    table.pack()
    data = twstock.realtime.get(stockname1.get())
    # print(data)
    # 表格內數值變數

    if data['realtime']['accumulate_trade_volume'] == "0":
        nopen = data['realtime']['open']
        nhigh = data['realtime']['high']
        nlow = data['realtime']['low']
        nvolume = data['realtime']['accumulate_trade_volume']
        nclose = data['realtime']['latest_trade_price']
    else:
        nopen = format(float(data['realtime']['open']), '.2f')
        nhigh = format(float(data['realtime']['high']), '.2f')
        nlow = format(float(data['realtime']['low']), '.2f')
        nvolume = data['realtime']['accumulate_trade_volume']
        nclose = format(float(data['realtime']['latest_trade_price']), '.2f')

    # 股票重新整理
    # nowstockre = threading.Thread(target=now_stock_re)
    # nowstockre.start()

    # 取得前一天日期
    today = datetime.date.today()
    day = today.day-1
    yesterday = today.replace(day=day)

    # 抓取前天收盤價
    try:
        df = yf.download(f"{stockname1.get()}.TW", yesterday)
        yclose = float(df.values[0][4])
    except:
        df = yf.download(f"{stockname1.get()}.TWO", yesterday)
        yclose = float(df.values[0][4])

    try:
        # 抓取當天收盤價或最後價格
        if data['realtime']['latest_trade_price'] != "-":
            nowprice = float(data['realtime']['latest_trade_price'])
        else:
            nowprice = float(data['realtime']['best_bid_price'][0])

        # 計算漲跌
        agio = nowprice-yclose
        price = format(abs(agio)/yclose * 100, '.2f')
        if data['realtime']['accumulate_trade_volume'] != "0":
            if agio > 0:
                updown = "▲"+format(agio, '.2f')+"("+price+"%)"
            elif agio == 0:
                updown = format(agio, '.2f')+"("+price+"%)"
            else:
                updown = "▼"+format(abs(agio), '.2f')+" ("+price+"%)"
        else:
            updown = "-"
    except:
        pass
    table.insert('', 1, values=(nopen, nhigh, nlow, nvolume, nclose, updown))

    win1_re_btn = tk.Button(frame2, text="重新整理", width=10, command=now)
    win1_re_btn.pack(pady=5)


def chart():
    plt.close()
    for widget in frame2.winfo_children():
        widget.destroy()
    now_switch = 0

    today = datetime.date.today()
    month = today.month-6
    yestermonth = today.replace(month=month)
    # 抓取前天收盤價
    try:
        df = yf.download(f"{stockname1.get()}.TW", yestermonth)
        yclose = float(df.values[0][4])
    except:
        df = yf.download(f"{stockname1.get()}.TWO", yestermonth)

    # for i in range(len(df)):
    daylist = []
    for i in range(0, len(df.index)):
        daylist.append(df.index[i])
    df.insert(0, column="Date", value=daylist)
    # print(df)

    df['Date'] = pd.to_datetime(df['Date'], format='%Y-%m-%d')
    df.set_index('Date', inplace=True)
    mc = mpf.make_marketcolors(up='r', down='g', inherit=True)
    s = mpf.make_mpf_style(base_mpf_style='yahoo', marketcolors=mc)
    fig, ax = mpf.plot(df, type='candlestick', volume=True,
                       style=s, mav=(5, 20), returnfig=True, block=False)
    chart_canvas = FigureCanvasTkAgg(fig, master=frame2)
    chart_canvas.draw()
    chart_canvas.get_tk_widget().pack()


def information():
    plt.close()
    now_switch = 0
    for widget in frame2.winfo_children():
        widget.destroy()
    ######## 抓取基本資料########
    try:
        # 取得前一天日期
        today = datetime.date.today()
        day = today.day-1
        yesterday = today.replace(day=day)
        df = yf.download(f"{stockname1.get()}.TW", yesterday)
        yclose = float(df.values[0][4])
        url = f'https://tw.stock.yahoo.com/quote/{stockname1.get()}.TW/profile'
    except:
        url = f'https://tw.stock.yahoo.com/quote/{stockname1.get()}.TWO/profile'

    html = requests.get(url)
    sp = BeautifulSoup(html.text, 'lxml')
    infolist = sp.find_all(
        'span', class_='As(st) Bxz(bb) Pstart(12px) Py(8px) Bgc($c-gray-hair) C($c-primary-text) Flx(n) W(104px) W(120px)--mobile W(152px)--wide Miw(u) Pend(12px) Mend(0)')
    infolist2 = sp.find_all('div', class_='Py(8px) Pstart(12px) Bxz(bb)')
    company_info = []
    for i in range(len(infolist)):
        company_info.append([infolist[i].text, infolist2[i].text])
        if infolist[i].text == "主要經營業務":
            break

    columns = ["項目", "內容"]
    table = tk.ttk.Treeview(
        frame2, height=28, show="headings", columns=columns)

    # 卷軸
    # yscrollbar=tk.Scrollbar(frame2)
    # yscrollbar.pack(side="right",fill='y')
    # table.configure(yscrollcommand=yscrollbar.set)

    table.column("項目", width=125, anchor="center")
    table.column("內容", width=400)

    table.heading("項目", text="項目")
    table.heading("內容", text="內容")
    table.pack()
    for i in company_info:
        if i[0] == '主要經營業務':
            tampinfo = list(company_info[-1][1].split("\r\n"))
            table.insert('', 'end', values=(i[0], tampinfo[0]))
            for i in range(1, len(tampinfo)):
                table.insert('', 'end', values=("", tampinfo[i]))
        else:
            table.insert('', 'end', values=(i[0], i[1]))


######### 股票資訊##########
win1 = tk.Tk()  # 股票資訊
win1.title('個股資訊')
win1.geometry('900x700+800+150')
show_switch = 0
f1 = tk.Frame(win1)
f1.pack()
stockname = tk.StringVar()
stockname1 = tk.StringVar()
stockname2 = tk.StringVar()
stockname3 = tk.StringVar()
stockname4 = tk.StringVar()
win1.protocol('WM_DELETE_WINDOW', cleanup_function)
label_stockname = tk.Label(f1, textvariable=stockname)
label_stockname.grid(row=0, column=0, padx=20)

label_stockname1 = tk.Label(f1, textvariable=stockname1)
label_stockname1.grid(row=1, column=0, padx=20)

label_stockname2 = tk.Label(f1, textvariable=stockname2)
label_stockname2.grid(row=0, column=1, rowspan=2, padx=20)

label_stockname3 = tk.Label(f1, textvariable=stockname3)
label_stockname3.grid(row=0, column=2, padx=20)

label_stockname4 = tk.Label(f1, textvariable=stockname4)
label_stockname4.grid(row=1, column=2, padx=20)


menu = tk.Frame(win1, width=500, height=50)
menu.pack()

# 按鈕
btn1 = tk.Button(menu, text="分時", width=10, command=now)
btn2 = tk.Button(menu, text="圖表", width=10, command=chart)
btn3 = tk.Button(menu, text="基本資料", width=10, command=information)

frame2 = tk.Frame(win1)
frame2.pack()
# 抓當前價格

############ 股票輸入、收藏#############
# 创建主窗口
win2 = tk.Toplevel(win1)
win2.title("股票搜尋與收藏")
win2.geometry('500x505+300+300')

# 创建输入框，用来输入股票代码
input_field = tk.Entry(win2)
input_field.grid(row=0, column=0)

# 创建按钮，用来搜索股票
search_button = tk.Button(win2, text="搜索", command=search)
search_button.grid(row=0, column=1)

# 创建按钮，用来加入收藏
add_button = tk.Button(win2, text="加入收藏", command=update_table)
add_button.grid(row=0, column=2)

# 重新整理按鈕
add_button = tk.Button(win2, text="重新整理", command=stock_re_button)
add_button.grid(row=0, column=3)
# 查不到股票的錯誤訊息
error_msg = tk.StringVar()
label_error_msg = tk.Label(win2, textvariable=error_msg, text="hiiii")
label_error_msg.grid(row=1, column=0)

# 创建表格，用来显示收藏的股票
total_list_stock = []
try:
    f = open('collect.txt')
    total_list_stock = list(f.read().split())
    f.close()
except:
    pass
frame33 = tk.Frame(win2)
frame33.grid(row=2, column=0, columnspan=4)
columns = ["代碼", "名稱", "成交價", "成交量"]

collet_table = ttk.Treeview(
    frame33, height=15, show="headings", columns=columns)
yscrollbar = tk.Scrollbar(frame33)
yscrollbar.pack(side="right", fill='y')
collet_table.configure(yscrollcommand=yscrollbar.set)

collet_table.column("代碼", width=100, anchor="center")
collet_table.column("名稱", width=120, anchor="center")
collet_table.column("成交價", width=70, anchor="ne")
collet_table.column("成交量", width=70, anchor="ne")

collet_table.heading("代碼", text="代碼")
collet_table.heading("名稱", text="名稱")
collet_table.heading("成交價", text="成交價")
collet_table.heading("成交量", text="成交量")

collet_table.pack()
# collet_table.insert(tk.END, "{:　<4}{:　<6}{:　<6}{:　>4}".format("代碼","名稱","成交價","成交量"))

collet_table.bind('<<TreeviewSelect>>', touch)
collet_table.pack()
# 刪除按鈕
del_button = tk.Button(frame33, text="刪除", width=10, command=delbox)
del_button.pack(pady=10)

# 股票重新整理
# stockre = threading.Thread(target=stock_re)
# stockre.setDaemon(True)
# stockre.start()
stock_re_button()
# 给加入收藏按钮添加事件处理函数

# 运行主窗口
win1.mainloop()
