import tkinter
import tkinter.filedialog
import glob
import os
import tkinter.ttk as ttk
import pandas as pd
import numpy as np
import time
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import configparser
import shutil
import sys
import csv
import json
import pathlib
from tkinter import messagebox
import subprocess
#config読み込み
config_ini = configparser.ConfigParser()
config_ini.read('config.ini', encoding='utf-8')
USERNAME=config_ini["default"]["username"]
PASSWORD=config_ini["default"]["password"]
CLASS=config_ini["default"]["class"]
CLASS_CODE=config_ini["default"]["class_code"]
BASE_URL="https://ist.ksc.kwansei.ac.jp/~ishiura/cgi-bin/submit/main.cgi?"

#ダウンロードパスを絶対パスに変換する
#パスが存在しなければ現在のディレクトリを設定する
downloadPath= config_ini["default"]["downloadPath"]
try:
    downloadPath=downloadPath.resolve()
except:
    downloadPath=pathlib.Path.cwd()
if not os.path.exists(downloadPath):
    downloadPath=pathlib.Path.cwd()
downloadPath=str(downloadPath)

#option
options = Options()
UA = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_4) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/13.1 Safari/605.1.15'
options.add_argument('--user-agent=' + UA)
options.add_argument("--headless")
options.add_experimental_option('prefs', {'download.prompt_for_download': False,})

#ログイン認証の突破用にurlの書き換え
def make_login_url(url):
    return "https://"+USERNAME+":"+PASSWORD+"@"+url[8:]

#スクレイピングで名簿と課題番号を取得してくる
def prepare_file():

    driver = webdriver.Chrome(ChromeDriverManager().install(),options=options)
    driver.command_executor._commands["send_command"] = (
        'POST',
        '/session/$sessionId/chromium/send_command'
    )
    driver.execute(
        'send_command',
        params={
            'cmd': 'Page.setDownloadBehavior',
            'params': {'behavior': 'allow', 'downloadPath': downloadPath}
        }
    )
    info=["S="+CLASS_CODE,"act_class=1","c="+CLASS]
    url=make_login_url(BASE_URL+"&".join(info))
    #テーブルの取得
    driver.get(url)
    trs=driver.find_elements(By.TAG_NAME,"tr")
    #学籍番号と名前のpd.DataFrameを作成
    df=[]    
    for i in trs[4:-1]:
        row=i.text.split()
        df.append([int(row[0])," ".join(row[1:3])])
    df=pd.DataFrame(df,columns=["student number","name"])
    df.to_csv(f"{downloadPath}/student_class{CLASS}.csv",index=False,encoding='shift_jis')

    #課題名をすべて取得し、辞書型で保存
    assign_dic=dict()
    for i in trs[3].text.split()[3:-1]:
        info=["S="+CLASS_CODE,"act_report=1","c="+CLASS,"r="+i]
        url=make_login_url(BASE_URL+"&".join(info))
        driver.get(url)
        ths=driver.find_elements(By.TAG_NAME,"tr")[3].find_elements(By.TAG_NAME,"th")
        new=[]
        for j in ths[6:-1]:
            new.append(j.text.split()[0])
        assign_dic[i]=new
    tf = open(f"{downloadPath}/assign_dic.json", "w")
    json.dump(assign_dic,tf)
    tf.close()
    driver.close()

#ファイルの準備用（初回のみ）
if not (os.path.exists(f"{downloadPath}/student_class{CLASS}.csv") and os.path.exists(f"{downloadPath}/assign_dic.json")):
    prepare_file()
if not os.path.exists(f"{downloadPath}/test_case"):
    os.makedirs(f"{downloadPath}/test_case")

# 外部ファイルからの読み込みのための関数
def resource_path(filename):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, filename)
    return os.path.join(os.path.abspath("."), filename)

#生徒の名簿を読み込む
student_dic = {}
student_path = f"{downloadPath}/student_class{CLASS}.csv"
with open(student_path,encoding='shift_jis') as f:
    reader = csv.reader(f)
    for i, row in enumerate(reader):
        if i==0:
            continue
        student_dic[int(row[0])] = row[1]


student_list = []
for key, val in student_dic.items():
    student_list.append(" ".join([str(key), val]))

#課題一覧を取得
tf = open(f"{downloadPath}/assign_dic.json", "r")
assign_dic = json.load(tf)
assign_list=list(assign_dic.keys())
#print(assign_dic)
#print(assign_list)

inp_df_temp=pd.DataFrame([["","","出力が違います",""]],columns=["入力例","出力例","コメント","(コマンドライン引数)"])


class Application(tkinter.Tk):
    def __init__(self):
        super().__init__()
        
        self.editor_width = 900
        self.editor_height = 500
        
        self.geometry("1260x670")
        self.title("採点システム")
        
          # 画像表示のキャンバス作成と配置
        # self.editor = tkinter.Canvas(
        #     self,
        #     width=self.editor_width,
        #     height=self.editor_height,
        #     bg="gray",
        #     scrollregion = (0,0,0,0)
        # )
        self.frame = tkinter.Frame(
            self,
            #width=100,#self.editor_width,
            #height=100#self.editor_height
        )
        self.frame.grid(column=0, row=0)
        # self.frame.place()
        self.editor = tkinter.Text(
            self.frame,
            wrap="none",
            width=120,#self.editor_width,
            height=30#self.editor_height
        )
        self.editor.pack()
        # self.editor.grid(row=0, column=0)
        
        # スクロールバーを作成
        xbar = tkinter.Scrollbar(orient='horizontal')  # バーの方向
        ybar = tkinter.Scrollbar(orient='vertical')  # バーの方向
        # キャンバスにスクロールバーを配置
        xbar.grid(
            row=1, column=0,  # キャンバスの下の位置を指定
            sticky=tkinter.W + tkinter.E  # 左右いっぱいに引き伸ばす
        )
        ybar.grid(
            row=0, column=1,  # キャンバスの右の位置を指定
            sticky=tkinter.N + tkinter.S  # 上下いっぱいに引き伸ばす
        )
        # スクロールバーのスライダーが動かされた時に実行する処理を設定
        xbar.config(command=self.editor.xview)
        ybar.config(command=self.editor.yview)
        # キャンバススクロール時に実行する処理を設定
        self.editor.config(xscrollcommand=xbar.set)
        self.editor.config(yscrollcommand=ybar.set)
        

        # 画像表示のキャンバス作成と配置
        self.editor_code = tkinter.Canvas(
            self,
            width=self.editor_width,
            height=150,
            scrollregion = (0,0,0,0)
            #scrollregion = (0,0,3000,600)
        )
        self.editor_code.grid(row=2, column=0)
        # スクロールバーを作成
        xbar2 = tkinter.Scrollbar(orient='horizontal')  # バーの方向
        ybar2 = tkinter.Scrollbar(orient='vertical')  # バーの方向
        # キャンバスにスクロールバーを配置
        xbar2.grid(
            row=3, column=0,  # キャンバスの下の位置を指定
            sticky=tkinter.W + tkinter.E  # 左右いっぱいに引き伸ばす
        )
        ybar2.grid(
            row=2, column=1,  # キャンバスの右の位置を指定
            sticky=tkinter.N + tkinter.S  # 上下いっぱいに引き伸ばす
        )
        # スクロールバーのスライダーが動かされた時に実行する処理を設定
        xbar2.config(command=self.editor_code.xview)
        ybar2.config(command=self.editor_code.yview)
        # キャンバススクロール時に実行する処理を設定
        self.editor_code.config(xscrollcommand=xbar2.set)
        self.editor_code.config(yscrollcommand=ybar2.set)
        
        # ボタンを配置するフレームの作成と配置
        self.button_frame = tkinter.Frame(width = 1200-self.editor_width,
                                          height = self.editor_height)
        self.button_frame.grid(row=0, column=2, rowspan=3)
        
        self.font = ("",12)
        
        #  # 課題コード選択ボックス
        # self.comb_assign_code = ttk.Combobox(self.button_frame,
        #                              values=assign_list, font=("",4), width=10)
        
        # # 課題選択ボックス
        # self.comb_assign = ttk.Combobox(self.button_frame,
        #                              values=student_list, font=("",), width=10)

        # 課題コード選択ボックス
        self.assign_code = ttk.Combobox(self.button_frame,
                                     values=assign_list, font=("",12), width=7)
        self.assign_code.current(0)
        self.assign_code.bind('<<ComboboxSelected>>', lambda event: self.assign.config(values=assign_dic[self.assign_code.get()]))
        # 課題選択ボックス
        self.assign = ttk.Combobox(self.button_frame,
                                     values=assign_dic[assign_list[0]], font=("",12), width=10)

        # ファイル読み込みボタンの作成と配置
        self.load_button = tkinter.Button(
            self.button_frame,
            text = "課題選択",
            command = self.push_download_button,
            bg = "light cyan",
            font=self.font, width=10
        )
        
        # 学生選択ボックス
        self.comb_student = ttk.Combobox(self.button_frame,
                                     values=student_list, font=("",11), width=20)
        # 学生選択ボタン
        self.choice_button = tkinter.Button(
            self.button_frame,
            text = "選択",
            command = self.push_choice_button,
            bg = "light cyan",
            font=self.font, width=10
        )
        
        # 前、次へボタンの作成
        self.before_button = tkinter.Button(
            self.button_frame,
            text = "前へ",
            command = self.push_before_button,
            bg = "light cyan",
            font=self.font, width = 15
        )
        self.next_button = tkinter.Button(
            self.button_frame,
            text = "次へ",
            command = self.push_next_button,
            bg = "light cyan",
            font=self.font, width = 15
        )
        
        # 個別に実行と戻すボタン
        self.execute_button = tkinter.Button(
            self.button_frame,
            text = "個別に実行",
            command = self.push_execute_button,
            bg = "light cyan",
            font = self.font, width=20
        )
        self.back_button = tkinter.Button(
            self.button_frame,
            text = "戻す",
            command = self.push_back_button,
            bg = "light cyan",
            font = self.font, width = 10
        )
        
        # OK、NGボタンの作成と配置
        self.ok_button = tkinter.Button(
            self.button_frame,
            text = "OK",
            command = self.push_ok_button,
            bg = "LightSkyBlue",
            font = self.font, width = 15
        )
        self.ng_button = tkinter.Button(
            self.button_frame,
            text = "NG",
            command = self.push_ng_button,
            bg = "IndianRed1",
            font = self.font, width = 15
        )
        
        # ファイル出力ボタンの作成と配置
        self.output_button = tkinter.Button(
            self.button_frame,
            text = "ファイル出力",
            command = self.push_output_button,
            bg = "seaGreen1",
            font=self.font
        )
         # ファイル出力ボタンの作成と配置
        self.setting_button= tkinter.Button(
            self.button_frame,
            text = "設定",
            command = self.push_setting_button,
            bg = "gray",
            font=self.font
        )
         # ファイル出力ボタンの作成と配置
        self.init_button= tkinter.Button(
            self.button_frame,
            text = "初期化",
            command = self.push_init_button,
            bg = "yellow",
            font=self.font
        )
        self.current_num = 0

        # 修正要求課題選択ボックス
        self.fix_assign = ttk.Combobox(self.button_frame,
                                     values=assign_list, font=("",12), width=7)
        # 修正要求ボタンの作成と配置
        self.fixrequest_button = tkinter.Button(
            self.button_frame,
            text = "修正要求",
            command = self.push_fixrequest_button,
            bg = "IndianRed1",
            font = self.font
        )
    
        self.editor_obj= None
        self.output_files=[]
        self.files = None
        self.file_len = 0
        self.ASSIGN=""
        self.ASSIGN_NUM=""
        self.CODE=None
        self.message = "課題を選択してください"
        
        self.message_st = tkinter.StringVar() #文字更新用のStringVarを定義
        self.message_st.set(self.message)
        self.msg_canvas = tkinter.Canvas(self.button_frame, height=40, width=1200-self.editor_width,
                                         background="white")
        self.lbl_message = tkinter.Label(self.button_frame,
                                         textvariable=self.message_st, background="white",
                                         font=self.font)
        
        
        self.student = ""
        self.lbl_student = tkinter.Label(self.button_frame,
                                         text=self.student, background="white",
                                         font=self.font)
        
        self.txt_set = set()
        self.txt_list = list(self.txt_set)
        self.comb_txt = ttk.Combobox(self.button_frame,
                                     values=self.txt_list, font=self.font)
        #入力例を格納
        self.input_df=inp_df_temp[:]
        #コード実行結果を格納
        self.result_code=None

        self.df = None
        self.output_file=None
        self.editor_code_obj=None
        self.time_limit=3
        #self.df = self.mk_df()
        self.state_dict={"4":"OK","1":"OK","0":"NG",np.nan:"未採点","未提出":"未提出"}

        # gridでウェイジェットの配置
        pady = 12
        self.msg_canvas.grid(row=0, column=0, columnspan=6, padx=2, pady=pady, 
                              sticky=tkinter.W+tkinter.E)
        self.lbl_message.grid(row=0, column=0, columnspan=6, padx=2, pady=pady, 
                              sticky=tkinter.W+tkinter.E)
        self.assign_code.grid(row=1, column=0, columnspan=2, padx=2, pady=pady)
        #ttk.Label(self.button_frame, text='の',font=("",12)).grid(row=1, column=1, columnspan=1, padx=0, pady=pady)
        self.assign.grid(row=1, column=2, columnspan=2, padx=10, pady=pady)
        
        self.load_button.grid(row=1, column=4, columnspan=2, padx=2, pady=pady)
        self.lbl_student.grid(row=2, column=0, columnspan=6, padx=10, pady=pady,
                              sticky=tkinter.W+tkinter.E)
        self.comb_student.grid(row=3, column=0, columnspan=5, padx=10, pady=pady)
        self.choice_button.grid(row=3, column=5, columnspan=1, padx=10, pady=pady)
        self.before_button.grid(row=4, column=0, columnspan=3, padx=10, pady=pady)
        self.next_button.grid(row=4, column=3, columnspan=3, padx=10, pady=pady)
        self.execute_button.grid(row=5, column=0, columnspan=4, padx=10, pady=pady)
        self.back_button.grid(row=5, column=4, columnspan=2, padx=10, pady=pady)
        self.comb_txt.grid(row=6, column=0, columnspan=6, padx=10, pady=pady,
                      sticky=tkinter.W+tkinter.E)
        self.ok_button.grid(row=7, column=0, columnspan=3, padx=10, pady=pady)
        self.ng_button.grid(row=7, column=3, columnspan=3, padx=10, pady=pady)
        self.output_button.grid(row=8, column=0, columnspan=6, padx=10, pady=pady,
                                sticky=tkinter.W+tkinter.E)
        self.setting_button.grid(row=9, column=0, columnspan=6, padx=10, pady=pady,
                                sticky=tkinter.W+tkinter.E)
        self.init_button.grid(row=10, column=0, columnspan=6, padx=10, pady=pady,
                                sticky=tkinter.W+tkinter.E)
        self.fix_assign.grid(row=11, column=0, columnspan=2, padx=2, pady=pady)
        self.fixrequest_button.grid(row=11, column=2, columnspan=6, padx=10, pady=pady,
                                sticky=tkinter.W+tkinter.E)
        
        # 採点ファイルの拡張子
        self.file_extension = None
        
        self.msg_box()
        
    def edit(self):
        # root = tkinter.Tk()
        self.editor.delete('1.0', tkinter.END)
        filepath = self.files[self.current_num]
        # print(filepath)
        with open(filepath, "r", encoding="UTF-8") as open_file:
            text = open_file.read()
            self.editor.insert(tkinter.END, text)
        # root.title(f'Text Files - {filepath}')
        #　ファイル名から学籍番号と氏名を取得し、表示する
        basename = os.path.splitext(os.path.basename(self.files[self.current_num]))[0]
        self.student = basename + " " + student_dic[int(basename)]
        self.msg_box()
           
    def push_download_button(self):
        week = self.assign_code.get()
        assign = self.assign.get()
        if week not in assign_dic:
            self.message = "alert_課題名が不正です"
        elif assign not in assign_dic[week]:
            self.message = "alert_課題名が不正です"
        else:
            self.message = "alert_ダウンロードを行います"
            self.msg_box()
            try:
                self.ASSIGN=assign
                self.ASSIGN_NUM = week
                self.output_file=f"{downloadPath}/eval_{self.ASSIGN}.xlsx"
                if os.path.exists(downloadPath+"/"+self.ASSIGN+'-'+str(CLASS)):
                    ret = messagebox.askyesno("確認", "すでに課題が存在します。\n再ダウンロードしますか？")
                    if ret == False:
                        self.push_load_button()
                        return
                info=["S="+CLASS_CODE,"act_download_multi=1","c="+CLASS,"r="+self.ASSIGN_NUM,"i="+self.ASSIGN]
                url=make_login_url(BASE_URL+"&".join(info))
                self.message = "ダウンロードを行います"
                
                #webdriver起動
                self.msg_box()
                driver = webdriver.Chrome(ChromeDriverManager().install(),options=options)
                driver.command_executor._commands["send_command"] = (
                    'POST',
                    '/session/$sessionId/chromium/send_command'
                )
                driver.execute(
                    'send_command',
                    params={
                        'cmd': 'Page.setDownloadBehavior',
                        'params': {'behavior': 'allow', 'downloadPath': downloadPath}
                    }
                )
                
                #テーブルの取得
                driver.get(url)
                time.sleep(1)
                driver.close()
                self.message = "alert_ダウンロードが終了しました"
            except:
                self.message = "alert_課題名が不正です"
            self.push_load_button()
        self.msg_box()
            
    def push_load_button(self,f=True):
        self.message = "alert_読み込んでいます"
        self.msg_box()
        if f:
            shutil.unpack_archive(downloadPath+"/"+self.ASSIGN+'-'+str(CLASS)+'.zip', downloadPath)
        folder_path = downloadPath+"/"+self.ASSIGN+'-'+str(CLASS)
        self.file_extension = self.ASSIGN.split(".")[-1]
        if self.file_extension == "c":
            self.files = glob.glob(f"{folder_path}/*.{self.file_extension}")
            # フォルダ内にcファイルが存在するか確認
            self.file_len = len(self.files)
            if self.file_len==0:
                basename = ""
            else:
                basename = os.path.splitext(os.path.basename(self.files[0]))[0]
            #cファイルがプロ2のものか判定（ファイル名が8桁の数字であることが条件）
            if self.file_len!=0 and len(basename)==8 and basename.isdigit():
                self.current_num = 0
                self.edit()
                self.CODE=self.ASSIGN
                self.message = ""
                if os.path.exists(self.output_file):
                    ret = messagebox.askyesno("確認", "すでに出力ファイルが存在します。\n前回の続きから再開しますか？")
                    if ret == True:
                        self.df=self.read_df()
                    else:
                        self.df = self.mk_df()
                else:
                    self.df = self.mk_df()
                self.push_setting_button()
                ret = messagebox.askyesno("確認", "自動で採点を行います。\n結果が上書きされますがよろしいですか？")
                self.result_code=self.execute_C(ret)
                if self.result_code is not None:
                    self.show_result(self.result_code[int(basename)])
                    self.message = "alert_コードの実行が完了しました！"
                self.output_files.append(self.ASSIGN)
            else:
                self.files = None
                self.result_code = None
                self.df = None
                self.message = "alert_フォルダが適切ではありません"
        elif self.file_extension == "txt":
            self.files = glob.glob(f"{folder_path}/*.{self.file_extension}")
            self.file_len = len(self.files)
            if self.file_len==0:
                basename = ""
            else:
                basename = os.path.splitext(os.path.basename(self.files[0]))[0]
            # ファイルがプロ2のものか判定（ファイル名が8桁の数字であることが条件）
            if self.file_len!=0 and len(basename)==8 and basename.isdigit():
                self.current_num = 0
                self.message = ""
                self.edit()
                self.df = self.mk_df()
            else:
                self.files = None
                self.message = "alert_フォルダが適切ではありません"
                self.student = ""
        self.msg_box()
        
    def push_choice_button(self):
        self.message = "alert_ファイルが存在しません"
        if self.files is not None:
            value = self.comb_student.get()
            value = value.split(" ")[0]
            for i, file in enumerate(self.files):
                if value in file:
                    self.message = ""
                    if self.result_code is not None:
                        self.show_result(self.result_code[int(value)])
                    if int(value) in self.df.index:
                        self.message=self.state_dict[self.df["判定"][int(value)]]
                    self.comb_txt.delete(0, tkinter.END)
                    if self.df["判定"][int(value)]=="0":
                        self.comb_txt.insert(tkinter.END, self.df["コメント"][int(value)])
                    self.current_num = i
                    self.edit()
                    break
        self.msg_box()
        
    def push_before_button(self):
        if self.files is not None:
            self.current_num = (self.current_num-1) % self.file_len
            filename = self.files[self.current_num]
            basename = os.path.splitext(os.path.basename(filename))[0]
            self.edit()
            self.message = ""
            ind = self.return_list_index(basename)
            if ind is not None:
                self.comb_student.delete(0, tkinter.END)
                self.comb_student.insert(tkinter.END, student_list[ind])
            self.comb_txt.delete(0, tkinter.END)
            if ind is not None  and self.df["判定"][int(basename)]=="0":               
                self.comb_txt.insert(tkinter.END, self.df["コメント"][int(basename)])
            if int(basename) in self.df.index:
                self.message=self.state_dict[self.df["判定"][int(basename)]]
                if self.result_code is not None:
                    self.show_result(self.result_code[int(basename)])
            
        else:
            self.message = "alert_ファイルが存在しません"
        self.msg_box()
        
    
    def push_next_button(self):
        if self.files is not None:
            self.current_num = (self.current_num+1) % self.file_len
            filename = self.files[self.current_num]
            basename = os.path.splitext(os.path.basename(filename))[0]
            self.edit()
            ind = self.return_list_index(basename)

            if ind is not None:
                self.comb_student.delete(0, tkinter.END)
                self.comb_student.insert(tkinter.END, student_list[ind])
            self.comb_txt.delete(0, tkinter.END)
            if ind is not None  and self.df["判定"][int(basename)]=="0":
                self.comb_txt.insert(tkinter.END, self.df["コメント"][int(basename)])
            self.message = ""
            if int(basename) in self.df.index:
                    self.message=self.state_dict[self.df["判定"][int(basename)]]
                    if self.result_code is not None:
                        self.show_result(self.result_code[int(basename)])
            
            
        else:
            self.message = "alert_ファイルが存在しません"
        self.msg_box()
        
    # 個別にcファイルを実行して、出力結果を表示する
    # 編集した場合は編集したプログラムを実行する、学生のプログラムに戻す場合は戻すボタンを押す
    def push_execute_button(self):
        if self.files is not None and self.file_extension == "c":
            ta_filepath = f"{os.path.splitext(self.files[self.current_num])[0]}_TA.{self.file_extension}"
            with open(ta_filepath, "w") as save_file:
                text = self.editor.get("1.0", tkinter.END)
                save_file.write(text)
            
            """
            ここに個別にcファイルを実行するためのコードを記述する
            ta_filepath を実行する
            """
            tests=self.input_df.values
            ta_exe_path=f"{os.path.splitext(self.files[self.current_num])[0]}_TA"
            ta_file_result=""
            #コンパイルを行ってエラーの場合はスルーする
            if subprocess.run("gcc " + ta_filepath + " -o " + ta_exe_path, shell=True).returncode:
                ta_file_result+="コンパイルエラーです。\n"
            else:
                for n,test in enumerate(tests):
                    input,output,comment,command=test
                    ta_file_result+=f"テストケース{n+1}\n input=[{input}]\n*---*---*---*---*---*---*---*---*---*\n"
                    try:
                        p=subprocess.run(ta_exe_path+" "+command,input=input,stdout=subprocess.PIPE,shell=False,encoding='utf-8',timeout=self.time_limit)
                        if p.returncode:
                            ta_file_result+="実行時エラーです。\n" 
                        else:
                            ta_file_result+=p.stdout+"\n"
                    except UnicodeDecodeError:
                        ta_file_result+="UnicodeDecodeError\n"
                    except Exception:
                        ta_file_result+="終わりません\n"

            self.show_result(ta_file_result)
            # 実行が終わると削除される
            os.remove(ta_filepath)
            
    # 表示するテキストを元のテキスト（学生のもの）に戻す
    def push_back_button(self):
        if self.files is None:
            return
        filename = self.files[self.current_num]
        basename = os.path.splitext(os.path.basename(filename))[0]
        self.show_result(self.result_code[int(basename)])
        if self.files is not None:
            self.edit()
        
    def push_ok_button(self):
        if self.df is not None:
            filename = self.files[self.current_num]
            basename = int(os.path.splitext(os.path.basename(filename))[0])
            comment = self.comb_txt.get()
            # コメントのコンボボックスにコメントを追加
            self.txt_set.add(comment)
            self.txt_list = list(self.txt_set)
            self.comb_txt.config(values=self.txt_list)
            self.df.loc[basename, "判定"] = "4"
            self.df.loc[basename, "コメント"] = comment
            self.df.loc[basename, "system"] = "!ok"
            self.comb_txt.delete(0, tkinter.END)
            self.message = "OK"
            # コンボボックスの表示を変更
            ind = self.return_list_index(str(basename))
            if ind is not None:
                student_list[ind] = student_list[ind].replace(" *", "")
                self.comb_student.config(values=student_list)
                self.comb_student.delete(0, tkinter.END)
                self.comb_student.insert(tkinter.END, student_list[ind])
                self.push_next_button()
            self.msg_box()
          
    def push_ng_button(self):
        if self.df is not None:
            filename = self.files[self.current_num]
            basename = int(os.path.splitext(os.path.basename(filename))[0])
            comment = self.comb_txt.get()
            # コメントのコンボボックスにコメントを追加
            self.txt_set.add(comment)
            self.txt_list = list(self.txt_set)
            self.comb_txt.config(values=self.txt_list)
            if len(comment)==0:
                self.message = "alert_コメントを入力してください"
            else:
                self.df.loc[basename, "判定"] = "0"
                self.df.loc[basename, "コメント"] = comment
                self.df.loc[basename, "system"] = "!NG/"+comment
                self.comb_txt.delete(0, tkinter.END)
                self.message = "NG"
                # コンボボックスの表示を変更
                ind = self.return_list_index(str(basename))
                if ind is not None:
                    student_list[ind] = student_list[ind].replace(" *", "")
                    self.comb_student.config(values=student_list)
                    self.comb_student.delete(0, tkinter.END)
                    self.comb_student.insert(tkinter.END, student_list[ind])
                    self.push_next_button()
            self.msg_box()
    
    #出力ボタン
    def push_output_button(self):
        if self.df is not None:
            null_num = self.df.isnull().sum()["判定"]
            if null_num == 0:
                ret = messagebox.askyesno("確認", "excelに出力しますか？")
                if ret == True:
                    self.output(self.output_file,self.ASSIGN_NUM,self.ASSIGN)
                else:
                    self.message = "alert_ファイルを出力しませんでした"
            else:
                ret = messagebox.askyesno("確認", "未採点の課題があります。\nexcelに出力しますか？")
                if ret == True:
                    self.output(self.output_file,self.ASSIGN_NUM,self.ASSIGN)
                else:
                    self.message = "alert_ファイルを出力しませんでした"
            if len(self.output_files)>1:
                for out in self.output_files:
                    if out==self.ASSIGN:
                        continue
                    ret = messagebox.askyesno("確認", out+"にも同じ結果を出力しますか？")
                    if ret == True:
                        self.output(f"{downloadPath}/eval_{out}.xlsx",self.ASSIGN_NUM,out)
                    else:
                        self.message = "alert_" + out+ "にはファイルを出力しませんでした"
                self.msg_box()
                
    def push_setting_button(self):
        '''モーダルダイアログボックスの作成'''
        dlg_modal = tkinter.Toplevel(self)
        dlg_modal.title("Modal Dialog") # ウィンドウタイトル
        dlg_modal.geometry("500x500")   # ウィンドウサイズ(幅x高さ)

        # モーダルにする設定
        dlg_modal.grab_set()        # モーダルにする
        dlg_modal.focus_set()       # フォーカスを新しいウィンドウをへ移す
        n=5
        items=[list(i) for i in self.input_df.values]+[[None,None,None,None] for i in range(n-len(self.input_df.values))]
        forms=[]
         # 項目をつくる
        def open_case():
            typ = [('Excelfile',['*.xlsx','*.xls'])] 
            dir = f"{downloadPath}/test_case"
            file = tkinter.filedialog.askopenfilename(filetypes = typ, initialdir = dir) 
            if os.path.exists(file):
                self.input_df=pd.read_excel(file).fillna("")
                items=[list(i) for i in self.input_df.values]+[[None,None,None,None] for i in range(n-len(self.input_df.values))]
                for i in range(0, n):
                    for j in range(len(list(self.input_df.columns))):
                        if items[i][j] is not None:
                            forms[i][j].delete("1.0", tkinter.END)
                            forms[i][j].insert("1.0",items[i][j])
                        else:
                            forms[i][j].delete("1.0", tkinter.END)
        def save():
            for i in range(n):
                for j in range(4):
                    if i==0 or j!=0 or (forms[i][j].get("1.0", tkinter.END)!="\n" and forms[i][j].get("1.0", tkinter.END)!=""):
                        items[i][j]=forms[i][j].get("1.0", tkinter.END).strip()
                    else:
                        items[i][j]=None
            self.input_df=pd.DataFrame(items,columns=["入力例","出力例","コメント","(コマンドライン引数)"]).dropna(how='any')
            filename=filename_input.get()
            if filename!="":
                self.input_df.to_excel(f"{downloadPath}/test_case/{filename}.xlsx",index=False)

        def time_limit_save():
            try:
                t=float(time_input.get())
                self.time_limit=t
            except:
                time_label["text"]="数値を入力→"
        
        cols=list(self.input_df.columns)
        for i in range(0, len(cols)):
            label_item = ttk.Label(dlg_modal,
            text=cols[i])
            label_item.grid(row=0, column=i+1,padx=5)
        
        for i in range(0, n):
            num = tkinter.Label(dlg_modal, text=str(i+1), font=("MSゴシック", "15"))
            num.grid(row=i+1,column=0)
            item1 = tkinter.Text(dlg_modal,height=1,width=10)
            item1.grid(row=i+1,column=1,padx=10,pady=10)
            if items[i][0] is not None:
                item1.delete("1.0", tkinter.END)
                item1.insert("1.0",items[i][0])
            item2 = tkinter.Text(dlg_modal,height=1,width=10)
            item2.grid(row=i+1,column=2,padx=10,pady=10)
            if items[i][1] is not None:
                item2.delete("1.0", tkinter.END)
                item2.insert("1.0",items[i][1])
            item3 = tkinter.Text(dlg_modal,height=1,width=20)
            item3.grid(row=i+1,column=3,padx=10,pady=10)
            if items[i][2] is not None:
                item3.delete("1.0", tkinter.END)
                item3.insert("1.0",items[i][2])
            item4 = tkinter.Text(dlg_modal,height=1,width=10)
            item4.grid(row=i+1,column=4,padx=10,pady=10)
            if items[i][3] is not None:
                item4.delete("1.0", tkinter.END)
                item4.insert("1.0",items[i][3])
            forms.append([item1,item2,item3,item4])
        
        load_button = tkinter.Button(
            dlg_modal,
            text = "読込",
            command = open_case,
            bg = "light cyan",
            font=self.font
        )
        load_button.grid(row=n+3,column=1)
        filename_label = tkinter.Label(dlg_modal, text="ファイル名", font=self.font)
        filename_label.grid(row=n+3,column=2)
        filename_input=tkinter.Entry(dlg_modal,width=20)
        if self.ASSIGN !="":
            filename_input.insert(tkinter.END,f"{self.ASSIGN}_test")
        else:
            filename_input.insert(tkinter.END,"A00_0")
        filename_input.grid(row=n+3,column=3)
        save_button = tkinter.Button(
            dlg_modal,
            text = "保存",
            command = save,
            bg = "light cyan",
            font=self.font
        )
        save_button.grid(row=n+3,column=4)
        time_title = tkinter.Label(dlg_modal, text="実行制限時間", font=self.font)
        time_title.grid(row=n+4,column=1)
        time_label = tkinter.Label(dlg_modal, text="入力→", font=self.font)
        time_label.grid(row=n+5,column=1)
        time_input=tkinter.Entry(dlg_modal,width=5)
        time_input.insert(tkinter.END,self.time_limit)
        time_input.grid(row=n+5,column=2)
        time_label = tkinter.Label(dlg_modal, text="秒", font=self.font)
        time_label.grid(row=n+5,column=3)
        save_button = tkinter.Button(
            dlg_modal,
            text = "実行時間を変更",
            command = time_limit_save,
            bg = "light cyan",
            font=self.font
        )
        save_button.grid(row=n+5,column=3)
        #print(self.input_df)
        dlg_modal.transient(self.master)   # タスクバーに表示しない
        # ダイアログが閉じられるまで待つ
        app.wait_window(dlg_modal)

    def push_init_button(self):
        ret = messagebox.askyesno("確認", "本当に初期化しますか？（出力されていない内容は破棄されます）")
        if ret == True:
            self.destroy()
            self.__init__()

    def push_fixrequest_button(self):
        assign_code = self.fix_assign.get()
        if assign_code not in assign_list:
            return
        ret = messagebox.askyesno("確認", f"{assign_code} に修正要求を出しますか？")
        if not ret:
            return
        self.stump_ok(assign_code, save_flag=ret)

    def output(self,output_file,ASSIGN_NUM,ASSIGN):
        self.df.to_excel(output_file)
        #print(self.df)
        self.message = "alert_excelファイルを出力しました"
        self.msg_box()
        ret = messagebox.askyesno("確認", "結果をレポートシステムに出力しますか？")
        if ret == True:
            info=["S="+CLASS_CODE,"act_report=1","c="+CLASS,"r="+ASSIGN_NUM,"eval_i="+ASSIGN]
            url=make_login_url(BASE_URL+"&".join(info))
            driver = webdriver.Chrome(ChromeDriverManager().install(),options=options)
            driver.get(url)
            flg = False
            for i in self.df.index.values:
                if type(self.df["system"][i]) is not str: 
                    continue
                if len(self.df["system"][i])==0:
                    continue
                flg = True
                a=driver.find_element(By.NAME,"eval:"+str(i)+":"+ASSIGN)
                a.clear()
                a.send_keys(self.df["system"][i])   
            if flg:
                a.send_keys(Keys.ENTER)
            driver.close()
    
    # code実行
    def execute_py(self,f):
        tests=self.input_df.values
        result={i:"提出されていません" for i in student_dic.keys()}
        for file in self.files:
            student_num=int(os.path.splitext(file)[0])
            result[student_num]=""
            #print(student_num)
            evaluation=None
            
            for n,test in enumerate(tests):
                input,output,comment,command=test
                result[student_num]+=f"テストケース{n+1}\n input=[{input}]\n"
                try:
                    p=subprocess.run("python "+file+" "+command,input=input,stdout=subprocess.PIPE,shell=False,encoding='cp932',timeout=self.time_limit)
                    if p.returncode:
                        result[student_num]+="エラーです。\n"
                        if f:
                            self.df.loc[student_num, "判定"] = "0"
                            self.df.loc[student_num, "コメント"] = "エラーです。"
                            self.df.loc[student_num, "system"] = "!NG/"+"エラーです。"
                            evaluation="エラーです。"
                            self.txt_set.add("エラーです。")
                            self.txt_list = list(self.txt_set)
                            self.comb_txt.config(values=self.txt_list)
                    else:
                        result[student_num]+=p.stdout+"\n"
                        if f:
                            if evaluation is None:
                                for out in output.split("\n"):
                                    if out not in p.stdout:
                                        self.df.loc[student_num, "判定"] = "0"
                                        self.df.loc[student_num, "コメント"] = comment
                                        self.df.loc[student_num, "system"] = "!NG/"+comment
                                        evaluation=comment
                                        self.txt_set.add(comment)
                                        self.txt_list = list(self.txt_set)
                                        self.comb_txt.config(values=self.txt_list)
                                        break
                except UnicodeDecodeError:
                    result[student_num]+="UnicodeDecodeError\n"
                    if f:
                        self.df.loc[student_num, "判定"] = "0"
                        self.df.loc[student_num, "コメント"] = "UnicodeDecodeError"
                        self.df.loc[student_num, "system"] = "!NG/"+"UnicodeDecodeError"
                        evaluation="UnicodeDecodeError"
                        self.txt_set.add("UnicodeDecodeError")
                        self.txt_list = list(self.txt_set)
                        self.comb_txt.config(values=self.txt_list)
                except Exception:
                    result[student_num]+="終わりません\n"
                    if f:
                        self.df.loc[student_num, "判定"] = "0"
                        self.df.loc[student_num, "コメント"] = "終わりません"
                        self.df.loc[student_num, "system"] = "!NG/"+"終わりません"
                        evaluation="終わりません"
                        self.txt_set.add("終わりません")
                        self.txt_list = list(self.txt_set)
                        self.comb_txt.config(values=self.txt_list)
            if f and evaluation is None:
                self.df.loc[student_num, "判定"] = "4"
                self.df.loc[student_num, "コメント"] = ""
                self.df.loc[student_num, "system"] = "!ok"
        return result
    
    # C言語のcode実行
    def execute_C(self,f):
        #print(self.df)
        tests=self.input_df.values
        result={i:"提出されていません" for i in student_dic.keys()}
        for file in self.files:
            student_num=int(os.path.split(file)[-1].split(".")[0])
            result[student_num]=""
            #print(student_num)
            evaluation=None
            #コンパイルを行ってエラーの場合はスルーする
            if subprocess.run("gcc " + file + " -o " + os.path.splitext(file)[0], shell=True).returncode:
                result[student_num]+="コンパイルエラーです。\n"
                if f:
                    self.df.loc[student_num, "判定"] = "0"
                    self.df.loc[student_num, "コメント"] = "コンパイルエラーです。"
                    self.df.loc[student_num, "system"] = "!NG/"+"コンパイルエラーです。"
                    evaluation="コンパイルエラーです。"
                    self.txt_set.add("コンパイルエラーです。")
                    self.txt_list = list(self.txt_set)
                    self.comb_txt.config(values=self.txt_list)
            else:
                for n,test in enumerate(tests):
                    input,output,comment,command=test
                    result[student_num]+=f"テストケース{n+1}\n input=[{input}]\n*---*---*---*---*---*---*---*---*---*\n"
                    try:
                        p=subprocess.run(os.path.splitext(file)[0]+" "+command,input=input,stdout=subprocess.PIPE,shell=False,encoding='utf-8',timeout=self.time_limit)
                        if p.returncode:
                            result[student_num]+="実行時エラーです。\n"
                            if f:
                                self.df.loc[student_num, "判定"] = "0"
                                self.df.loc[student_num, "コメント"] = "実行時エラーです。"
                                self.df.loc[student_num, "system"] = "!NG/"+"実行時エラーです。"
                                evaluation="実行時エラーです。"
                                self.txt_set.add("実行時エラーです。")
                                self.txt_list = list(self.txt_set)
                                self.comb_txt.config(values=self.txt_list)
                        else:
                            result[student_num]+=p.stdout+"\n"
                            if f:
                                if evaluation is None:
                                    for out in output.split("\n"):
                                        if out not in p.stdout:
                                            self.df.loc[student_num, "判定"] = "0"
                                            self.df.loc[student_num, "コメント"] = comment
                                            self.df.loc[student_num, "system"] = "!NG/"+comment
                                            evaluation=comment
                                            self.txt_set.add(comment)
                                            self.txt_list = list(self.txt_set)
                                            self.comb_txt.config(values=self.txt_list)
                                            break
                    except UnicodeDecodeError:
                        result[student_num]+="UnicodeDecodeError\n"
                        if f:
                            self.df.loc[student_num, "判定"] = "0"
                            self.df.loc[student_num, "コメント"] = "UnicodeDecodeError"
                            self.df.loc[student_num, "system"] = "!NG/"+"UnicodeDecodeError"
                            evaluation="UnicodeDecodeError"
                            self.txt_set.add("UnicodeDecodeError")
                            self.txt_list = list(self.txt_set)
                            self.comb_txt.config(values=self.txt_list)
                    except Exception:
                        result[student_num]+="終わりません\n"
                        if f:
                            self.df.loc[student_num, "判定"] = "0"
                            self.df.loc[student_num, "コメント"] = "終わりません"
                            self.df.loc[student_num, "system"] = "!NG/"+"終わりません"
                            evaluation="終わりません"
                            self.txt_set.add("終わりません")
                            self.txt_list = list(self.txt_set)
                            self.comb_txt.config(values=self.txt_list)
            if f and evaluation is None:
                self.df.loc[student_num, "判定"] = "4"
                self.df.loc[student_num, "コメント"] = ""
                self.df.loc[student_num, "system"] = "!ok"
            #print(self.df)
        return result
    
    # コードの出力結果を表示
    def show_result(self, val):
        if self.result_code is not None:
            if self.editor_code_obj is not None:
                self.editor_code.delete(self.editor_code_obj)
            self.editor_code_obj=self.editor_code.create_text(10, 10, anchor="nw", text=val,
                                                              fill="black", font=self.font)
            # テキストの境界ボックスを取得
            bbox = self.editor_code.bbox(self.editor_code_obj)
            # テキストの幅と高さを計算
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
            self.editor_code.config(scrollregion = (0,0,text_width+10,text_height))
    
    def msg_box(self):
        self.lbl_student.config(text=self.student)
        alert_flg = self.message.split("_")[0]
        if alert_flg == "alert":
            self.message = self.message.split("_")[1]
            self.message_st.set(self.message) #StringVarに反映
            message_x = 170
            message_y = 23
            while message_x >= 30:
                self.lbl_message.place_forget() #ラベル消去
                self.lbl_message.place(x=message_x, y=message_y) #ラベル再配置
                time.sleep(0.1)
                self.update()
                message_x -= 10
        else:
            self.message_st.set(self.message) #StringVarに反映
            self.lbl_message.config(text=self.message)
            self.lbl_message.place(x=30, y=23)
            
    def read_df(self):
        df=pd.read_excel(self.output_file, index_col=0)
        
        for file in self.files:
            basename = int(os.path.splitext(os.path.basename(file))[0])
            if df["判定"][basename] != "4":
                # コンボボックスの要素のうち未提出かok以外のものに*を記入
                ind = self.return_list_index(str(basename))
                student_list[ind] = student_list[ind].replace(" *", "") + " *"
        self.comb_student.config(values=student_list)
        self.txt_set=set(df["コメント"].dropna())
        self.txt_list = list(self.txt_set)
        self.comb_txt.config(values=self.txt_list)
        return df

    def mk_df(self):
        nan = np.full((len(student_list), 4), np.nan)
        df = pd.DataFrame(data=nan, index=student_dic.keys(), columns=["名前","判定", "コメント","system"])
        df["名前"] = student_dic.values()
        df["判定"] = "未提出"
        df["コメント"] = ""
        df["system"] = ""
        for file in self.files:
            basename = int(os.path.splitext(os.path.basename(file))[0])
            df.at[basename, "判定"] = np.nan
            # コンボボックスの要素のうち未提出以外のものに*を記入
            ind = self.return_list_index(str(basename))
            student_list[ind] = student_list[ind].replace(" *", "") + " *"
        self.comb_student.config(values=student_list)
        
        return df
    
    # 入力に一致するstudent_listの要素のインデックスを返す関数
    def return_list_index(self, txt):
        for i, val in enumerate(student_list):
            if txt in val:
                return i
        return None
    
    def stump_ok(self, assign,save_flag=False):
        driver = webdriver.Chrome(ChromeDriverManager().install(),options=options)
        driver.command_executor._commands["send_command"] = (
            'POST',
            '/session/$sessionId/chromium/send_command'
        )
        driver.execute(
            'send_command',
            params={
                'cmd': 'Page.setDownloadBehavior',
                'params': {'behavior': 'allow', 'downloadPath': downloadPath}
            }
        )
        info=["S="+CLASS_CODE,"act_report=1","c="+CLASS,"r="+assign]
        url=make_login_url(BASE_URL+"&".join(info))
        driver.get(url)
        trs=driver.find_elements(By.TAG_NAME,"tr")
        oks=[]
        ngs=[]
        ths=trs[3].find_elements(By.TAG_NAME,"th")
        cols=["late"]+[i.text.split('\n')[0] for i in ths[6:-1]]
        student=dict()
        for j in range(4,len(trs)-1):
            ths=trs[j].find_elements(By.TAG_NAME,"td")
            num=ths[0].text
            student[num]=[]
            f=-1
            if len(ths[4].find_elements(By.TAG_NAME,"font"))==0:
                student[num].append(0)
            else:
                student[num].append(1)
            for th in ths[6:-1]:
                if "○"in th.text or "△" in th.text:
                    if "ok" in th.text.lower():
                        student[num].append(4)
                        if f!=-1:
                            continue
                        f=1
                    elif "ng" in th.text.lower():
                        student[num].append(0)
                        if f==0:
                            continue
                        f=2
                    else:
                        student[num].append(None)
                        f=0
                else:
                    student[num].append(None)
            if f==1:
                oks.append(num)
            elif f==2:
                if "修正要求" not in ths[3].text:
                    ngs.append(num)
        df=pd.DataFrame.from_dict(student, orient="index", columns=cols)
        savename="pro2.xlsx"
        if os.path.exists(savename):
            with pd.ExcelWriter(savename,engine="openpyxl", mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer,sheet_name=f"{assign}", na_rep=0)
        else:
            with pd.ExcelWriter(savename,engine="openpyxl") as writer:
                df.to_excel(writer,sheet_name=f"{assign}", na_rep=0)
        #print(oks)
        #print(ngs)
        #print(df)
        if save_flag:
            info=["S="+CLASS_CODE,"act_report=1","c="+CLASS,"r="+assign,"eval_r="+assign]
            url=make_login_url(BASE_URL+"&".join(info))
            driver.get(url)
            for ok in oks:
                a=driver.find_element(By.NAME,"eval:"+str(ok)+":"+assign)
                a.clear()
                a.send_keys("!ok")
            if len(oks)!=0:
                a.send_keys(Keys.ENTER)
            for ng in ngs:
                info=["S="+CLASS_CODE,"act_reject=1","s="+ng,"r="+assign,"return_to=class "+CLASS+" report "+assign,"exec=1","revise=1"]
                url=make_login_url(BASE_URL+"&".join(info))
                driver.get(url)
            driver.close()
    
if __name__ == "__main__":
    app = Application()
    app.mainloop()