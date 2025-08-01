import pyautogui
import tkinter as tk
import tkinter.ttk as ttk
from datetime import datetime

'''
画面表示
    ・時計
    ・メール一覧
    ・タスク一覧
    ・時間割
    ・メモ
コマンド
    ・ctrl + P 画面の表示/非表示切り替え
    ・ctrl + s 画面表示内容の設定変更
    ・ctrl + r 再起動（MeMO編集時に）
    ・escape プログラムの終了

左端に黒い縦のフレームがあるので ctrl+p で非表示にしてほかのウインドウにフォーカスを置き、
こっちにフォーカスを戻す場合は、この左端のフレームに触れる。
'''
class Display():
    def __init__(self, 
                 task: list,
                 time_table: list,
                 time_lst: list,
                 until_td: int,
                 until_tm: int,
                 messages: list):
        
        self.tasks = task #タスク一覧
        self.time_table = time_table #時間割の教科名
        self.time_lst = time_lst #時間割の時間
        self.until_td = until_td #今日までのタスク数
        self.until_tm = until_tm #明日までのタスク数
        self.messages = messages #メール情報のリスト
        
        #画面サイズの取得
        self.w, self.h = pyautogui.size()

        #ウインドウ状態
        self.status = {"OPEN": True, 
                       "RESTART": False,
                       "SELECT": False,
                       "saved": ""}
        
        #各機能の表示 ON/OFF
        self.display = {"時計": True,
                        "メール": True,
                        "タスク": True,
                        "時間割": True,
                        "メモ": True}

        self.root = tk.Tk()
        self.root.geometry(f"{self.w}x{self.h}")

        self.root.overrideredirect(True) #タイトルバーの非表示
        self.root.attributes("-topmost", True) #常に前面に表示
        self.root["bg"] = "white"
        
        #白色を透明にする（透過）
        self.root.wm_attributes("-transparentcolor", "white")
        
        #キーボードイベント
        self.root.bind_all("<KeyPress>", self.key_event)

        #左端のフォーカス補助フレーム
        tk.Frame(self.root,
                 background="black",
                 width=2).pack(side=tk.LEFT, fill=tk.Y)
        
        self.frame_R = tk.Frame(self.root,
                                 background="white")
        self.frame_R.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.frame_L1 = tk.Frame(self.root,
                                 background="white")
        self.frame_L1.pack(anchor=tk.W)
        
        self.frame_L2 = tk.Frame(self.root,
                                 background="white")
        self.frame_L2.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.frame_L3 = tk.Frame(self.root,
                                 background="white")
        self.frame_L3.pack(side=tk.BOTTOM, expand=True, anchor=tk.CENTER)

        self.slct_frame = tk.Frame(self.frame_R, 
                                   background="white")
        
        self.call()
        
        #クリック時にテキストボックス以外がクリックされたならフォーカス解除
        def unfocus_textbox(event: tk.Event):
            widget = event.widget
            if not isinstance(widget, tk.Text):
                self.root.focus_set()
        self.root.bind("<Button-1>", unfocus_textbox)
        
        self.root.mainloop()
    
    def call(self): #呼び出し
        self.Clock()
        self.Task()
        self.Mail()
        self.TimeTable()
        self.Memo()

    def Task(self):
        if self.display["タスク"]:
            if len(self.tasks) > 0:
                #今日/明日までのタスクの件数を表示
                tk.Label(self.frame_L2,
                        text=f"今日まで：{self.until_td}件",
                        font=(None, 20),
                        bg="white",
                        fg="chartreuse").pack(anchor=tk.W)

                tk.Label(self.frame_L2,
                        text=f"明日まで：{self.until_tm}件",
                        font=(None, 20),
                        bg="white",
                        fg="chartreuse").pack(anchor=tk.W)
                
                #スクロール用キャンバス
                canvas = tk.Canvas(self.frame_L2, 
                                        background="white",
                                        width=self.w,
                                        height=int(self.h * 0.3),
                                        highlightthickness=0,
                                        bd=0)
                #内容を格納するフレーム
                sub_frame = tk.Frame(self.frame_L2,
                                        background="white")
                                        
                #スクロールバーのスタイル設定
                style = ttk.Style()
                style.theme_use("clam")
                style.configure("Vertical.TScrollbar",
                                background="DarkSeaGreen1",
                                arrowcolor="cyan",
                                bordercolor="green",
                                borderwidth=0.1)
                #有効時と非有効時の色
                style.map("Vertical.TScrollbar",
                        background=[("active", "DarkSeaGreen1"), 
                                    ("!active", "DarkSeaGreen")])
                #スクロールバー
                scroll = ttk.Scrollbar(self.frame_L2, 
                                        orient=tk.VERTICAL,
                                        command=canvas.yview,
                                        style="Vertical.TScrollbar")
                canvas.configure(yscrollcommand=scroll.set)

                scroll.pack(side=tk.LEFT, fill=tk.Y, anchor=tk.N)
                canvas.pack(fill=tk.BOTH, expand=True)
                #キャンバス上にサブフレームを設置
                canvas.create_window((0, 0), window=sub_frame, anchor="nw")    

                #マウスホイールでスクロール
                def on_mousewheel(event: tk.Event):
                    canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
                #キャンバスに入ったらマウスホイールをバインド
                def bind_to_mousewheel(event: tk.Event):
                    canvas.bind("<MouseWheel>", on_mousewheel)
                #キャンバスを離れたらマウスホイールを解除
                def unbind_from_mousewheel(event: tk.Event):
                    canvas.unbind_all("<MouseWheel>")
                #サブフレームのサイズ変更に応じてスクロール範囲を更新
                def on_frame_configure(event: tk.Event):
                    canvas.configure(scrollregion=canvas.bbox("all"))

                sub_frame.bind("<Configure>", on_frame_configure)
                canvas.bind("<Enter>", bind_to_mousewheel)
                canvas.bind("<Leave>", unbind_from_mousewheel)

                day_limit = ""
                #タスク一覧の表示
                for task in self.tasks:
                    if task[0] == "0":
                        day_limit = "[本日まで]"   
                    elif task[0] == "1":
                        day_limit = "[明日まで]"
                    
                    #タスクの締切日とタイトル
                    tk.Label(sub_frame,
                            text=f"{day_limit}{task[1]}：",
                            font=(None, 30), 
                            bg="white", 
                            fg="green").pack(anchor=tk.W)
                    
                    #内容表示
                    for content in task[2:]:
                        tk.Label(sub_frame,
                                text=f"{content}",
                                font=(None, 25),
                                bg="white",
                                fg="chartreuse3",
                                wraplength=int(self.w*0.4),
                                justify="left").pack(anchor=tk.W, padx=40)

            else:
                tk.Label(self.frame_L2,
                        text="タスクはありません",
                        font=(None, 35), bg="white", 
                        fg="green").pack(anchor=tk.W)
        
    def Mail(self):
        if self.display["メール"]:
            #スクロール用キャンバス
            canvas = tk.Canvas(self.frame_L3, 
                               background="white",
                               width=self.w,
                               height=int(self.h * 0.3),
                               highlightthickness=0,
                               bd=0)
            #内容を格納するフレーム
            sub_frame = tk.Frame(self.frame_L3,
                                    background="white")
            
            #スクロール設定
            style = ttk.Style()
            style.theme_use("clam")
            style.configure("Vertical.TScrollbar",
                            background="DarkSeaGreen1",
                            arrowcolor="cyan",
                            bordercolor="green",
                            borderwidth=0.1)
            style.map("Vertical.TScrollbar",
                    background=[("active", "DarkSeaGreen1"), 
                                ("!active", "DarkSeaGreen")])
            #スクロール
            scroll = ttk.Scrollbar(self.frame_L3, 
                                    orient=tk.VERTICAL,
                                    command=canvas.yview,
                                    style="Vertical.TScrollbar")
            canvas.configure(yscrollcommand=scroll.set)

            scroll.pack(side=tk.LEFT, fill=tk.Y, anchor=tk.N)
            canvas.pack(fill=tk.BOTH, expand=True)
            canvas.create_window((0, 0), window=sub_frame, anchor="nw")  

            #マウスホイールでスクロール
            def on_mousewheel(event: tk.Event):
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            #キャンバスに入ったらマウスホイールをバインド
            def bind_to_mousewheel(event: tk.Event):
                canvas.bind("<MouseWheel>", on_mousewheel)
            #キャンバスを離れたらマウスホイールを解除
            def unbind_from_mousewheel(event: tk.Event):
                canvas.unbind_all("<MouseWheel>")
            #サブフレームのサイズ変更に応じてスクロール範囲を更新
            def on_frame_configure(event: tk.Event):
                canvas.configure(scrollregion=canvas.bbox("all"))

            sub_frame.bind("<Configure>", on_frame_configure)
            canvas.bind("<Enter>", bind_to_mousewheel)
            canvas.bind("<Leave>", unbind_from_mousewheel)

            #メール一覧の表示
            for content in self.messages:
                #差出人
                tk.Label(sub_frame,
                        text=content["From"],
                        font=(None, 25), 
                        bg="white", 
                        fg="green").pack(anchor=tk.W)
                #件名
                tk.Label(sub_frame,
                        text=content["Subject"],
                        font=(None, 20), 
                        bg="white", 
                        fg="chartreuse3",
                        wraplength=int(self.w*0.4),
                        justify="left").pack(anchor=tk.W, padx=40)
            
    def Clock(self):
        if self.display["時計"]:
            def update_time():
                now_time = datetime.now().strftime("%Y/%m/%d  %H:%M:%S")
                self.clock_lbl["text"] = now_time
                #1000ms後に自信を再呼び出し
                self.root.after(1000, update_time)
            
            #時刻
            self.clock_lbl = tk.Label(self.frame_L1,
                                    font=("", 35),
                                    bg="white",
                                    fg="green1")
            self.clock_lbl.pack(anchor=tk.W)
            update_time()
    
    def TimeTable(self):
        if self.display["時間割"]:
            sub_frame = tk.Frame(self.frame_R,
                                background="white")
            sub_frame.pack(anchor=tk.E)
            
            #時間と時間割
            for (subject, time) in zip(self.time_table, self.time_lst):
                tk.Label(sub_frame,
                        text=f"{time} : {subject}",
                        font=(None, 30),
                        bg="white",
                        fg="green").pack(anchor=tk.W, padx=3)
        
    def Memo(self):
        if self.display["メモ"]:
            #テキストボックス
            self.txt_box = tk.Text(self.frame_R,
                            font=("", 17),
                            bg="white",
                            fg="green1",
                            width=int(self.w*0.03),
                            relief=tk.SOLID,
                            border=5,
                            insertbackground="green2")
            self.txt_box.pack(side=tk.BOTTOM, fill=tk.X, anchor=tk.E)
            
            #保存状態を参照
            self.txt_box.insert("1.0", self.status["saved"])
            
            #フォーカスが当たったときの背景色の変化
            def on_focus_in(event: tk.Event):
                self.txt_box.config(bg="gray23")
                self.txt_box["relief"] = tk.GROOVE
                self.txt_box["border"] = 10
            #フォーカスが外れた時の背景色の変化                
            def on_focus_out(event: tk.Event):
                self.txt_box.config(bg="white")
                self.txt_box["relief"] = tk.SOLID
                self.txt_box["border"] = 5

            self.txt_box.bind("<FocusIn>", on_focus_in)
            self.txt_box.bind("<FocusOut>", on_focus_out)

    def key_event(self, event: tk.Event):
        key = event.keysym
        #ctrlキーの判定
        ctrl = (event.state & 0x4) != 0
        
        if key == "Escape":
            #プログラムの終了
            self.root.destroy()
        
        elif key == "p" and ctrl:
            #画面収納、再表示
            if self.status["OPEN"]: #表示状態なら
                #現在の入力状態を保存
                self.status["saved"] = self.txt_box.get("1.0", "end -1c")
                
                #現在表示されているものを削除する
                for packed in self.frame_L1.pack_slaves():
                    packed.pack_forget()
                for packed in self.frame_L2.pack_slaves():
                    packed.pack_forget()
                for packed in self.frame_L3.pack_slaves():
                    packed.pack_forget()
                for packed in self.frame_R.pack_slaves():
                    packed.pack_forget()
                for packed in self.slct_frame.pack_slaves():
                    packed.pack_forget()
                
                self.status["OPEN"] = False
                self.status["SELECT"] = False

            else: #非表示なら
                self.call() #再表示
                self.status["OPEN"] = True
        
        elif key == "r" and ctrl:
            #再起動処理
            self.status["RESTART"] = True
            self.root.destroy()

        elif key == "s" and ctrl:
            #表示コンテンツの設定
            self.change_display()
    
    def change_display(self):   
        if self.status["SELECT"] == False: #設定画面が開いていないなら
            self.status["SELECT"] = True
            
            #選択用フレームを配置
            self.slct_frame.pack(pady=15)

            self.var = {} #チェックボックス取得結果格納用
            #チェックボックスを作成して表示
            for item in self.display:
                self.var[item] = tk.BooleanVar()
                self.var[item].set(self.display[item])

                tk.Checkbutton(self.slct_frame, 
                            text=item,
                            variable=self.var[item],
                            font=("", 20),
                            background="white",
                            fg="green1").pack(side=tk.RIGHT)

            #設定の適用
            def get_choice():
                self.status["SELECT"] = False
                self.status["OPEN"] = True
                
                #MeMOの現在状況を保存
                self.status["saved"] = self.txt_box.get("1.0", "end -1c")

                #設定の保存
                for item in self.var:
                    self.display[item] = self.var[item].get()
                
                #現在表示されているものを削除
                for packed in self.frame_L1.pack_slaves():
                    packed.pack_forget()
                for packed in self.frame_L2.pack_slaves():
                    packed.pack_forget()
                for packed in self.frame_L3.pack_slaves():
                    packed.pack_forget()
                for packed in self.frame_R.pack_slaves():
                    packed.pack_forget()
                for packed in self.slct_frame.pack_slaves():
                    packed.pack_forget()
                
                self.call() #再呼び出し

            tk.Button(self.slct_frame,
                      text="適用",
                      font=("", 20),
                      background="white",
                      fg="green2",
                      command=get_choice).pack(padx=10)
        
        else:
            #設定画面が開いているなら閉じる
            
            self.status["SELECT"] = False

            for packed in self.slct_frame.pack_slaves():
                packed.pack_forget()