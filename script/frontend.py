import sqlite3
import sys
import threading
import time
import tkinter as tk
from os import environ
from tkinter import BOTTOM, RIGHT, E, S, messagebox, ttk

import customtkinter as ctk
from dotenv import load_dotenv
import webbrowser


import backend
import gsheet


class App():
    def __init__(self) -> None:
        # 環境変数の読み込み
        load_dotenv(dotenv_path="../setting/.env")
        self.root = ctk.CTk()
        self.prog_window = None
        self.general_font = ("Arial", 20)
        self.width = None
        self.height = None
        #登録処理が終わっているかの確認用（Trueが終わっているという意味）
        self.finished = False

        # このAppはインターネットに接続していないと使えない。接続が確認できない場合はFalseを返し強制終了
        ping = backend.Ping()
        result = ping.ping()
        if result == True:
            pass
        elif result == False:
            messagebox.showerror(
                "警告", "インターネットに接続されていないため、アプリケーションを起動できませんでした。インターネットに接続上で起動するようお願いします。")
            sys.exit()

    
    def screen_size(self, screen, width: int, height: int):
        """
        画面サイズを指定したら、自動的に中央に配置されるように計算する
        """
        screen.update_idletasks()
        ww = screen.winfo_screenwidth()
        wh = screen.winfo_screenheight()
        screen.geometry(str(width)+"x"+str(height)+"+" +
                        str(int(ww/2-width/2))+"+"+str(int(wh/2-height/2)))

    def progress(self, content: str):
        # ウィンドウの作成
        self.prog_window = ctk.CTkToplevel()
        self.screen_size(self.prog_window, 350, 150)  # 画面サイズの決定
        self.prog_window.overrideredirect(True)  # 最大化・最小化無効
        label = ctk.CTkLabel(self.prog_window, text=content,
                             text_font=self.general_font)
        label.pack(side="top", fill="both",pady=15)
        progressbar = ctk.CTkProgressBar(
            self.prog_window, progress_color="#006400", mode="indeterminate")
        progressbar.pack(side="bottom", fill="x", padx=10,pady=15)
        progressbar.start()
        # もう一つの登録処理が終了するまで待機
        while not self.finished:
            time.sleep(0.1)
        progressbar.stop()          # プログレスバーの停止
        self.prog_window.destroy()  # ウィンドウの消去

    def register(self, get_list: list):
        """
        sqlite3とGoogleのスプレッドシートにそれぞれ登録する。また、登録する前に念のためネットワーク接続を確認している。
        """

        # 登録処理が終わっているかの確認用（Trueが終わっているという意味）
        self.finished = False

        try:
            # DBの登録前にインターネットに接続出来ているかを確認
            ping = backend.Ping()
            result = ping.ping()
            if result == True:
                # sqlite3に追加
                ope = backend.DB_operation()
                ope.db_register(get_list)
                # Googleスプレッドシート側に追加
                google_ope = gsheet.Google_spreadsheet_operation()
                google_ope.sp_insert(get_list)
                messagebox.showinfo("通知", "登録が完了しました")
            elif result == False:
                messagebox.showerror("警告", "インターネットに接続していないため、登録ができません")
        except sqlite3.IntegrityError:
            messagebox.showerror("通知", f"管理番号が重複しています")
        except Exception as e:
            messagebox.showerror("通知", f"エラーが発生しました。\n{e}")
        finally:
            # 登録処理が終了したことにする
            self.finished = True

    def delete(self):
        """
        DBから削除を行う。また、同時にGoogleのスプレッドシートからも削除する。削除前に、インターネットに接続しているかも確認する。
        """

        # 登録処理が終わっているかの確認用（Trueが終わっているという意味）
        self.finished = False
        try:
            # DBの削除前にインターネットに接続出来ているかを確認
            ping = backend.Ping()
            result = ping.ping()
            if result == True:
                ope = backend.DB_operation()
                ope.db_delete(self.del_control_num_entry.get())
                # Googleスプレッドシート側に追加
                google_ope = gsheet.Google_spreadsheet_operation()
                google_ope.sp_delete(self.del_control_num_entry.get())
                messagebox.showinfo("通知", "削除が完了しました")
            elif result == False:
                messagebox.showerror(
                    "警告", "インターネットに接続していないため、登録ができません")
        except Exception as e:
            messagebox.showerror("エラー", f"エラーが発生しました。\n{e}")
        finally:
            # 登録処理が終了したことにする
            self.finished = True

    def waste(self):
        """
        廃棄するときに「廃棄」を登録すための関数。
        """

        # 登録処理が終わっているかを確認用（Trueが終わっているという意味）
        self.finished = False
        try:
            # DBの登録前にインターネットに接続出来ているかを確認
            ping = backend.Ping()
            result = ping.ping()
            if result == True:
                ope = backend.DB_operation()
                ope.db_waste(self.del_control_num_entry.get())
                # Googleスプレッドシート側に追加
                google_ope = gsheet.Google_spreadsheet_operation()
                google_ope.sp_waste(self.del_control_num_entry.get())
                messagebox.showinfo("通知", "登録が完了しました")
            elif result == False:
                messagebox.showerror(
                    "警告", "インターネットに接続していないため、登録ができません")
        except Exception as e:
            messagebox.showerror("エラー", f"エラーが発生しました。\n{e}")
        finally:
            # 登録処理が終了したことにする
            self.finished = True

    def window_setup(self, root: ctk.CTk) -> None:
        """
        このAppのテーマカラーやタイトル、画面サイズの設定を行う。
        """

        root.set_appearance_mode("System")          # OSのテーマカラーに合わせる。(dark or light)
        ctk.set_default_color_theme("dark-blue")    # テーマカラーをダークブルーに統一

        # 画面サイズを取得
        self.width = root.winfo_screenwidth()
        self.height = root.winfo_screenheight()
        # 画面サイズと最小サイズの決定
        root.geometry(f"{self.width}x{self.height}+0+0")
        root.minsize(1000, 750)

        root.grid_columnconfigure(0, weight=1)
        root.grid_rowconfigure(0, weight=1)

    def main_screen(self, root: ctk.CTk) -> None:
        # タイトルの変更
        root.title("高額物品管理システム(メインメニュー)")

        # フレームの用意(base)
        frame = ctk.CTkFrame(master=root, width=1000, height=600,
                             corner_radius=15, border_color="#347ab6", border_width=3)
        frame.grid(row=0, column=0)

        # メインメニューにおけるタイトル表示
        Title_label = ctk.CTkLabel(
            master=frame, text="高額物品管理システム", text_font=("MS ゴシック", 45))
        Title_label.grid(row=0, column=0, padx=150, pady=50)

        # ボタン用意
        register_button = ctk.CTkButton(master=frame, width=300, height=60, border_width=0, corner_radius=8,
                                        text="備品管理登録", text_font=self.general_font, command=lambda: self.register_screen(root))
        register_button.grid(row=1, column=0, pady=20)
        delete_button = ctk.CTkButton(master=frame, width=300, height=60, border_width=0, corner_radius=8,
                                      text="備品登録削除/廃棄", text_font=self.general_font, command=lambda: self.delete_screen(root))
        delete_button.grid(row=2, column=0, pady=20)
        settings_button = ctk.CTkButton(master=frame, width=300, height=60, border_width=0, corner_radius=8,
                                        text="備品一覧", text_font=self.general_font, command=lambda: self.list_screen(root))
        settings_button.grid(row=3, column=0, pady=20)
        explain_button = ctk.CTkButton(master=frame, width=300, height=60, border_width=0, corner_radius=8,
                                       text="説明", text_font=self.general_font, command=lambda: self.list_screen(root))
        explain_button.grid(row=4, column=0, pady=20)
        end_button = ctk.CTkButton(master=frame, width=300, height=60, border_width=0, corner_radius=8,
                                   text="終了", hover_color="red", text_font=self.general_font, command=self.end_exe)
        end_button.grid(row=5, column=0, pady=20)

    def register_screen(self, root: ctk.CTk):
        # タイトルの変更
        root.title("高額物品管理システム(登録)")

        # .envファイルから団体名を読み込む。このときカンマ区切りでリストに追加
        org = list(environ['all_org'].split(','))

        # 登録画面用のフレーム用意(base)
        frame = ctk.CTkFrame(master=root, width=self.width,
                             height=self.height, corner_radius=0, border_width=0)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.tkraise()  # 作成したフレームを一番上に持ってくる

        # 入力欄用のフレーム用意
        entry_frame = ctk.CTkFrame(
            master=frame, width=1200, height=800, corner_radius=0, border_width=0)
        entry_frame.pack(pady=50, anchor=tk.CENTER)

        # 入力欄の作成とそれの説明用のラベル
        goods_label = ctk.CTkLabel(
            master=entry_frame, text="物品名", text_font=self.general_font)
        goods_label.grid(row=1, column=0, padx=15, pady=15)
        self.goods_entry = ctk.CTkEntry(master=entry_frame, placeholder_text="抽象的な名前", width=700,
                                        height=50, border_width=2, corner_radius=8, text_font=self.general_font)
        self.goods_entry.grid(row=1, column=1, padx=15, pady=15)

        goods_detail_label = ctk.CTkLabel(
            master=entry_frame, text="物品の型番", text_font=self.general_font)
        goods_detail_label.grid(row=2, column=0, padx=15, pady=15)
        self.goods_detail_entry = ctk.CTkEntry(master=entry_frame, placeholder_text="その商品が特定できる程度",
                                               width=700, height=50, border_width=2, corner_radius=8, text_font=self.general_font)
        self.goods_detail_entry.grid(row=2, column=1, padx=15, pady=15)

        use_group_label = ctk.CTkLabel(
            master=entry_frame, text="使用団体", text_font=self.general_font)
        use_group_label.grid(row=3, column=0, padx=15, pady=15)
        self.use_group_box = ctk.CTkComboBox(master=entry_frame, width=500, height=50, values=org,
                                             dropdown_text_font=self.general_font, text_font=self.general_font, corner_radius=8)
        self.use_group_box.set("団体名を選んでください")
        self.use_group_box.grid(row=3, column=1, padx=15, pady=15)

        name_label = ctk.CTkLabel(
            master=entry_frame, text="代表者名", text_font=self.general_font)
        name_label.grid(row=4, column=0, padx=15, pady=15)
        self.name_entry = ctk.CTkEntry(master=entry_frame, placeholder_text="顧問の名前", width=500,
                                       height=50, border_width=2, corner_radius=8, text_font=self.general_font)
        self.name_entry.grid(row=4, column=1, padx=15, pady=15)

        purchase_date_label = ctk.CTkLabel(
            master=entry_frame, text="購入日", text_font=self.general_font)
        purchase_date_label.grid(row=5, column=0, padx=15, pady=15)
        self.purchase_date_entry = ctk.CTkEntry(master=entry_frame, placeholder_text="2022年01月01日→20220101(8桁)",
                                                width=500, height=50, border_width=2, corner_radius=8, text_font=self.general_font)
        self.purchase_date_entry.grid(row=5, column=1, padx=15, pady=15)

        control_num_label = ctk.CTkLabel(
            master=entry_frame, text="管理番号", text_font=self.general_font)
        control_num_label.grid(row=6, column=0, padx=15, pady=15)
        self.control_num_entry = ctk.CTkEntry(master=entry_frame, placeholder_text="別途内部資料参考",
                                              width=500, height=50, border_width=2, corner_radius=8, text_font=self.general_font)
        self.control_num_entry.grid(row=6, column=1, padx=15, pady=15)

        note_label = ctk.CTkLabel(
            master=entry_frame, text="備考欄", text_font=self.general_font)
        note_label.grid(row=7, column=0, padx=15, pady=15)
        self.note_entry = ctk.CTkEntry(master=entry_frame, placeholder_text="特にない場合は「None」と入力", width=700,
                                       height=50, border_width=0, corner_radius=8, text_font=self.general_font)
        self.note_entry.grid(row=7, column=1, padx=15, pady=15)

        # 登録と終了ボタン用のframeを作成する
        button_frame = ctk.CTkFrame(
            master=frame, corner_radius=0, border_width=0)
        button_frame.pack(pady=5, anchor=tk.CENTER)

        register_button = ctk.CTkButton(master=button_frame, width=100, height=40, text_font=(
            "MS ゴシック", 14), border_width=0, corner_radius=8, text="登録", command=lambda: thread_work())
        register_button.pack(padx=10, pady=10, side=tk.RIGHT)

        end_button = ctk.CTkButton(master=button_frame, width=200, height=40, border_width=0, corner_radius=8,
                                   text="メインメニューに戻る",  hover_color="red", text_font=("MS ゴシック", 14), command=lambda: frame.destroy())
        end_button.pack(padx=10, pady=10, side=tk.RIGHT)

        def thread_work():
            # 入力値の取得
            get_list = []  # 入力欄の値を格納するリスト
            for i in self.goods_entry.get(), self.goods_detail_entry.get(), self. use_group_box.get(), self.name_entry.get(), self.purchase_date_entry.get(), self.control_num_entry.get(), self.note_entry.get():
                # 空欄があると登録できないようにしている。問題がなければリストに追加
                if i == "":
                    messagebox.showwarning("通知", "空欄の所があります")
                    return None
                elif str.isdecimal(self.purchase_date_entry.get()) == False:
                    messagebox.showerror(
                        "入力値エラー", "購入日の入力値が正しく入力されていません\nex)数字以外が入力されている")
                    return None
                elif len(self.purchase_date_entry.get()) != 8:
                    messagebox.showerror(
                        "入力値エラー", "８桁の数字で入力してください")
                    return None
                else:
                    get_list.append(i)
            # buttonを無効化
            register_button['state'] = "disable"
            end_button['state'] = "disable"
            # プログレスバーと登録処理をマルチスレッドで実行
            register_thread = threading.Thread(target=self.register, args=(get_list,))
            progress_thread = threading.Thread(
                target=self.progress, args=("データベースに登録中...",))
            register_thread.start()
            progress_thread.start()
            # buttonを有効化
            register_button['state'] = "normal"
            end_button['state'] = "normal"

    def delete_screen(self, root: ctk.CTk):
        # タイトルの変更
        root.title("高額物品管理システム(削除・廃棄)")

        # 削除画面用のフレームの用意(base)
        frame = ctk.CTkFrame(master=root, width=self.width,
                             height=self.height, corner_radius=0, border_width=0)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.tkraise()  # 作成したフレームを一番上に持ってくる。

        # Entry用のフレーム用意
        entry_frame = ctk.CTkFrame(
            master=frame, width=1200, height=800, corner_radius=0, border_width=0)
        entry_frame.pack(pady=50, anchor=tk.CENTER)

        control_num_label = ctk.CTkLabel(
            master=entry_frame, text="管理番号を入力してください", text_font=self.general_font)
        control_num_label.grid(row=1, column=0, padx=15, pady=15)
        self.del_control_num_entry = ctk.CTkEntry(master=entry_frame, placeholder_text="", width=400,
                                                  height=50, border_width=2, corner_radius=8, text_font=self.general_font)
        self.del_control_num_entry.grid(row=2, column=3, padx=15, pady=15)

        # 登録と終了ボタン用のframeを作成する
        button_frame = ctk.CTkFrame(
            master=frame, corner_radius=0, border_width=0)
        button_frame.pack(pady=5, anchor=tk.CENTER)

        delete_button = ctk.CTkButton(master=button_frame, width=100, height=40, text_font=(
            "MS ゴシック", 14), border_width=0, corner_radius=8, text="完全に削除", fg_color="#ff758c", command=lambda: thread_delete())
        delete_button.pack(padx=10, pady=10, side=tk.RIGHT)

        waste_button = ctk.CTkButton(master=button_frame, width=100, height=40, text_font=(
            "MS ゴシック", 14), border_width=0, corner_radius=8, text="廃棄", fg_color="#ff7eb3", command=lambda: thread_waste())
        waste_button.pack(padx=10, pady=10, side=tk.RIGHT)

        end_button = ctk.CTkButton(master=button_frame, width=200, height=40, border_width=0, corner_radius=8,
                                   text="メインメニューに戻る",  fg_color="red", text_font=("MS ゴシック", 14), command=lambda: frame.destroy())
        end_button.pack(padx=10, pady=10, side=tk.RIGHT)

        def thread_delete():
            if self.del_control_num_entry.get() == "":
                messagebox.showwarning("通知", "何も入力されていません")
                return None
            else:
                confirm = messagebox.askokcancel(
                    "確認", "データベースから完全に削除をしてもよろしいですか?\n※登録を間違えたとき以外は使用しないでください。")
                if confirm == True:
                    # buttonを無効化
                    delete_button['state'] = "disable"
                    waste_button['state'] = "disable"
                    end_button['state'] = "disable"
                    # プログレスバーと登録処理をマルチスレッドで実行
                    delete_thread = threading.Thread(target=self.delete)
                    progress_thread = threading.Thread(
                        target=self.progress, args=("データベースに登録中...",))
                    delete_thread.start()
                    progress_thread.start()
                    # buttonを有効化
                    delete_button['state'] = "normal"
                    waste_button['state'] = "normal"
                    end_button['state'] = "normal"

        def thread_waste():
            if self.del_control_num_entry.get() == "":
                messagebox.showwarning("通知", "何も入力されていません")
                return None
            else:
                confirm = messagebox.askokcancel(
                    "確認", "データベースに対して廃棄済みとして登録しますがよろしいですか?")
                if confirm == True:
                    # buttonを無効化
                    delete_button['state'] = "disable"
                    waste_button['state'] = "disable"
                    end_button['state'] = "disable"
                    # プログレスバーと登録処理をマルチスレッドで実行
                    waste_thread = threading.Thread(target=self.waste)
                    progress_thread = threading.Thread(
                        target=self.progress, args=("データベースに登録中...",))
                    waste_thread.start()
                    progress_thread.start()
                    # buttonを有効化
                    delete_button['state'] = "normal"
                    waste_button['state'] = "normal"
                    end_button['state'] = "normal"

    def list_screen(self, root: ctk.CTk):
        """
        データベースに登録されているデータを表形式で表示する。
        """
        root.minsize(1920, 1080)
        frame = ctk.CTkFrame(master=root, width=self.width,
                             height=self.height, corner_radius=0, border_width=0)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.tkraise()

        list_frame = ctk.CTkFrame(
            master=frame, corner_radius=0, border_width=0)
        list_frame.pack(pady=10, ipadx=4, ipady=4)

        column_all = ("物品名", "物品の型番", "使用団体",  "代表者名",
                      "購入日", "管理番号", "備考", "廃棄")
        list_tree = ttk.Treeview(
            list_frame, columns=column_all, selectmode="none", height=45)
        ctk_textbox_scrollbar_y = ctk.CTkScrollbar(
            list_frame, orientation="vertical", command=list_tree.yview, hover=True, width=20)
        ctk_textbox_scrollbar_y.pack(side=RIGHT, fill='y')
        list_tree["yscrollcommand"] = ctk_textbox_scrollbar_y.set
        ctk_textbox_scrollbar_x = ctk.CTkScrollbar(
            list_frame, orientation="horizontal", command=list_tree.xview, hover=True, width=10)
        ctk_textbox_scrollbar_x.pack(side=BOTTOM, fill='x')
        list_tree["xscrollcommand"] = ctk_textbox_scrollbar_x.set
        list_tree.pack(side='left')

        style = ttk.Style()
        style.configure("Treeview.Heading", font=("MS ゴシック", 24))
        style.configure("Treeview", font=("Arial", 16))

        list_tree.heading('#0', text='')
        list_tree.heading('物品名', text='物品名', anchor='center')
        list_tree.heading('物品の型番', text='物品の型番', anchor='center')
        list_tree.heading('購入日', text='購入日', anchor='center')
        list_tree.heading('管理番号', text='管理番号', anchor='center')
        list_tree.heading('使用団体', text='使用団体', anchor='center')
        list_tree.heading('代表者名', text='代表者名', anchor='center')
        list_tree.heading('備考', text='備考', anchor='center')
        list_tree.heading('廃棄', text='廃棄', anchor='center')

        list_tree.column('#0', width=0, stretch=False)
        list_tree.column('物品名', anchor='center', width=200, stretch=False)
        list_tree.column('物品の型番', anchor='center', width=250, stretch=False)
        list_tree.column('購入日', anchor='center', width=125, stretch=False)
        list_tree.column('管理番号', anchor='w', width=140, stretch=False)
        list_tree.column('使用団体', anchor='center', width=280, stretch=False)
        list_tree.column('代表者名', anchor='center', width=200, stretch=False)
        list_tree.column('備考', anchor='w', width=550, stretch=False)
        list_tree.column('廃棄', anchor='center', width=75, stretch=False)

        # list_tree.insert(parent='', index='end',values=(1, 'KAWASAKI',80))

        ope = backend.DB_operation()
        a = ope.db_extract()

        i = 0
        for r in a:
            list_tree.insert("", "end", tags=i, values=r)
            if i & 2:
                # tagが奇数(レコードは偶数)の場合のみ、背景色の設定
                list_tree.tag_configure(i, background="red")
            i += 1

        button_frame = ctk.CTkFrame(
            master=frame, corner_radius=0, border_width=0)
        button_frame.pack(pady=10)
        end_button = ctk.CTkButton(master=button_frame, width=200, height=40, border_width=0, corner_radius=8,
                                   text="メインメニューに戻る",  fg_color="red", hover_color="red", text_font=("MS ゴシック", 14), command=lambda: return_button())
        end_button.pack(padx=10, pady=5, side=tk.RIGHT)

        def return_button():
            frame.destroy()
            root.minsize(1000, 750)

    def end_exe(self):
        """
        終了するかを確認したうえで、完全に終了する
        """
        confirm = messagebox.askyesno("確認", "終了してもよろしいでしょうか")
        if confirm == True:
            sys.exit()


app = App()
app.window_setup(app.root)
app.main_screen(app.root)

app.root.mainloop()