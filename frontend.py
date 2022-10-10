# -*- coding: utf-8 -*-

import sys
import tkinter as tk
from tkinter import messagebox, S, E
import backend
import gsheet
import customtkinter as ctk
import sqlite3
from tkinter import ttk


class App():
    def __init__(self) -> None:
        self.root = ctk.CTk()

        self.general_font = ("Arial", 20)
        self.width = None
        self.height = None

        # このAppはインターネットに接続していないと使えない。接続が確認できない場合はFalseを返し強制終了
        ping = backend.Ping()
        result = ping.ping()
        if result == True:
            pass
        elif result == False:
            messagebox.showerror(
                "警告", "インターネットに接続されていないため、アプリケーションを起動できませんでした。インターネットに接続上で起動するようお願いします。")
            sys.exit()

    def window_setup(self, root: ctk.CTk) -> None:
        """
        このAppのテーマカラーやタイトル、画面サイズの設定を行う。
        """

        # OSのテーマカラーに合わせる。(dark or light)
        root.set_appearance_mode("System")
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
        goods_entry = ctk.CTkEntry(master=entry_frame, placeholder_text="抽象的な名前", width=700,
                                   height=50, border_width=2, corner_radius=8, text_font=self.general_font)
        goods_entry.grid(row=1, column=1, padx=15, pady=15)

        goods_detail_label = ctk.CTkLabel(
            master=entry_frame, text="物品の型番", text_font=self.general_font)
        goods_detail_label.grid(row=2, column=0, padx=15, pady=15)
        goods_detail_entry = ctk.CTkEntry(master=entry_frame, placeholder_text="その商品が特定できる程度",
                                          width=700, height=50, border_width=2, corner_radius=8, text_font=self.general_font)
        goods_detail_entry.grid(row=2, column=1, padx=15, pady=15)

        use_group_label = ctk.CTkLabel(
            master=entry_frame, text="使用団体", text_font=self.general_font)
        use_group_label.grid(row=3, column=0, padx=15, pady=15)
        use_group_box = ctk.CTkComboBox(master=entry_frame, width=450, height=50, values=["学生会執行部", "体育祭実行委員会", "環境委員会", "広報委員会", "図書委員会", "産技祭実行委員会", "柔道部", "ラグビー部", "剣道部", "卓球部", "硬式野球部", "バスケットボール部", "硬式テニス部", "ソフトテニス部", "水泳部", "陸上競技部", "バドミントン部", "サッカー部",
                                        "バレーボール部", "弓道部", "茶道部", "写真部", "電気通信部", "吹奏楽部", "石灰費(ラグビー部)", "石灰費(サッカー部)", "石灰費(硬式野球部)", "石灰費(陸上競技部)", "熱中症対策費(学生会執行部)", "コロナ対策費(学生会執行部)", "慶弔費", "新入生卒業生記念品(学生会執行部)"], dropdown_text_font=self.general_font, text_font=self.general_font, corner_radius=8)
        use_group_box.set("団体名を選んでください")
        use_group_box.grid(row=3, column=1, padx=15, pady=15)

        name_label = ctk.CTkLabel(
            master=entry_frame, text="代表者名", text_font=self.general_font)
        name_label.grid(row=4, column=0, padx=15, pady=15)
        name_entry = ctk.CTkEntry(master=entry_frame, placeholder_text="顧問の名前", width=450,
                                  height=50, border_width=2, corner_radius=8, text_font=self.general_font)
        name_entry.grid(row=4, column=1, padx=15, pady=15)

        purchase_date_label = ctk.CTkLabel(
            master=entry_frame, text="購入日", text_font=self.general_font)
        purchase_date_label.grid(row=5, column=0, padx=15, pady=15)
        purchase_date_entry = ctk.CTkEntry(master=entry_frame, placeholder_text="2022年11月20日→20221001(8桁)",
                                           width=450, height=50, border_width=2, corner_radius=8, text_font=self.general_font)
        purchase_date_entry.grid(row=5, column=1, padx=15, pady=15)

        control_num_label = ctk.CTkLabel(
            master=entry_frame, text="管理番号", text_font=self.general_font)
        control_num_label.grid(row=6, column=0, padx=15, pady=15)
        control_num_entry = ctk.CTkEntry(master=entry_frame, placeholder_text="別途内部資料参考",
                                         width=450, height=50, border_width=2, corner_radius=8, text_font=self.general_font)
        control_num_entry.grid(row=6, column=1, padx=15, pady=15)

        note_label = ctk.CTkLabel(
            master=entry_frame, text="備考欄", text_font=self.general_font)
        note_label.grid(row=7, column=0, padx=15, pady=15)
        note_entry = ctk.CTkEntry(master=entry_frame, placeholder_text="", width=700,
                                  height=50, border_width=0, corner_radius=8, text_font=self.general_font)
        note_entry.grid(row=7, column=1, padx=15, pady=15)

        # 登録と終了ボタン用のframeを作成する
        button_frame = ctk.CTkFrame(
            master=frame, corner_radius=0, border_width=0)
        button_frame.pack(pady=5, anchor=tk.CENTER)

        register_button = ctk.CTkButton(master=button_frame, width=100, height=40, text_font=(
            "MS ゴシック", 14), border_width=0, corner_radius=8, text="登録", command=lambda: register())
        register_button.pack(padx=10, pady=10, side=tk.RIGHT)

        end_button = ctk.CTkButton(master=button_frame, width=200, height=40, border_width=0, corner_radius=8,
                                   text="メインメニューに戻る",  hover_color="red", text_font=("MS ゴシック", 14), command=lambda: frame.destroy())
        end_button.pack(padx=10, pady=10, side=tk.RIGHT)

        def register():
            """
            sqlite3とGoogleのスプレッドシートにそれぞれ登録する。また、登録する前に念のためネットワーク接続を確認している。
            """

            # 入力値の取得
            get_list = []  # 入力欄の値を格納するリスト
            for i in goods_entry.get(), goods_detail_entry.get(), use_group_box.get(), name_entry.get(), purchase_date_entry.get(), control_num_entry.get(), note_entry.get():
                # 空欄があると登録できないようにしている。問題がなければリストに追加
                if i == "":
                    messagebox.showwarning("通知", "空欄の所があります")
                    return None
                elif str.isdecimal(purchase_date_entry.get()) == False:
                    messagebox.showerror(
                        "入力値エラー", "購入日の入力値が正しく入力されていません\nex)数字以外が入力されている")
                    return None
                else:
                    get_list.append(i)

            try:
                # DBの登録前にインターネットに接続出来ているかを確認
                ping = backend.Ping()
                result = ping.ping()
                if result == True:
                    # sqlite3に追加
                    ope.db_register(get_list)
                    a.sp_insert(get_list)
                    # Googleスプレッドシート側に追加
                    ope = backend.DB_operation()
                    a = gsheet.Google_spreadsheet_operation()
                    messagebox.showinfo("通知", "登録が完了しました")
                elif result == False:
                    messagebox.showerror("警告", "インターネットに接続していないため、登録ができません")
            except sqlite3.IntegrityError:
                messagebox.showerror("通知", f"管理番号が重複しています")
            except Exception as e:
                messagebox.showerror("通知", f"エラーが発生しました。\n{e}")

            return None

    def delete_screen(self, root: ctk.CTk):
        #タイトルの変更
        root.title("高額物品管理システム(削除・廃棄)")

        #削除画面用のフレームの用意(base)
        frame = ctk.CTkFrame(master=root, width=self.width,
                             height=self.height, corner_radius=0, border_width=0)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.tkraise()     #作成したフレームを一番上に持ってくる。


        # Entry用のフレーム用意
        entry_frame = ctk.CTkFrame(
            master=frame, width=1200, height=800, corner_radius=0, border_width=0)
        entry_frame.pack(pady=50, anchor=tk.CENTER)

        control_num_label = ctk.CTkLabel(
            master=entry_frame, text="管理番号を入力してください", text_font=self.general_font)
        control_num_label.grid(row=1, column=0, padx=15, pady=15)
        control_num_entry = ctk.CTkEntry(master=entry_frame, placeholder_text="", width=400,
                                         height=50, border_width=2, corner_radius=8, text_font=self.general_font)
        control_num_entry.grid(row=2, column=3, padx=15, pady=15)

        # 登録と終了ボタン用のframeを作成する
        button_frame = ctk.CTkFrame(
            master=frame, corner_radius=0, border_width=0)
        button_frame.pack(pady=5, anchor=tk.CENTER)

        delete_button = ctk.CTkButton(master=button_frame, width=100, height=40, text_font=(
            "MS ゴシック", 14), border_width=0, corner_radius=8, text="完全に削除", fg_color="#ff758c", command=lambda: delete(control_num_entry.get()))
        delete_button.pack(padx=10, pady=10, side=tk.RIGHT)

        waste_button = ctk.CTkButton(master=button_frame, width=100, height=40, text_font=(
            "MS ゴシック", 14), border_width=0, corner_radius=8, text="廃棄", fg_color="#ff7eb3", command=lambda: waste(control_num_entry.get()))
        waste_button.pack(padx=10, pady=10, side=tk.RIGHT)

        end_button = ctk.CTkButton(master=button_frame, width=200, height=40, border_width=0, corner_radius=8,
                                   text="メインメニューに戻る",  fg_color="red", text_font=("MS ゴシック", 14), command=lambda: frame.destroy())
        end_button.pack(padx=10, pady=10, side=tk.RIGHT)

        def delete(control_num: str):
            """
            DBから削除を行う。また、同時にGoogleのスプレッドシートからも削除する。削除前に、インターネットに接続しているかも確認する。
            """
            if control_num_entry.get() == "":
                messagebox.showwarning("通知", "何も入力されていません")
            else:
                confirm = messagebox.askokcancel(
                    "確認", "データベースから完全に削除をしてもよろしいですか?\n※登録を間違えたとき以外は使用しないでください。")
                if confirm == True:
                    try:
                        # DBの削除前にインターネットに接続出来ているかを確認
                        ping = backend.Ping()
                        result = ping.ping()
                        if result == True:
                            ope = backend.DB_operation()
                            ope.db_delete(control_num)
                            messagebox.showinfo("通知", "削除が完了しました")
                        elif result == False:
                            messagebox.showerror(
                                "警告", "インターネットに接続していないため、登録ができません")
                    except Exception as e:
                        messagebox.showerror("エラー", f"エラーが発生しました。\n{e}")
            return None

        def waste(control_num: str):
            """
            廃棄するときに廃棄ということを登録すための関数。
            """
            if control_num_entry.get() == "":
                messagebox.showwarning("通知", "何も入力されていません")
                return None
            else:
                confirm = messagebox.askokcancel(
                    "確認", "データベースに対して廃棄済みとして登録しますがよろしいですか?")
                if confirm == True:
                    try:
                        # DBの登録前にインターネットに接続出来ているかを確認
                        ping = backend.Ping()
                        result = ping.ping()
                        if result == True:
                            ope = backend.DB_operation()
                            ope.db_waste(control_num)
                            messagebox.showinfo("通知", "登録が完了しました")
                        elif result == False:
                            messagebox.showerror(
                                "警告", "インターネットに接続していないため、登録ができません")
                    except Exception as e:
                        messagebox.showerror("エラー", f"エラーが発生しました。\n{e}")
            return None

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
        list_frame.pack(pady=10)

        column_all = ("物品名", "購入日", "管理番号", "使用団体", "代表者名", "備考", "廃棄")
        list_tree = ttk.Treeview(
            list_frame, columns=column_all, selectmode="none", height=45)
        style = ttk.Style()
        style.configure("Treeview.Heading", font=("MS ゴシック", 24))
        style.configure("Treeview", font=("Arial", 16))

        list_tree.heading('#0', text='')
        list_tree.heading('物品名', text='物品名', anchor='center')
        list_tree.heading('購入日', text='購入日', anchor='center')
        list_tree.heading('管理番号', text='管理番号', anchor='center')
        list_tree.heading('使用団体', text='使用団体', anchor='center')
        list_tree.heading('代表者名', text='代表者名', anchor='center')
        list_tree.heading('備考', text='備考', anchor='center')
        list_tree.heading('廃棄', text='廃棄', anchor='center')

        list_tree.column('#0', width=0, stretch=False)
        list_tree.column('物品名', anchor='center', width=250, stretch=False)
        list_tree.column('購入日', anchor='center', width=200, stretch=False)
        list_tree.column('管理番号', anchor='w', width=200, stretch=False)
        list_tree.column('使用団体', anchor='center', width=250, stretch=False)
        list_tree.column('代表者名', anchor='center', width=250, stretch=False)
        list_tree.column('備考', anchor='w', width=600, stretch=False)
        list_tree.column('廃棄', anchor='w', width=85, stretch=False)
        list_tree.pack()

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
                                   text="メインメニューに戻る",  fg_color="red", text_font=("MS ゴシック", 14), command=lambda: return_button())
        end_button.pack(padx=10, pady=10, side=tk.RIGHT)

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
