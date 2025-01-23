import tkinter as tk
from tkinter import messagebox,ttk,filedialog
import os
import sqlite3
import datetime
import subprocess
import json


# グローバル変数
database_name = ""
target_words = []

#--------------------------------------------------------------------------------------------------------------
# メッセージ欄にメッセージを表示
def update_message_window(message, tag="info"):
    message_window.tag_configure("info", foreground="blue")
    message_window.tag_configure("error", foreground="red")
    message_window.insert(tk.END, f"{message}\n", tag)
    message_window.see(tk.END)
#--------------------------------------------------------------------------------------------------------------
# メッセージ欄の初期化
def clear_message_window():
    message_window.delete(0., tk.END)
#--------------------------------------------------------------------------------------------------------------
# データベースを選択ボタン
def select_database():
    global database_name
    
    run_ocr_button.config(state='disabled')
    add_word_button.config(state='disabled')
    delete_word_button.config(state='disabled')
    if word_entry.get() !=  "":
        word_entry.config(state='normal')
        word_entry.delete(0, tk.END)
    word_entry.config(state='disabled')
    clear_button.config(state='disabled')
    
    f_type = [("データベース", "*.db")]
    i_file = os.path.abspath(os.path.dirname(__file__))
    #selected_file = filedialog.askopenfilename(filetype = f_type)
    selected_file = filedialog.askopenfilename(filetype = f_type, initialdir = i_file)
    database_name = os.path.basename(selected_file)
    if not selected_file:
        update_message_window("データベースファイルが選択されませんでした。", "error")
        return
    load_table_names()
#--------------------------------------------------------------------------------------------------------------
# データベース内のテーブル名を読み込み反映
def load_table_names():
    global database_name
    try:

        #ツリー内のデータを非表示
        for item in record_tree.get_children():
            record_tree.delete(item)

        for item in table_tree.get_children():
            table_tree.delete(item)

        conn = sqlite3.connect(database_name)
        cur = conn.cursor()
        # テーブル名を取得
        query= "SELECT name FROM sqlite_master WHERE type = 'table'"								
        cur.execute(query)
        tables = cur.fetchall()
            
        #ツリービューにテーブル名を反映
        for table in tables:
            table_tree.insert("", "end", values=table)

        database_label['text'] = database_name
        update_message_window(f"データベース {database_name} に接続しました。", "info")
    except Exception as e:
        update_message_window(f"データベースの読み込み中にエラーが発生しました: {e}", "error")

    finally:
        conn.close() 
#--------------------------------------------------------------------------------------------------------------
# 選択したテーブルのデータを読み込む
def load_table_data(event):
    global target_words
    selected_table = table_tree.focus()
    if selected_table:
        table_name = table_tree.item(selected_table,'value')[0]
        try:
            # ツリー内のデータを非表示
            for item in record_tree.get_children():
                record_tree.delete(item)

            conn = sqlite3.connect(database_name)
            cur = conn.cursor()                
            # データを取得
            query = f'SELECT * FROM "{table_name}"'
            cur.execute(query)
            rows = cur.fetchall()

            target_words = []
            # ツリービューにデータを反映
            for row in rows:
                record_tree.insert("", "end", values=row)
                target_words.append(row[0])
    
            run_ocr_button.config(state='normal')
            word_entry.config(state='normal')
            add_word_button.config(state='disabled')
            delete_word_button.config(state='disabled')
            
        except Exception as e:
            update_message_window(f"テーブルデータの読み込み中にエラーが発生しました: {e}", "error")
        finally:
            conn.close() 

#--------------------------------------------------------------------------------------------------------------
# 削除ボタンの有効化
def select_record(event):
    selected_record = record_tree.focus()
    if selected_record:  # 選択されたアイテムが存在する場合
        delete_word_button.config(state='normal')
    else:
        delete_word_button.config(state='disabled')
#--------------------------------------------------------------------------------------------------------------
# 削除ボタン
def delete_selected_word():
    selected_table = table_tree.focus()
    table_name = table_tree.item(selected_table, 'value')[0]
    selected_record = record_tree.focus()
    word_name = record_tree.item(selected_record, 'value')[0]
    try:
        conn = sqlite3.connect(database_name)
        cur = conn.cursor()
        # テーブルのカラム情報を取得
        query = f'PRAGMA table_info("{table_name}")'
        cur.execute(query)
        table_info = cur.fetchall()
        column_name = table_info[0][1]

        # 指定のデータを削除
        query = f'DELETE FROM "{table_name}" WHERE "{column_name}" = ?'
        cur.execute(query, (word_name,))
        conn.commit()

        update_message_window(f"単語 '{word_name}' を削除しました。", "info")
        load_table_data(None)

    except Exception as e:
        update_message_window(f"単語削除中にエラーが発生しました: {e}", "error")

        
    # 削除ボタンを無効化
    if not record_tree.get_children():
        delete_word_button.config(state='disabled')
    else:
        delete_word_button.config(state='normal') 
#--------------------------------------------------------------------------------------------------------------
# 追加ボタンとクリアボタンの有効、無効化
def enter_word(event):
    if word_entry.get() != "":
        add_word_button.config(state='normal')
        clear_button.config(state='normal')
    else:
        add_word_button.config(state='disabled')
        clear_button.config(state='disabled')
#--------------------------------------------------------------------------------------------------------------
# 追加ボタン
def add_word():
    selected_table = table_tree.focus()
    table_name = table_tree.item(selected_table, 'value')[0]
    new_word = word_entry.get().strip()
    date = format(datetime.date.today(), '%Y/%m/%d')
    try:
        # ツリー内のデータを非表示
        for item in record_tree.get_children():
            record_tree.delete(item)
        
        conn = sqlite3.connect(database_name)
        cur = conn.cursor()
        # 入力したデータを追加
        query = f'INSERT INTO "{table_name}" VALUES(?,?)'
        cur.execute(query, (new_word,date))
        conn.commit()

        update_message_window(f"単語 '{new_word}' を追加しました。", "info")
        load_table_data(None)

    except Exception as e:
        update_message_window(f"単語追加中にエラーが発生しました: {e}", "error")

    finally:
        clear()
#--------------------------------------------------------------------------------------------------------------
# クリアボタン
def clear():
    if word_entry.get() != "":
        word_entry.delete(0, tk.END)
        add_word_button.config(state='disabled')
        clear_button.config(state='disabled')
#--------------------------------------------------------------------------------------------------------------        
#検索実行ボタン
def search(target_words):
    update_message_window( f"OCR処理を開始します。対象文字列リスト: {target_words}", "info")
    
    try:
        # リストをJSON形式の文字列に変換
        target_strings = json.dumps(target_words)
        
        # OCRスクリプトを実行
        subprocess.run(["python", "ocr.py", target_strings],check=True)
        update_message_window("OCR処理が正常に終了しました。", "info")
    except subprocess.CalledProcessError as e:
        update_message_window(f"OCR処理中にエラーが発生しました: {e}", "error")
#-------------------------------------------------------------------------------------------------------------- 
#終了ボタン
def close_app():
    if messagebox.askokcancel("終了確認","アプリケーションを終了しますか?"):
        root.destroy()
#--------------------------------------------------------------------------------------------------------------

## ボタン仮押し用
#def submit():
#    print("ボタンが押されました")
#--------------------------------------------------------------------------------------------------------------


root = tk.Tk()
root.title("単語検索")



# フレームの作成
frame = tk.Frame(root)
frame.grid(row=0, column=0, padx=10, pady=10,sticky="nswe")

# データベース名ラベル
database_label = tk.Label(frame, text="", width=75, anchor='w', background='white')
database_label.grid(row=0, column=0, columnspan = 2, padx=10, pady=5, sticky='ew')

# タブ一覧ラベル
tables_label = tk.Label(frame, text="タブ一覧", width=20, anchor='e')
tables_label.grid(row=1, column=0, padx=10, pady=5, sticky='w')



# 新規追加用フレーム
tree_frame = tk.Frame(frame, width=800, height=460)  # 固定サイズのフレーム
tree_frame.grid(row=2, column=0, columnspan=2, padx=20, pady=5, sticky='nsew')

# タブ用のTreeview
table_columns = ("Table_num",)
table_tree = ttk.Treeview(tree_frame, show="headings", columns=table_columns, height=15)
table_tree.heading("Table_num", text="タブ名", anchor='w')
table_tree.column("Table_num", width=500, stretch=False)  # カラム幅をTreeviewの幅より大きく設定
table_tree.bind("<<TreeviewSelect>>", load_table_data)
table_tree.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")

# 縦スクロールバー
table_v_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=table_tree.yview)
table_tree.configure(yscroll=table_v_scrollbar.set)
table_v_scrollbar.grid(row=0, column=1, sticky=tk.W+"ns")

# 横スクロールバー
table_h_scrollbar = ttk.Scrollbar(tree_frame, orient="horizontal", command=table_tree.xview)
table_tree.configure(xscroll=table_h_scrollbar.set)
table_h_scrollbar.grid(row=1, column=0, sticky="ew")


# 単語一覧ラベル
records_label = tk.Label(frame, text="単語一覧", anchor='w')
records_label.grid(row=1, column=1, padx=10, pady=5, sticky="w")

# 単語用のTreeview
record_columns = ("Word_num", "Word_ymd")
record_tree = ttk.Treeview(tree_frame, show="headings", columns=record_columns, height=15)
record_tree.heading("Word_num", text="単語名")
record_tree.column("Word_num", width=450,stretch=False)
record_tree.heading("Word_ymd", text="登録日")
record_tree.column("Word_ymd", width=75,stretch=False)
record_tree.bind("<<TreeviewSelect>>", select_record)
record_tree.grid(row=0, column=1, padx=(50,5), pady=5, sticky="nsew")

# 縦スクロールバー
record_v_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=record_tree.yview)
record_tree.configure(yscroll=record_v_scrollbar.set)
record_v_scrollbar.grid(row=0, column=2, sticky=tk.W+"ns")

# 横スクロールバー
record_h_scrollbar = ttk.Scrollbar(tree_frame, orient="horizontal", command=record_tree.xview)
record_tree.configure(xscroll=record_h_scrollbar.set)
record_h_scrollbar.grid(row=1, column=1, padx=(50,5), sticky="ew")

# フレーム内のレイアウト調整
tree_frame.grid_propagate(False)  # フレームのサイズ自動調整を無効化
tree_frame.grid_rowconfigure(0, weight=1)
tree_frame.grid_columnconfigure(0, weight=1)



# 新規追加、削除用フレーム
add_delete_frame = tk.Frame(frame)
add_delete_frame.grid(row=2, column=2, rowspan=5, columnspan=2, padx=20, pady=5, sticky='n')

# 新規追加用ラベル
add_delete_label = tk.Label(add_delete_frame, text="新規追加", anchor='w')
add_delete_label.grid(row=0, column=0, padx=10, pady=5, sticky='w')

# タブ名エントリ
#table_label = tk.Label(add_delete_frame, text="タブ名", anchor='w')
#table_label.grid(row=1, column=0, padx=10, pady=5, sticky='w')
# 入力欄
#table_entry = tk.Entry(add_delete_frame, width=50, state=tk.DISABLED)
#table_entry.grid(row=1, column=1, padx=10, pady=5, sticky='ew')

# 単語名エントリ
date_label = tk.Label(add_delete_frame, text="単語名", anchor='w')
date_label.grid(row=2, column=0, padx=10, pady=5, sticky='w')
# 入力欄
word_entry = tk.Entry(add_delete_frame, width=50, state=tk.DISABLED)
word_entry.bind("<KeyRelease>", enter_word)
word_entry.grid(row=2, column=1, padx=10, pady=5, sticky='ew')

# 追加ボタン
add_word_button = tk.Button(add_delete_frame, text="追加", width=8, command=add_word, state=tk.DISABLED)
add_word_button.grid(row=3, column=1, padx=(50,10), pady=(50,5), sticky=tk.W)
# 削除ボタン
delete_word_button = tk.Button(add_delete_frame, text="削除", width=8, command=delete_selected_word, state=tk.DISABLED)
delete_word_button.grid(row=3, column=1, padx=(10,70), pady=(50,5), sticky=tk.E)



# メッセージウィンドウ
message_window = tk.Text(frame, height=5, width=50, bg='white')
message_window.insert("1.0","文字列検索アプリケーションを起動しました\r\n")
message_window.grid(row=4, column=0, columnspan=4, padx=10, pady=10, sticky='ew')



# ボタン用フレーム
button_frame = tk.Frame(frame)
button_frame.grid(row=5, column=0, columnspan=4, pady=10, sticky='e')

# データベース参照ボタン
select_db_button = tk.Button(button_frame, text="データベースを選択", width=15, command=select_database)
select_db_button.pack(side='left', padx=5)

# 検索実行ボタン
run_ocr_button = tk.Button(button_frame, text="検索実行", width=10, command=lambda: search(target_words), state=tk.DISABLED)
run_ocr_button.pack(side='left', padx=5)

# クリアボタン
clear_button = tk.Button(button_frame, text="入力クリア", width=8, command=clear, state=tk.DISABLED)
clear_button.pack(side='left', padx=5)

# 終了ボタン
exit_button = tk.Button(button_frame, text="終了", width=8, bg='yellow', command=close_app)
exit_button.pack(side='left', padx=5)


root.mainloop()