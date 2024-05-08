import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
from tkinterdnd2 import DND_FILES, TkinterDnD
from datetime import datetime, timedelta
import csv


class WorkTimeApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("勤務時間丸めツール")
        
        # ウィンドウサイズを設定し、最小サイズを定義
        self.geometry("800x600")
        self.minsize(800, 600)

        self.file_path = None
        self.round_minutes = tk.IntVar(value=15)

        self.upload_frame = tk.Frame(self)
        self.upload_frame.pack(fill=tk.BOTH, expand=True)

        self.settings_frame = tk.Frame(self)
        self.result_frame = tk.Frame(self)

        self.create_upload_ui()

        # 出勤時間と退勤時間の丸め選択用変数
        self.start_round_option = tk.StringVar(value="そのまま")
        self.end_round_option = tk.StringVar(value="そのまま")
        self.break_start_round_option = tk.StringVar(value="そのまま") 
        self.break_end_round_option = tk.StringVar(value="そのまま")

    def on_drop(self, event):
        file_path = event.data.strip('{}')
        if file_path:
            self.file_path = file_path
            df = self.read_file(self.file_path)
            self.upload_frame.pack_forget()
            self.create_settings_ui(df)
            self.settings_frame.pack(fill=tk.BOTH, expand=True)
    


    def create_upload_ui(self):
        for widget in self.upload_frame.winfo_children():
            widget.destroy()


        # ドラッグ＆ドロップエリアを明確にする
        drop_frame = tk.Frame(self.upload_frame, bd=2, relief="groove", bg="#e1e4e8")
        drop_frame.pack(padx=20, pady=20, fill=tk.BOTH, expand=True)

        drop_label = tk.Label(drop_frame, text="ファイルをここにドラッグ＆ドロップ\nまたは", pady=20, padx=20,bg="#e1e4e8")
        drop_label.pack(expand=True)


        open_button = tk.Button(drop_frame, text="ファイルを開く", command=self.open_file)
        open_button.pack(pady=10)

        drop_frame.drop_target_register(DND_FILES)
        drop_frame.dnd_bind('<<Drop>>', self.on_drop)

    def create_settings_ui(self, data):
        for widget in self.settings_frame.winfo_children():
            widget.destroy()

        # スクロール可能なテキストエリアを設定
        text_area = scrolledtext.ScrolledText(self.settings_frame, wrap=tk.WORD, height=10)
        text_area.grid(row=0, column=0, columnspan=2, sticky='ew', padx=5, pady=5)

        text_area.insert(tk.INSERT, "日付\t出勤\t休始\t休終\t退勤\t実働\n")
        for row in data:
            text_area.insert(tk.INSERT, f"{row[0]}\t{row[1]}\t{row[2]}\t{row[3]}\t{row[6]\t{row[15]}\n")
        text_area.configure(state='disabled')

        # 出勤時間の丸め選択
        start_round_label = tk.Label(self.settings_frame, text="出勤時間の丸め選択：")
        start_round_label.grid(row=3, column=0, pady=5, sticky='e')

        start_round_menu = ttk.Combobox(self.settings_frame, textvariable=self.start_round_option, values=["8:30", "9:00", "そのまま"], width=15)
        start_round_menu.grid(row=3, column=1, pady=5, sticky='w')

        # 退勤時間の丸め選択
        end_round_label = tk.Label(self.settings_frame, text="退勤時間の丸め選択：")
        end_round_label.grid(row=4, column=0, pady=5, sticky='e')

        end_round_menu = ttk.Combobox(self.settings_frame, textvariable=self.end_round_option, values=["12:40", "12:00 & 18:30", "そのまま"], width=15)
        end_round_menu.grid(row=4, column=1, pady=5, sticky='w')

        # 休憩開始時刻の丸め選択
        break_start_round_label = tk.Label(self.settings_frame, text="休憩開始時刻の丸め選択：")
        break_start_round_label.grid(row=5, column=0, pady=5, sticky='e')

        break_start_round_menu = ttk.Combobox(self.settings_frame, textvariable=self.break_start_round_option, values=["12:00","12:40", "そのまま"], width=15)
        break_start_round_menu.grid(row=5, column=1, pady=5, sticky='w')

        # 休憩終了時刻の丸め選択
        break_end_round_label = tk.Label(self.settings_frame, text="休憩終了時刻の丸め選択：")
        break_end_round_label.grid(row=6, column=0, pady=5, sticky='e')

        break_end_round_menu = ttk.Combobox(self.settings_frame, textvariable=self.break_end_round_option, values=["14:25","15:00", "そのまま"], width=15)
        break_end_round_menu.grid(row=6, column=1, pady=5, sticky='w')

        execute_button = tk.Button(self.settings_frame, text="計算を実行", command=self.execute_calculation)
        execute_button.grid(row=2, column=0, columnspan=2, pady=20)        

        # グリッドの配置を調整
        self.settings_frame.columnconfigure(0, weight=1)
        self.settings_frame.columnconfigure(1, weight=1)
        self.settings_frame.rowconfigure(0, weight=1)


    def create_result_ui(self, hours, minutes, df):
        for widget in self.result_frame.winfo_children():
            widget.destroy()  # 既存のウィジェットを削除

        result_label = tk.Label(self.result_frame, text=f"総勤務時間: {hours}時間{minutes}分")
        result_label.pack(pady=20)

        # 丸めた時刻を表示するテキストエリアを追加
        text_area = scrolledtext.ScrolledText(self.result_frame, wrap=tk.WORD, height=10)
        text_area.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        text_area.insert(tk.INSERT, "日付\t出勤\t休始\t休終\t退勤\t実働\n")
        for row in df:
            text_area.insert(tk.INSERT, f"{row[0]}\t{row[1]}\t{row[2]}\t{row[3]}\t{row[6]}\t{row[15]}\n")
        text_area.configure(state='disabled')

        ok_button = tk.Button(self.result_frame, text="OK", command=self.reset_app)
        ok_button.pack(pady=10)
    
    def open_file(self):
        self.file_path = filedialog.askopenfilename()
        if self.file_path:
            df = self.read_file(self.file_path)
            self.upload_frame.pack_forget()
            self.create_settings_ui(df)
            self.settings_frame.pack(fill=tk.BOTH, expand=True)

    def read_file(self, file_path):
        data = []
        with open(file_path, 'r', encoding='utf-8') as file:
            reader = csv.reader(file)
            next(reader)  # ヘッダー行をスキップ
            for row in reader:
                if row[1] != '':  # 出勤時間が記録されている行のみを追加
                    data.append(row)
        return data

    def execute_calculation(self):
        hours, minutes, data = self.process_file(self.file_path)  # 丸めた時刻を含むデータフレームを受け取る
        self.settings_frame.pack_forget()
        self.create_result_ui(hours, minutes, data)  # 丸めた時刻を含むデータフレームを渡す
        self.result_frame.pack(fill=tk.BOTH, expand=True)

    def reset_app(self):
        self.result_frame.pack_forget()
        self.create_upload_ui()
        self.upload_frame.pack(fill=tk.BOTH, expand=True)

    def round_time(self, time_str, option, type):
        if time_str == '' or time_str in ['−', '-'] or option == "そのまま":
            return time_str

        time_obj = datetime.strptime(time_str, '%H:%M')
        if type == "start":
            if option == "8:30" and time_obj < datetime.strptime("9:00", '%H:%M'):
                return "8:30" if (datetime.strptime("8:30", '%H:%M') - time_obj).total_seconds() / 60 <= 20 else time_str
            elif option == "9:00":
                return "9:00" if (datetime.strptime("9:00", '%H:%M') - time_obj).total_seconds() / 60 <= 20 else time_str
        elif type == "end":
            if option == "12:40" and time_obj > datetime.strptime("12:00", '%H:%M'):
                return "12:40" if (time_obj - datetime.strptime("12:40", '%H:%M')).total_seconds() / 60 <= 20 else time_str
            elif option == "12:00&18:30":
                if time_obj > datetime.strptime("11:30", '%H:%M'):
                    return "12:00" if (datetime.strptime("12:00", '%H:%M') - time_obj).total_seconds() / 60 <= 20 else time_str
                elif time_obj > datetime.strptime("18:00", '%H:%M'):
                    return "18:30" if (datetime.strptime("18:30", '%H:%M') - time_obj).total_seconds() / 60 <= 20 else time_str

        return time_str

    def process_file(self, file_path):
        data = self.read_file(file_path)
    
        for row in data:
            # 出勤時間の丸め処理
            row[1] = self.round_time(row[1], self.start_round_option.get(), "start")
        
            # 退勤時間の丸め処理
            row[6] = self.round_time(row[6], self.end_round_option.get(), "end")
        
            # 休憩時間の丸め処理
            row[3] = self.round_break_end_time(row[3], self.break_end_round_option.get())
        
            # 休憩開始時刻の丸め処理
            row[2] = self.round_break_start_time(row[2], self.break_start_round_option.get())
        
            # 退勤時刻が記録されていない場合、休憩開始時刻を勤務終了時刻として扱う処理
            if ":" not in row[6] and ":" in row[2]:
                row[6] = row[2]
    
        total_work_seconds = 0
        for row in data:
            start_time = datetime.strptime(row[1], '%H:%M')
            end_time = datetime.strptime(row[6], '%H:%M')
            work_time = end_time - start_time
            break_time = calculate_break_time(row[2], row[3])
            total_work_seconds += (work_time - break_time).total_seconds()
    
        total_hours = total_work_seconds // 3600
        total_minutes = (total_work_seconds % 3600) // 60
    
        return int(total_hours), int(total_minutes), data


    def round_break_end_time(self, time_str, option):
        if time_str == '' or time_str in ['−', '-'] or option == "そのまま":
            return time_str
        time_obj = datetime.strptime(time_str, '%H:%M')
        if option == "15:00":
            round_time_obj = datetime.strptime("15:00", '%H:%M')
        elif option == "14:25":
            round_time_obj = datetime.strptime("14:25", '%H:%M')
        
        if time_obj > round_time_obj and (time_obj - round_time_obj).total_seconds() / 60 <= 20:
            return option
        else:
            return time_str
    def round_break_start_time(self, time_str, option):
        if time_str == '' or time_str in ['−', '-'] or option == "そのまま":
            return time_str

        time_obj = datetime.strptime(time_str, '%H:%M')
        if option == "12:00" and datetime.strptime("11:30", '%H:%M') <= time_obj <= datetime.strptime("12:00", '%H:%M'):
            return "12:00"
        elif option == "12:40" and datetime.strptime("12:20", '%H:%M') <= time_obj <= datetime.strptime("12:40", '%H:%M'):
            return "12:40"

        return time_str

def calculate_break_time(start, end):
    if start == '' or end == '' or start in ['−', '-'] or end in ['−', '-']:
        return timedelta(0)
    start_time = datetime.strptime(start, '%H:%M')
    end_time = datetime.strptime(end, '%H:%M')
    return end_time - start_time




if __name__ == "__main__":
    app = WorkTimeApp()
    app.mainloop()
