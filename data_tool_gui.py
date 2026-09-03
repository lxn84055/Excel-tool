import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import os
import threading
from datetime import datetime
import re
import csv
import time
import numpy as np

class DataProcessingGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel/Word 数据处理工具")
        self.root.geometry("800x700")
        
        style = ttk.Style()
        style.theme_use('clam')
        
        self.main_canvas = tk.Canvas(root, highlightthickness=0)
        self.main_scrollbar = ttk.Scrollbar(root, orient="vertical", command=self.main_canvas.yview)
        self.main_canvas.configure(yscrollcommand=self.main_scrollbar.set)
        self.main_canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.main_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        self.main_frame = ttk.Frame(self.main_canvas, padding="10")
        self.canvas_window = self.main_canvas.create_window((0, 0), window=self.main_frame, anchor="nw")
        self.main_frame.bind("<Configure>", self.on_frame_configure)
        self.main_canvas.bind("<Configure>", self.on_canvas_configure)
        self.bind_mousewheel_recursive(root)
        
        self.notebook = ttk.Notebook(self.main_frame)
        self.create_merge_tab()
        self.create_convert_tab()
        self.create_split_tab()
        self.create_batch_tab()
        self.create_progress_area()
        self.create_log_area()
        self.create_help_button()
        
        self.notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.help_button_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=3)
        self.progress_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=3)
        self.log_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=3)
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.rowconfigure(0, weight=1)
        
        self.is_processing = False
        self.last_progress_time = 0
        self.root.minsize(700, 600)
        self.center_window()
        
        self.interactive_widgets = []
        self.collect_interactive_widgets()
    
    def collect_interactive_widgets(self):
        for widget in self.get_all_widgets(self.main_frame):
            if isinstance(widget, (ttk.Button, ttk.Entry, ttk.Checkbutton, 
                                  ttk.Radiobutton, ttk.Combobox, tk.Listbox,
                                  ttk.Scale, ttk.Spinbox)):
                if widget != self.cancel_button:
                    self.interactive_widgets.append(widget)
    
    def lock_interface(self):
        self.is_processing = True
        for tab_id in self.notebook.tabs():
            self.notebook.tab(tab_id, state='disabled')
        for widget in self.interactive_widgets:
            try: widget.config(state='disabled')
            except: pass
        try: self.main_scrollbar.grid_remove()
        except: pass
        self.unbind_mousewheel_recursive(self.root)
        self.cancel_button.config(state='normal')
        self.notebook.grid_remove()
        self.help_button_frame.grid_remove()
        self.progress_frame.grid_remove()
        self.log_frame.grid_remove()
        self.progress_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5)
        self.log_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        self.notebook.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.update_idletasks()
    
    def unlock_interface(self):
        self.is_processing = False
        for tab_id in self.notebook.tabs():
            self.notebook.tab(tab_id, state='normal')
        for widget in self.interactive_widgets:
            try: widget.config(state='normal')
            except: pass
        try: self.main_scrollbar.grid()
        except: pass
        self.bind_mousewheel_recursive(self.root)
        self.cancel_button.config(state='disabled')
        self.notebook.grid_remove()
        self.help_button_frame.grid_remove()
        self.progress_frame.grid_remove()
        self.log_frame.grid_remove()
        self.notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.help_button_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=3)
        self.progress_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=3)
        self.log_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=3)
        self.root.update_idletasks()
    
    def unbind_mousewheel_recursive(self, widget):
        try:
            widget.unbind("<MouseWheel>")
            widget.unbind("<Button-4>")
            widget.unbind("<Button-5>")
        except: pass
        for child in widget.winfo_children():
            self.unbind_mousewheel_recursive(child)
    
    def get_all_widgets(self, parent):
        widgets = []
        try:
            for child in parent.winfo_children():
                widgets.append(child)
                widgets.extend(self.get_all_widgets(child))
        except: pass
        return widgets
    
    def detect_header(self, df):
        if df.empty: return False
        column_names = df.columns.tolist()
        header_score = 0
        data_score = 0
        for col in column_names:
            col_str = str(col)
            if re.match(r'^[A-Za-z\u4e00-\u9fff][A-Za-z0-9\u4e00-\u9fff_\s]*$', col_str):
                header_score += 1
            if re.match(r'^\d+$', col_str): data_score += 1
            elif 'Unnamed' in col_str: data_score += 1
        return header_score > data_score
    
    def detect_header_quick(self, file_path):
        try:
            ext = os.path.splitext(file_path)[1].lower()
            if ext == '.csv': return self.detect_header_csv_quick(file_path)
            elif ext in ['.xlsx', '.xls']: return self.detect_header_excel_quick(file_path)
            else: return True
        except: return True
    
    def detect_header_csv_quick(self, file_path):
        try:
            encoding = self.detect_file_encoding(file_path)
            with open(file_path, 'r', encoding=encoding, errors='ignore') as f:
                lines = []
                for i, line in enumerate(f):
                    if i >= 3: break
                    lines.append(line.strip())
            if not lines: return False
            delimiter = self.detect_delimiter(lines[0])
            first_row = lines[0].split(delimiter)
            second_row = lines[1].split(delimiter) if len(lines)>1 else None
            header_score = 0; data_score = 0
            for cell in first_row:
                cell = cell.strip().strip('"').strip("'")
                if re.match(r'^[A-Za-z\u4e00-\u9fff][A-Za-z0-9\u4e00-\u9fff_\s]*$', cell): header_score += 1
                if re.match(r'^\d+$', cell): data_score += 1
            if second_row:
                for c1, c2 in zip(first_row, second_row):
                    c1 = c1.strip().strip('"').strip("'")
                    c2 = c2.strip().strip('"').strip("'")
                    if re.match(r'^[A-Za-z\u4e00-\u9fff]', c1) and re.match(r'^\d+$', c2): header_score += 2
            return header_score > data_score
        except: return True
    
    def detect_header_excel_quick(self, file_path):
        try:
            ext = os.path.splitext(file_path)[1].lower()
            if ext == '.xlsx': df = pd.read_excel(file_path, engine='openpyxl', nrows=3, dtype=object)
            else:
                try: df = pd.read_excel(file_path, engine='xlrd', nrows=3, dtype=object)
                except: df = pd.read_excel(file_path, engine='openpyxl', nrows=3, dtype=object)
            if df.empty: return False
            column_names = df.columns.tolist()
            header_score = 0; data_score = 0
            for col in column_names:
                col_str = str(col)
                if re.match(r'^[A-Za-z\u4e00-\u9fff][A-Za-z0-9\u4e00-\u9fff_\s]*$', col_str): header_score += 1
                if re.match(r'^\d+$', col_str): data_score += 1
                elif 'Unnamed' in col_str: data_score += 1
            return header_score > data_score
        except: return True
    
    def get_columns_quick(self, file_path):
        try:
            ext = os.path.splitext(file_path)[1].lower()
            has_header = self.detect_header_quick(file_path)
            if ext == '.csv': return self.get_csv_columns_quick(file_path, has_header)
            elif ext in ['.xlsx', '.xls']: return self.get_excel_columns_quick(file_path, has_header)
            else: return [], True
        except: return [], True
    
    def get_csv_columns_quick(self, file_path, has_header):
        try:
            encoding = self.detect_file_encoding(file_path)
            with open(file_path, 'r', encoding=encoding, errors='ignore') as f:
                lines = []
                for i, line in enumerate(f):
                    if i >= 5: break
                    lines.append(line.strip())
            if not lines: return [], has_header
            delimiter = self.detect_delimiter(lines[0])
            if has_header:
                columns = [c.strip().strip('"').strip("'") for c in lines[0].split(delimiter)]
            else:
                max_cols = max(len(line.split(delimiter)) for line in lines)
                columns = [f"列{i+1}" for i in range(max_cols)]
            return columns, has_header
        except: return [], has_header
    
    def get_excel_columns_quick(self, file_path, has_header):
        try:
            ext = os.path.splitext(file_path)[1].lower()
            if ext == '.xlsx': df = pd.read_excel(file_path, engine='openpyxl', nrows=5, dtype=object)
            else:
                try: df = pd.read_excel(file_path, engine='xlrd', nrows=5, dtype=object)
                except: df = pd.read_excel(file_path, engine='openpyxl', nrows=5, dtype=object)
            if has_header: columns = df.columns.tolist()
            else:
                max_cols = len(df.columns)
                columns = [f"列{i+1}" for i in range(max_cols)]
            return columns, has_header
        except: return [], has_header
    
    def normalize_columns(self, df, reference_columns=None):
        if reference_columns is None: reference_columns = df.columns.tolist()
        current_columns = df.columns.tolist()
        if len(current_columns) == len(reference_columns):
            if current_columns != reference_columns: df.columns = reference_columns
        else:
            matched = []
            for i in range(min(len(current_columns), len(reference_columns))):
                matched.append(reference_columns[i])
            while len(matched) < len(current_columns):
                matched.append(f"列{len(matched)+1}")
            df.columns = matched
        return df
    
    def read_csv_without_header(self, file_path):
        encoding = self.detect_file_encoding(file_path)
        with open(file_path, 'r', encoding=encoding, errors='ignore') as f:
            sample = f.read(4096)
            f.seek(0)
            delimiter = self.detect_delimiter(sample)
            csv_reader = csv.reader(f, delimiter=delimiter)
            rows = [row for row in csv_reader if row]
        if not rows: return pd.DataFrame()
        max_cols = max(len(row) for row in rows)
        default_columns = [f"列{i+1}" for i in range(max_cols)]
        processed_rows = []
        for row in rows:
            if len(row) < max_cols: row += ['']*(max_cols-len(row))
            elif len(row) > max_cols: row = row[:max_cols]
            processed_rows.append(row)
        return pd.DataFrame(processed_rows, columns=default_columns, dtype=object)
    
    def read_excel_without_header(self, file_path):
        ext = os.path.splitext(file_path)[1].lower()
        if ext == '.xlsx': df = pd.read_excel(file_path, engine='openpyxl', header=None, dtype=object)
        else:
            try: df = pd.read_excel(file_path, engine='xlrd', header=None, dtype=object)
            except: df = pd.read_excel(file_path, engine='openpyxl', header=None, dtype=object)
        default_columns = [f"列{i+1}" for i in range(len(df.columns))]
        df.columns = default_columns
        return df
    
    def remove_blank_data(self, df, remove_blank_rows=True, remove_blank_cols=True, remove_blank_cells=True):
        original_shape = df.shape
        if remove_blank_rows:
            df = df.dropna(how='all')
            if original_shape[0] - df.shape[0] > 0: self.log(f"去除空白行: {original_shape[0] - df.shape[0]}行")
        if remove_blank_cols:
            df = df.dropna(axis=1, how='all')
            if original_shape[1] - df.shape[1] > 0: self.log(f"去除空白列: {original_shape[1] - df.shape[1]}列")
        if remove_blank_cells:
            for col in df.columns:
                if df[col].dtype == 'object':
                    df[col] = df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
                    df[col] = df[col].apply(lambda x: np.nan if isinstance(x, str) and x=='' else x)
        return df
    
    def read_csv_file_robust(self, file_path):
        encoding = self.detect_file_encoding(file_path)
        rows = []
        max_columns = 0
        with open(file_path, 'r', encoding=encoding, errors='ignore') as f:
            sample = f.read(4096)
            f.seek(0)
            delimiter = self.detect_delimiter(sample)
            csv_reader = csv.reader(f, delimiter=delimiter)
            for row in csv_reader:
                if row:
                    rows.append(row)
                    max_columns = max(max_columns, len(row))
        if not rows: raise Exception("CSV文件为空")
        processed = []
        for row in rows:
            if len(row) < max_columns: row += ['']*(max_columns-len(row))
            elif len(row) > max_columns: row = row[:max_columns]
            processed.append(row)
        if len(processed) > 1:
            columns = processed[0]
            data = processed[1:]
            df = pd.DataFrame(data, columns=columns, dtype=object)
        else: df = pd.DataFrame(processed, dtype=object)
        return df
    
    def detect_file_encoding(self, file_path):
        encodings = ['utf-8', 'utf-8-sig', 'gbk', 'gb2312', 'gb18030', 'big5', 'latin1']
        for enc in encodings:
            try:
                with open(file_path, 'r', encoding=enc) as f: f.read(1024)
                return enc
            except: continue
        return 'latin1'
    
    def detect_delimiter(self, sample_text):
        delimiters = [',', ';', '\t', '|']
        counts = {d: sample_text.count(d) for d in delimiters}
        if counts:
            max_d = max(counts, key=counts.get)
            if counts[max_d] > 0: return max_d
        return ','
    
    def read_excel_file(self, file_path):
        ext = os.path.splitext(file_path)[1].lower()
        if ext == '.csv': return self.read_csv_file_robust(file_path)
        elif ext == '.xlsx': return pd.read_excel(file_path, engine='openpyxl', dtype=object)
        elif ext == '.xls':
            try: return pd.read_excel(file_path, engine='xlrd', dtype=object)
            except: return pd.read_excel(file_path, engine='openpyxl', dtype=object)
        else: return pd.read_excel(file_path, dtype=object)
    
    def save_excel_file(self, df, file_path):
        ext = os.path.splitext(file_path)[1].lower()
        if ext == '.csv': df.to_csv(file_path, index=False, encoding='utf-8-sig')
        elif ext == '.xlsx': df.to_excel(file_path, index=False, engine='openpyxl')
        elif ext == '.xls': df.to_excel(file_path, index=False, engine='xlwt')
        elif ext == '.txt': df.to_csv(file_path, index=False, sep='\t', encoding='utf-8-sig')
        else: df.to_excel(file_path, index=False, engine='openpyxl')
        return True
    
    def on_frame_configure(self, event): self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))
    def on_canvas_configure(self, event): self.main_canvas.itemconfig(self.canvas_window, width=event.width)
    
    def bind_mousewheel_recursive(self, widget):
        widget.bind("<MouseWheel>", self.on_mousewheel_windows)
        widget.bind("<Button-4>", self.on_mousewheel_linux)
        widget.bind("<Button-5>", self.on_mousewheel_linux)
        for child in widget.winfo_children(): self.bind_mousewheel_recursive(child)
    
    def on_mousewheel_windows(self, event):
        if self.is_processing: return "break"
        if event.delta > 0: self.main_canvas.yview_scroll(-3, "units")
        else: self.main_canvas.yview_scroll(3, "units")
        return "break"
    def on_mousewheel_linux(self, event):
        if self.is_processing: return "break"
        if event.num == 4: self.main_canvas.yview_scroll(-3, "units")
        elif event.num == 5: self.main_canvas.yview_scroll(3, "units")
        return "break"
    
    def center_window(self):
        self.root.update_idletasks()
        w, h = 800, 700
        x = (self.root.winfo_screenwidth()//2) - (w//2)
        y = (self.root.winfo_screenheight()//2) - (h//2)
        self.root.geometry(f'{w}x{h}+{x}+{y}')
    
    def create_progress_area(self):
        self.progress_frame = ttk.LabelFrame(self.main_frame, text="处理进度", padding="5")
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.progress_frame, variable=self.progress_var, maximum=100, length=750, mode='determinate')
        self.progress_bar.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        self.progress_label = ttk.Label(self.progress_frame, text="0%")
        self.progress_label.grid(row=0, column=2, padx=10)
        self.status_label = ttk.Label(self.progress_frame, text="就绪", foreground="green")
        self.status_label.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=5)
        self.indeterminate_progress = ttk.Progressbar(self.progress_frame, mode='indeterminate', length=750)
        self.indeterminate_progress.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        self.indeterminate_progress.grid_remove()
        self.cancel_button = ttk.Button(self.progress_frame, text="取消", command=self.cancel_operation, state='disabled')
        self.cancel_button.grid(row=2, column=2, padx=10)
        self.cancel_flag = False
        self.last_progress_time = 0
    
    def create_log_area(self):
        self.log_frame = ttk.LabelFrame(self.main_frame, text="操作日志", padding="5")
        self.log_text = scrolledtext.ScrolledText(self.log_frame, height=5, width=90)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E))
        ttk.Button(self.log_frame, text="清除日志", command=self.clear_log).grid(row=1, column=0, pady=2)
    
    def create_help_button(self):
        self.help_button_frame = ttk.Frame(self.main_frame)
        ttk.Button(self.help_button_frame, text="📖 功能说明", command=self.show_help).pack(side=tk.LEFT, padx=5)
        ttk.Button(self.help_button_frame, text="📐 公式说明", command=self.show_formula_help).pack(side=tk.LEFT, padx=5)
        ttk.Button(self.help_button_frame, text="🔍 功能搜索", command=self.show_search).pack(side=tk.LEFT, padx=5)
    
    def show_help(self):
        win = tk.Toplevel(self.root); win.title("功能使用说明"); win.geometry("700x600")
        text = scrolledtext.ScrolledText(win, width=80, height=35)
        text.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        content = self.get_help_content()
        text.insert(tk.END, content); text.config(state='disabled')
        ttk.Button(win, text="关闭", command=win.destroy).pack(pady=10)
    
    def get_help_content(self):
        return """
╔══════════════════════════════════════════════════════════════╗
║                    功能使用说明                                ║
╚══════════════════════════════════════════════════════════════╝

【一、数据合并】
功能：将多个Excel/CSV文件合并为一个文件
- 支持选择特定列（按列名或列号）
- 支持垂直/水平合并
- 支持去除空白行/列/字符
- 支持合并所有Sheet（Excel多工作表）
- 支持添加数据来源标识、去重

【二、格式转换】
- Excel/CSV → Word/PDF/PPT/TXT
- Word → Excel
- Word → PDF
- PPT → PDF

【三、数据拆分】
1. 按列值拆分：根据某列的不同值拆分为多个文件
2. 按列位置拆分：
   - 前后拆分：输入列号（可多个），按位置切成多段
   - 指定列提取：只提取指定列
3. 按行数拆分：每N行一个文件
4. 按特定行拆分：在指定行号处拆分

【四、批量处理】
- 清理数据（去重+去空白）
- 去除空行
- 公式转数值（保持列头不变）

【五、预览列名】
快速查看列名，支持点击添加到输入框。
"""
    
    def show_formula_help(self):
        win = tk.Toplevel(self.root); win.title("公式使用说明"); win.geometry("700x600")
        text = scrolledtext.ScrolledText(win, width=80, height=35)
        text.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        content = """
【常见Excel公式】
=SUM(A1:A10)       求和
=AVERAGE(A1:A10)   平均值
=MAX(A1:A10)       最大值
=MIN(A1:A10)       最小值
=COUNT(A1:A10)     计数
=IF(A1>50,"通过","不通过")  条件判断
=VLOOKUP(A1,B1:D100,2,FALSE) 查找
=CONCATENATE(A1,B1) 合并文本
=TODAY()           今天日期
=YEAR(A1)          取年份

【公式转数值】
将公式计算结果转为纯数值，列头保持不变。
"""
        text.insert(tk.END, content); text.config(state='disabled')
        ttk.Button(win, text="关闭", command=win.destroy).pack(pady=10)
    
    def show_search(self):
        win = tk.Toplevel(self.root); win.title("功能搜索"); win.geometry("600x500")
        search_frame = ttk.Frame(win); search_frame.pack(pady=10, padx=10, fill=tk.X)
        ttk.Label(search_frame, text="搜索:").pack(side=tk.LEFT, padx=5)
        search_entry = ttk.Entry(search_frame, width=40); search_entry.pack(side=tk.LEFT, padx=5)
        result_text = scrolledtext.ScrolledText(win, width=70, height=20)
        result_text.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        all_functions = [
            ("数据合并","将多个Excel/CSV文件合并为一个文件"),
            ("垂直合并","按行追加合并"),
            ("水平合并","按列拼接合并"),
            ("合并所有Sheet","合并Excel中的所有工作表"),
            ("格式转换","在不同格式间转换"),
            ("Excel转Word","转为Word表格"),
            ("Excel转PDF","转为PDF表格"),
            ("Excel转PPT","转为PPT幻灯片"),
            ("Excel转TXT","转为文本文件"),
            ("数据拆分","拆分为多个小文件"),
            ("按列值拆分","根据列值拆分"),
            ("按列位置拆分","按列位置前后拆分或提取"),
            ("按行数拆分","每N行一个文件"),
            ("按特定行拆分","在指定行处拆分"),
            ("批量处理","批量处理文件夹中的文件"),
            ("清理数据","去重+去空白"),
            ("去除空行","删除空行"),
            ("公式转数值","公式结果转为纯数值"),
        ]
        def do_search():
            keyword = search_entry.get().strip().lower()
            result_text.delete(1.0, tk.END)
            if not keyword:
                result_text.insert(tk.END, "请输入搜索关键词\n"); return
            found = False
            for name, desc in all_functions:
                if keyword in name.lower() or keyword in desc.lower():
                    result_text.insert(tk.END, f"▶ {name}\n  {desc}\n\n")
                    found = True
            if not found: result_text.insert(tk.END, f"未找到与 '{keyword}' 相关的功能\n")
        def show_all():
            result_text.delete(1.0, tk.END)
            result_text.insert(tk.END, "所有功能列表：\n\n")
            for name, desc in all_functions:
                result_text.insert(tk.END, f"▶ {name}\n  {desc}\n\n")
        ttk.Button(search_frame, text="搜索", command=do_search).pack(side=tk.LEFT, padx=5)
        ttk.Button(search_frame, text="显示全部", command=show_all).pack(side=tk.LEFT, padx=5)
        show_all()
        ttk.Button(win, text="关闭", command=win.destroy).pack(pady=10)
    
    def clear_log(self):
        if not self.is_processing: self.log_text.delete(1.0, tk.END)
    def log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    def update_progress(self, value, status_text=None, force_update=False):
        current = time.time()
        if not force_update and current - self.last_progress_time < 0.1: return
        self.last_progress_time = current
        self.progress_var.set(value)
        self.progress_label.config(text=f"{int(value)}%")
        if status_text: self.status_label.config(text=status_text)
        self.root.update_idletasks()
    def start_indeterminate(self, status_text="处理中..."):
        self.indeterminate_progress.grid(); self.indeterminate_progress.start(10)
        self.status_label.config(text=status_text, foreground="blue")
        self.cancel_button.config(state='normal'); self.cancel_flag = False
        self.lock_interface(); self.root.update_idletasks()
    def stop_indeterminate(self, status_text="完成", success=True):
        self.indeterminate_progress.stop(); self.indeterminate_progress.grid_remove()
        if success:
            self.status_label.config(text=status_text, foreground="green")
            self.progress_var.set(100); self.progress_label.config(text="100%")
        else: self.status_label.config(text=status_text, foreground="red")
        self.unlock_interface(); self.root.update_idletasks()
    def cancel_operation(self):
        self.cancel_flag = True
        self.status_label.config(text="正在取消...", foreground="orange")
        self.cancel_button.config(state='disabled')
        self.log("用户请求取消操作")
    def check_cancel(self):
        if self.cancel_flag: raise Exception("操作已被用户取消")
    
    # ======================== 标签页创建 ========================
    def create_merge_tab(self):
        merge_frame = ttk.Frame(self.notebook, padding="5")
        self.notebook.add(merge_frame, text="数据合并")
        file_frame = ttk.LabelFrame(merge_frame, text="选择要合并的文件", padding="5")
        file_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=3)
        file_list_frame = ttk.Frame(file_frame)
        file_list_frame.grid(row=0, column=0, columnspan=3, pady=3, sticky=(tk.W, tk.E))
        self.file_listbox = tk.Listbox(file_list_frame, height=3, width=70, selectmode=tk.MULTIPLE)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(file_list_frame, orient="vertical", command=self.file_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_listbox.configure(yscrollcommand=scrollbar.set)
        btn_frame = ttk.Frame(file_frame); btn_frame.grid(row=1, column=0, columnspan=3, pady=3)
        ttk.Button(btn_frame, text="添加文件", command=self.add_files).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_frame, text="移除选中", command=self.remove_selected_files).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_frame, text="清空列表", command=self.clear_file_list).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_frame, text="预览列名", command=self.preview_columns).pack(side=tk.LEFT, padx=3)
        column_frame = ttk.LabelFrame(merge_frame, text="列选择", padding="5")
        column_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=3)
        ttk.Label(column_frame, text="按列名:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.columns_by_name = ttk.Entry(column_frame, width=45)
        self.columns_by_name.grid(row=0, column=1, columnspan=2, sticky=tk.W, pady=2)
        ttk.Label(column_frame, text="按列号:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.columns_by_index = ttk.Entry(column_frame, width=45)
        self.columns_by_index.grid(row=1, column=1, columnspan=2, sticky=tk.W, pady=2)
        option_frame = ttk.LabelFrame(merge_frame, text="合并选项", padding="5")
        option_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=3)
        ttk.Label(option_frame, text="方式:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.merge_type = tk.StringVar(value="vertical")
        ttk.Radiobutton(option_frame, text="垂直", variable=self.merge_type, value="vertical").grid(row=0, column=1, sticky=tk.W, padx=3)
        ttk.Radiobutton(option_frame, text="水平", variable=self.merge_type, value="horizontal").grid(row=0, column=2, sticky=tk.W, padx=3)
        self.remove_blank_rows_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(option_frame, text="去空白行", variable=self.remove_blank_rows_var).grid(row=1, column=0, sticky=tk.W, pady=2)
        self.remove_blank_cols_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(option_frame, text="去空白列", variable=self.remove_blank_cols_var).grid(row=1, column=1, sticky=tk.W, pady=2)
        self.remove_blank_cells_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(option_frame, text="去空白字符", variable=self.remove_blank_cells_var).grid(row=1, column=2, sticky=tk.W, pady=2)
        self.add_source_var = tk.BooleanVar()
        ttk.Checkbutton(option_frame, text="来源标识", variable=self.add_source_var).grid(row=2, column=0, sticky=tk.W, pady=2)
        self.remove_dup_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(option_frame, text="去重复行", variable=self.remove_dup_var).grid(row=2, column=1, sticky=tk.W, pady=2)
        self.merge_all_sheets_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(option_frame, text="合并所有Sheet", variable=self.merge_all_sheets_var).grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=2)
        output_frame = ttk.LabelFrame(merge_frame, text="输出设置", padding="5")
        output_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=3)
        ttk.Label(output_frame, text="输出:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.output_path = tk.StringVar()
        ttk.Entry(output_frame, textvariable=self.output_path, width=45).grid(row=0, column=1, padx=3)
        ttk.Button(output_frame, text="浏览", command=self.select_output_file).grid(row=0, column=2)
        self.merge_button = ttk.Button(merge_frame, text="开始合并", command=self.start_merge, width=15)
        self.merge_button.grid(row=4, column=0, columnspan=3, pady=3)
    
    def create_convert_tab(self):
        convert_frame = ttk.Frame(self.notebook, padding="5")
        self.notebook.add(convert_frame, text="格式转换")
        type_frame = ttk.LabelFrame(convert_frame, text="转换类型", padding="5")
        type_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=3)
        self.convert_type = tk.StringVar(value="excel_to_word")
        ttk.Radiobutton(type_frame, text="Excel→Word", variable=self.convert_type, value="excel_to_word").grid(row=0, column=0, padx=3, pady=2)
        ttk.Radiobutton(type_frame, text="Word→Excel", variable=self.convert_type, value="word_to_excel").grid(row=0, column=1, padx=3, pady=2)
        ttk.Radiobutton(type_frame, text="Excel→PDF", variable=self.convert_type, value="excel_to_pdf").grid(row=0, column=2, padx=3, pady=2)
        ttk.Radiobutton(type_frame, text="Excel→PPT", variable=self.convert_type, value="excel_to_ppt").grid(row=1, column=0, padx=3, pady=2)
        ttk.Radiobutton(type_frame, text="Excel→TXT", variable=self.convert_type, value="excel_to_txt").grid(row=1, column=1, padx=3, pady=2)
        ttk.Radiobutton(type_frame, text="Word→PDF", variable=self.convert_type, value="word_to_pdf").grid(row=1, column=2, padx=3, pady=2)
        ttk.Radiobutton(type_frame, text="PPT→PDF", variable=self.convert_type, value="ppt_to_pdf").grid(row=2, column=0, padx=3, pady=2)
        input_frame = ttk.LabelFrame(convert_frame, text="输入文件", padding="5")
        input_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=3)
        self.convert_input = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.convert_input, width=55).grid(row=0, column=0, padx=3)
        ttk.Button(input_frame, text="浏览", command=self.select_input_file).grid(row=0, column=1)
        output_frame = ttk.LabelFrame(convert_frame, text="输出文件", padding="5")
        output_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=3)
        self.convert_output = tk.StringVar()
        ttk.Entry(output_frame, textvariable=self.convert_output, width=55).grid(row=0, column=0, padx=3)
        ttk.Button(output_frame, text="浏览", command=self.select_convert_output).grid(row=0, column=1)
        self.convert_button = ttk.Button(convert_frame, text="开始转换", command=self.start_convert, width=15)
        self.convert_button.grid(row=3, column=0, columnspan=3, pady=3)
    
    def create_split_tab(self):
        split_frame = ttk.Frame(self.notebook, padding="5")
        self.notebook.add(split_frame, text="数据拆分")
        file_frame = ttk.LabelFrame(split_frame, text="选择文件", padding="5")
        file_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=3)
        self.split_file_path = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.split_file_path, width=55).grid(row=0, column=0, padx=3)
        ttk.Button(file_frame, text="浏览", command=self.select_split_file).grid(row=0, column=1)
        ttk.Button(file_frame, text="预览列名", command=self.preview_split_columns).grid(row=0, column=2, padx=3)
        method_frame = ttk.LabelFrame(split_frame, text="拆分方式", padding="5")
        method_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=3)
        self.split_method = tk.StringVar(value="by_columns")
        ttk.Radiobutton(method_frame, text="按列值", variable=self.split_method, value="by_columns", command=self.toggle_split_method).grid(row=0, column=0, padx=3)
        ttk.Radiobutton(method_frame, text="按列位置", variable=self.split_method, value="by_column_position", command=self.toggle_split_method).grid(row=0, column=1, padx=3)
        ttk.Radiobutton(method_frame, text="按行数", variable=self.split_method, value="by_rows", command=self.toggle_split_method).grid(row=0, column=2, padx=3)
        ttk.Radiobutton(method_frame, text="按特定行", variable=self.split_method, value="by_specific_rows", command=self.toggle_split_method).grid(row=0, column=3, padx=3)
        self.column_split_frame = ttk.LabelFrame(split_frame, text="按列值拆分设置", padding="5")
        self.column_split_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=3)
        ttk.Label(self.column_split_frame, text="拆分列:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.split_columns = ttk.Entry(self.column_split_frame, width=40)
        self.split_columns.grid(row=0, column=1, sticky=tk.W, pady=2)
        self.column_position_frame = ttk.LabelFrame(split_frame, text="按列位置拆分设置", padding="5")
        self.column_position_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=3)
        self.column_position_type = tk.StringVar(value="before_after")
        ttk.Radiobutton(self.column_position_frame, text="前后拆分", variable=self.column_position_type, value="before_after", command=self.toggle_column_position_type).grid(row=0, column=0, padx=3)
        ttk.Radiobutton(self.column_position_frame, text="指定列提取", variable=self.column_position_type, value="specific_columns", command=self.toggle_column_position_type).grid(row=0, column=1, padx=3)
        self.before_after_frame = ttk.Frame(self.column_position_frame)
        self.before_after_frame.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=2)
        ttk.Label(self.before_after_frame, text="拆分位置(多列用逗号):").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.split_position = ttk.Entry(self.before_after_frame, width=35)
        self.split_position.grid(row=0, column=1, sticky=tk.W, pady=2)
        ttk.Button(self.before_after_frame, text="从预览选择", command=self.preview_position_columns).grid(row=0, column=2, padx=3)
        self.specific_columns_frame = ttk.Frame(self.column_position_frame)
        self.specific_columns_frame.grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=2)
        ttk.Label(self.specific_columns_frame, text="指定列号:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.specific_column_numbers = ttk.Entry(self.specific_columns_frame, width=35)
        self.specific_column_numbers.grid(row=0, column=1, sticky=tk.W, pady=2)
        ttk.Button(self.specific_columns_frame, text="从预览选择", command=self.preview_specific_columns).grid(row=0, column=2, padx=3)
        self.row_split_frame = ttk.LabelFrame(split_frame, text="按行数拆分设置", padding="5")
        self.row_split_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=3)
        ttk.Label(self.row_split_frame, text="每文件行数:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.rows_per_file = ttk.Entry(self.row_split_frame, width=15)
        self.rows_per_file.grid(row=0, column=1, sticky=tk.W, pady=2)
        self.rows_per_file.insert(0, "1000")
        self.specific_row_frame = ttk.LabelFrame(split_frame, text="按特定行拆分设置", padding="5")
        self.specific_row_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=3)
        ttk.Label(self.specific_row_frame, text="拆分行号:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.specific_rows = ttk.Entry(self.specific_row_frame, width=35)
        self.specific_rows.grid(row=0, column=1, sticky=tk.W, pady=2)
        output_frame = ttk.LabelFrame(split_frame, text="输出设置", padding="5")
        output_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=3)
        ttk.Label(output_frame, text="输出目录:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.split_output_dir = tk.StringVar(value="./split_output")
        ttk.Entry(output_frame, textvariable=self.split_output_dir, width=35).grid(row=0, column=1, pady=2)
        ttk.Button(output_frame, text="浏览", command=self.select_split_output).grid(row=0, column=2, padx=3)
        self.split_button = ttk.Button(split_frame, text="开始拆分", command=self.start_split, width=15)
        self.split_button.grid(row=7, column=0, columnspan=3, pady=3)
        self.column_position_frame.grid_remove(); self.row_split_frame.grid_remove(); self.specific_row_frame.grid_remove(); self.specific_columns_frame.grid_remove()
    
    def create_batch_tab(self):
        batch_frame = ttk.Frame(self.notebook, padding="5")
        self.notebook.add(batch_frame, text="批量处理")
        dir_frame = ttk.LabelFrame(batch_frame, text="选择文件夹", padding="5")
        dir_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=3)
        self.batch_dir = tk.StringVar()
        ttk.Entry(dir_frame, textvariable=self.batch_dir, width=55).grid(row=0, column=0, padx=3)
        ttk.Button(dir_frame, text="浏览", command=self.select_batch_dir).grid(row=0, column=1)
        option_frame = ttk.LabelFrame(batch_frame, text="处理选项", padding="5")
        option_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=3)
        self.batch_operation = tk.StringVar(value="clean")
        ttk.Radiobutton(option_frame, text="清理数据", variable=self.batch_operation, value="clean").grid(row=0, column=0, padx=10)
        ttk.Radiobutton(option_frame, text="去除空行", variable=self.batch_operation, value="remove_empty").grid(row=0, column=1, padx=10)
        ttk.Radiobutton(option_frame, text="公式转数值", variable=self.batch_operation, value="formula_to_value").grid(row=0, column=2, padx=10)
        self.batch_button = ttk.Button(batch_frame, text="开始批量处理", command=self.start_batch, width=15)
        self.batch_button.grid(row=2, column=0, columnspan=2, pady=3)
    
    def toggle_split_method(self):
        method = self.split_method.get()
        self.column_split_frame.grid_remove(); self.column_position_frame.grid_remove(); self.row_split_frame.grid_remove(); self.specific_row_frame.grid_remove()
        if method == "by_columns": self.column_split_frame.grid()
        elif method == "by_column_position":
            self.column_position_frame.grid(); self.toggle_column_position_type()
        elif method == "by_rows": self.row_split_frame.grid()
        elif method == "by_specific_rows": self.specific_row_frame.grid()
    
    def toggle_column_position_type(self):
        pos_type = self.column_position_type.get()
        self.before_after_frame.grid_remove(); self.specific_columns_frame.grid_remove()
        if pos_type == "before_after": self.before_after_frame.grid()
        elif pos_type == "specific_columns": self.specific_columns_frame.grid()
    
    # ========== 预览方法 ==========
    def preview_columns(self):
        if self.is_processing: return
        files = list(self.file_listbox.get(0, tk.END))
        if not files:
            messagebox.showwarning("警告", "请先添加文件"); return
        try:
            reference_columns = None
            for file in files:
                cols, has_header = self.get_columns_quick(file)
                if has_header:
                    reference_columns = cols; break
            if reference_columns is None:
                cols, _ = self.get_columns_quick(files[0])
                reference_columns = cols if cols else [f"列{i+1}" for i in range(10)]
            self._show_column_preview(reference_columns, self.columns_by_name)
        except Exception as e:
            messagebox.showerror("错误", f"读取文件失败: {str(e)}")
    
    def preview_split_columns(self):
        if self.is_processing: return
        file_path = self.split_file_path.get()
        if not file_path:
            messagebox.showwarning("警告", "请先选择文件"); return
        try:
            cols, _ = self.get_columns_quick(file_path)
            self._show_column_preview(cols, self.split_columns)
        except Exception as e:
            messagebox.showerror("错误", f"读取文件失败: {str(e)}")
    
    def preview_position_columns(self):
        if self.is_processing: return
        file_path = self.split_file_path.get()
        if not file_path:
            messagebox.showwarning("警告", "请先选择文件"); return
        try:
            cols, _ = self.get_columns_quick(file_path)
            win = tk.Toplevel(self.root); win.title("选择拆分位置（可多选）"); win.geometry("400x500")
            lb = tk.Listbox(win, selectmode=tk.MULTIPLE, height=20, width=50)
            lb.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
            for i, col in enumerate(cols, 1): lb.insert(tk.END, f"{i}. {col}")
            btn_frame = ttk.Frame(win); btn_frame.pack(pady=10)
            def set_positions():
                selected = lb.curselection()
                if not selected: return
                nums = sorted([idx+1 for idx in selected])
                self.split_position.delete(0, tk.END); self.split_position.insert(0, ','.join(map(str, nums)))
                win.destroy()
            ttk.Button(btn_frame, text="设置拆分位置", command=set_positions).pack(side=tk.LEFT, padx=5)
            ttk.Button(btn_frame, text="关闭", command=win.destroy).pack(side=tk.LEFT, padx=5)
        except Exception as e:
            messagebox.showerror("错误", f"读取文件失败: {str(e)}")
    
    def preview_specific_columns(self):
        if self.is_processing: return
        file_path = self.split_file_path.get()
        if not file_path:
            messagebox.showwarning("警告", "请先选择文件"); return
        try:
            cols, _ = self.get_columns_quick(file_path)
            win = tk.Toplevel(self.root); win.title("选择指定列"); win.geometry("400x500")
            lb = tk.Listbox(win, selectmode=tk.MULTIPLE, height=20, width=50)
            lb.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
            for i, col in enumerate(cols, 1): lb.insert(tk.END, f"{i}. {col}")
            btn_frame = ttk.Frame(win); btn_frame.pack(pady=10)
            def add_columns():
                selected = lb.curselection()
                if not selected: return
                nums = sorted([idx+1 for idx in selected])
                cur = self.specific_column_numbers.get().strip()
                if cur:
                    existing = [n.strip() for n in cur.split(',') if n.strip()]
                    for n in nums:
                        if str(n) not in existing: existing.append(str(n))
                    new_text = ','.join(existing)
                else: new_text = ','.join(map(str, nums))
                self.specific_column_numbers.delete(0, tk.END); self.specific_column_numbers.insert(0, new_text)
                win.destroy()
            ttk.Button(btn_frame, text="添加选中列号", command=add_columns).pack(side=tk.LEFT, padx=5)
            ttk.Button(btn_frame, text="关闭", command=win.destroy).pack(side=tk.LEFT, padx=5)
        except Exception as e:
            messagebox.showerror("错误", f"读取文件失败: {str(e)}")
    
    def _show_column_preview(self, columns, target_entry):
        win = tk.Toplevel(self.root); win.title("列名预览"); win.geometry("400x500")
        lb = tk.Listbox(win, selectmode=tk.MULTIPLE, height=20, width=50)
        lb.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        for i, col in enumerate(columns, 1): lb.insert(tk.END, f"{i}. {col}")
        btn_frame = ttk.Frame(win); btn_frame.pack(pady=10)
        def add_selected():
            selected = lb.curselection()
            if not selected: return
            selected_cols = []
            for idx in selected:
                display = lb.get(idx)
                col_name = re.sub(r'^\d+\.\s*', '', display)
                selected_cols.append(col_name)
            cur = target_entry.get().strip()
            if cur:
                existing = [c.strip() for c in cur.split(',') if c.strip()]
                for col in selected_cols:
                    if col not in existing: existing.append(col)
                new_text = ', '.join(existing)
            else: new_text = ', '.join(selected_cols)
            target_entry.delete(0, tk.END); target_entry.insert(0, new_text)
            win.destroy()
        ttk.Button(btn_frame, text="添加选中列名", command=add_selected).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="关闭", command=win.destroy).pack(side=tk.LEFT, padx=5)
    
    def parse_column_selection(self, columns_str, df_columns):
        selected = []
        if self.columns_by_name.get().strip():
            names = [c.strip() for c in self.columns_by_name.get().split(',') if c.strip()]
            for name in names:
                if name in df_columns: selected.append(name)
        if self.columns_by_index.get().strip():
            idx_str = self.columns_by_index.get().strip()
            indices = []
            for part in idx_str.split(','):
                part = part.strip()
                if '-' in part:
                    s, e = part.split('-')
                    indices.extend(range(int(s), int(e)+1))
                else: indices.append(int(part))
            for idx in indices:
                if 1 <= idx <= len(df_columns): selected.append(df_columns[idx-1])
        return list(dict.fromkeys(selected))
    
    # ========== 文件选择方法 ==========
    def add_files(self):
        if self.is_processing: return
        files = filedialog.askopenfilenames(title="选择文件", filetypes=[("所有支持的文件","*.xlsx *.xls *.csv"),("Excel文件","*.xlsx *.xls"),("CSV文件","*.csv"),("所有文件","*.*")])
        for f in files:
            if f not in self.file_listbox.get(0, tk.END):
                self.file_listbox.insert(tk.END, f)
                self.log(f"添加文件: {os.path.basename(f)}")
    def remove_selected_files(self):
        if self.is_processing: return
        selected = self.file_listbox.curselection()
        for idx in reversed(selected): self.file_listbox.delete(idx)
    def clear_file_list(self):
        if self.is_processing: return
        self.file_listbox.delete(0, tk.END)
    def select_output_file(self):
        if self.is_processing: return
        path = filedialog.asksaveasfilename(title="保存文件", defaultextension=".xlsx", filetypes=[("Excel文件","*.xlsx"),("CSV文件","*.csv"),("所有文件","*.*")])
        if path: self.output_path.set(path)
    def select_input_file(self):
        if self.is_processing: return
        ct = self.convert_type.get()
        if ct in ["excel_to_word","excel_to_pdf","excel_to_ppt","excel_to_txt"]:
            path = filedialog.askopenfilename(title="选择文件", filetypes=[("所有支持的文件","*.xlsx *.xls *.csv"),("Excel文件","*.xlsx *.xls"),("CSV文件","*.csv")])
        elif ct in ["word_to_excel","word_to_pdf"]:
            path = filedialog.askopenfilename(title="选择Word文件", filetypes=[("Word文件","*.docx *.doc")])
        elif ct == "ppt_to_pdf":
            path = filedialog.askopenfilename(title="选择PPT文件", filetypes=[("PPT文件","*.pptx *.ppt")])
        else: path = filedialog.askopenfilename(title="选择文件", filetypes=[("所有文件","*.*")])
        if path: self.convert_input.set(path)
    def select_convert_output(self):
        if self.is_processing: return
        ct = self.convert_type.get()
        if ct == "excel_to_word":
            path = filedialog.asksaveasfilename(title="保存Word文件", defaultextension=".docx", filetypes=[("Word文件","*.docx")])
        elif ct == "word_to_excel":
            path = filedialog.asksaveasfilename(title="保存文件", defaultextension=".xlsx", filetypes=[("Excel文件","*.xlsx"),("CSV文件","*.csv")])
        elif ct == "excel_to_txt":
            path = filedialog.asksaveasfilename(title="保存文本文件", defaultextension=".txt", filetypes=[("文本文件","*.txt")])
        elif ct in ["excel_to_pdf","word_to_pdf","ppt_to_pdf"]:
            path = filedialog.asksaveasfilename(title="保存PDF文件", defaultextension=".pdf", filetypes=[("PDF文件","*.pdf")])
        elif ct == "excel_to_ppt":
            path = filedialog.asksaveasfilename(title="保存PPT文件", defaultextension=".pptx", filetypes=[("PPT文件","*.pptx")])
        else: path = filedialog.asksaveasfilename(title="保存文件", filetypes=[("所有文件","*.*")])
        if path: self.convert_output.set(path)
    def select_split_file(self):
        if self.is_processing: return
        path = filedialog.askopenfilename(title="选择文件", filetypes=[("所有支持的文件","*.xlsx *.xls *.csv"),("Excel文件","*.xlsx *.xls"),("CSV文件","*.csv")])
        if path: self.split_file_path.set(path)
    def select_split_output(self):
        if self.is_processing: return
        path = filedialog.askdirectory(title="选择输出目录")
        if path: self.split_output_dir.set(path)
    def select_batch_dir(self):
        if self.is_processing: return
        path = filedialog.askdirectory(title="选择文件夹")
        if path: self.batch_dir.set(path)
    def update_file_types(self): pass
    
    # ========== 处理功能 ==========
    def start_merge(self):
        if self.is_processing: return
        files = list(self.file_listbox.get(0, tk.END))
        if not files: messagebox.showwarning("警告", "请先添加要合并的文件"); return
        output_path = self.output_path.get()
        if not output_path: messagebox.showwarning("警告", "请指定输出文件路径"); return
        threading.Thread(target=self.merge_thread, args=(files, output_path), daemon=True).start()
    
    def merge_thread(self, files, output_path):
        try:
            self.start_indeterminate("正在合并文件...")
            self.log("开始合并文件...")
            reference_columns = None
            for f in files:
                cols, has_header = self.get_columns_quick(f)
                if has_header:
                    reference_columns = cols; break
            if reference_columns is None:
                cols, _ = self.get_columns_quick(files[0])
                reference_columns = cols if cols else [f"列{i+1}" for i in range(10)]
            selected_columns = self.parse_column_selection(self.columns_by_name.get(), reference_columns)
            dfs = []
            total = len(files)
            for i, file in enumerate(files):
                self.check_cancel()
                if i == 0 or i == total-1 or i % 5 == 0:
                    self.log(f"读取文件 {i+1}/{total}: {os.path.basename(file)}")
                self.update_progress((i / total) * 40, f"读取文件 {i+1}/{total}")
                ext = os.path.splitext(file)[1].lower()
                if self.merge_all_sheets_var.get():
                    if ext in ('.xlsx', '.xls'):
                        self.log(f"正在读取所有Sheet: {os.path.basename(file)}")
                        try:
                            all_sheets = pd.read_excel(file, sheet_name=None, dtype=object)
                            if all_sheets:
                                df = pd.concat(all_sheets.values(), ignore_index=True)
                                self.log(f"已合并 {len(all_sheets)} 个Sheet，共 {len(df)} 行")
                            else: df = pd.DataFrame()
                        except Exception as e:
                            self.log(f"读取所有Sheet失败: {str(e)}，尝试读取第一个Sheet")
                            df = self.read_excel_file(file)
                    else:
                        self.log(f"CSV文件没有多个Sheet，按单个Sheet读取: {os.path.basename(file)}")
                        df = self.read_excel_file(file)
                else:
                    df = self.read_excel_file(file)
                    if not self.detect_header(df):
                        if ext == '.csv': df = self.read_csv_without_header(file)
                        else: df = self.read_excel_without_header(file)
                df = self.normalize_columns(df, reference_columns)
                if selected_columns:
                    avail = [c for c in selected_columns if c in df.columns]
                    df = df[avail]
                if self.remove_blank_rows_var.get() or self.remove_blank_cols_var.get() or self.remove_blank_cells_var.get():
                    df = self.remove_blank_data(df,
                        remove_blank_rows=self.remove_blank_rows_var.get(),
                        remove_blank_cols=self.remove_blank_cols_var.get(),
                        remove_blank_cells=self.remove_blank_cells_var.get())
                if self.add_source_var.get():
                    df['数据来源'] = os.path.basename(file)
                dfs.append(df)
            self.check_cancel()
            self.update_progress(60, "正在合并数据...", force_update=True)
            if self.merge_type.get() == "vertical":
                merged = pd.concat(dfs, ignore_index=True)
            else:
                merged = pd.concat(dfs, axis=1)
            self.update_progress(75, "正在处理数据...", force_update=True)
            if self.remove_blank_rows_var.get() or self.remove_blank_cols_var.get():
                merged = self.remove_blank_data(merged,
                    remove_blank_rows=self.remove_blank_rows_var.get(),
                    remove_blank_cols=self.remove_blank_cols_var.get(),
                    remove_blank_cells=False)
            if self.remove_dup_var.get():
                self.check_cancel()
                merged = merged.drop_duplicates()
            self.update_progress(90, "正在保存结果...", force_update=True)
            self.save_excel_file(merged, output_path)
            self.update_progress(100, "合并完成", force_update=True)
            self.log(f"合并完成！结果已保存到: {output_path}")
            self.stop_indeterminate("合并完成", success=True)
            messagebox.showinfo("成功", f"合并完成！\n输出文件: {output_path}")
        except Exception as e:
            if "取消" in str(e): self.stop_indeterminate("已取消", success=False)
            else:
                self.stop_indeterminate("处理失败", success=False)
                self.log(f"合并失败: {str(e)}")
                messagebox.showerror("错误", f"合并失败: {str(e)}")
    
    def start_convert(self):
        if self.is_processing: return
        input_path = self.convert_input.get(); output_path = self.convert_output.get()
        if not input_path: messagebox.showwarning("警告", "请选择输入文件"); return
        if not output_path: messagebox.showwarning("警告", "请指定输出文件路径"); return
        threading.Thread(target=self.convert_thread, args=(input_path, output_path), daemon=True).start()
    
    def convert_thread(self, input_path, output_path):
        try:
            self.start_indeterminate("正在转换...")
            ct = self.convert_type.get()
            if ct == "excel_to_word": self.convert_excel_to_word(input_path, output_path)
            elif ct == "word_to_excel": self.convert_word_to_excel(input_path, output_path)
            elif ct == "excel_to_pdf": self.convert_excel_to_pdf(input_path, output_path)
            elif ct == "excel_to_ppt": self.convert_excel_to_ppt(input_path, output_path)
            elif ct == "excel_to_txt": self.convert_excel_to_txt(input_path, output_path)
            elif ct == "word_to_pdf": self.convert_word_to_pdf(input_path, output_path)
            elif ct == "ppt_to_pdf": self.convert_ppt_to_pdf(input_path, output_path)
            self.update_progress(100, "转换完成", force_update=True)
            self.stop_indeterminate("转换完成", success=True)
            messagebox.showinfo("成功", f"转换完成！\n输出文件: {output_path}")
        except Exception as e:
            if "取消" in str(e): self.stop_indeterminate("已取消", success=False)
            else:
                self.stop_indeterminate("转换失败", success=False)
                self.log(f"转换失败: {str(e)}")
                messagebox.showerror("错误", f"转换失败: {str(e)}")
    
    def convert_excel_to_txt(self, input_path, output_path):
        self.update_progress(20, "正在读取文件...", force_update=True)
        df = self.read_excel_file(input_path)
        if not self.detect_header(df):
            ext = os.path.splitext(input_path)[1].lower()
            if ext == '.csv': df = self.read_csv_without_header(input_path)
            else: df = self.read_excel_without_header(input_path)
        df = self.remove_blank_data(df)
        self.check_cancel()
        self.update_progress(70, "正在保存文本文件...", force_update=True)
        df.to_csv(output_path, index=False, sep='\t', encoding='utf-8-sig')
        self.log(f"Excel转TXT完成: {output_path}")
    
    def convert_excel_to_word(self, input_path, output_path):
        from docx import Document
        self.update_progress(20, "正在读取文件...", force_update=True)
        df = self.read_excel_file(input_path)
        if not self.detect_header(df):
            ext = os.path.splitext(input_path)[1].lower()
            if ext == '.csv': df = self.read_csv_without_header(input_path)
            else: df = self.read_excel_without_header(input_path)
        df = self.remove_blank_data(df)
        self.check_cancel()
        self.update_progress(50, "正在创建Word文档...", force_update=True)
        doc = Document(); doc.add_heading('数据转换结果', level=1)
        table = doc.add_table(rows=1, cols=len(df.columns)); table.style = 'Light Grid Accent 1'
        for i, col in enumerate(df.columns): table.rows[0].cells[i].text = str(col)
        total_rows = len(df); batch = max(1, total_rows//10)
        for idx, (_, row) in enumerate(df.iterrows()):
            self.check_cancel()
            cells = table.add_row().cells
            for i, val in enumerate(row):
                cells[i].text = '' if pd.isna(val) else str(val)
            if idx % batch == 0:
                self.update_progress(50 + (idx/total_rows)*40)
        self.update_progress(90, "正在保存Word文件...", force_update=True)
        doc.save(output_path)
        self.log(f"Excel转Word完成: {output_path}")
    
    def convert_word_to_excel(self, input_path, output_path):
        from docx import Document
        self.update_progress(30, "正在读取Word文件...", force_update=True)
        doc = Document(input_path)
        if not doc.tables: raise Exception("Word文档中没有表格")
        self.check_cancel()
        self.update_progress(60, "正在提取表格数据...", force_update=True)
        data = [[cell.text for cell in row.cells] for row in doc.tables[0].rows]
        df = pd.DataFrame(data[1:], columns=data[0])
        df = self.remove_blank_data(df)
        self.save_excel_file(df, output_path)
        self.log(f"Word转Excel完成: {output_path}")
    
    def convert_excel_to_pdf(self, input_path, output_path):
        try:
            from reportlab.lib.pagesizes import A4, landscape
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
            from reportlab.lib import colors
        except ImportError:
            raise Exception("缺少reportlab库，请安装: pip install reportlab")
        self.update_progress(20, "正在读取Excel文件...", force_update=True)
        df = self.read_excel_file(input_path)
        if not self.detect_header(df):
            ext = os.path.splitext(input_path)[1].lower()
            if ext == '.csv': df = self.read_csv_without_header(input_path)
            else: df = self.read_excel_without_header(input_path)
        df = self.remove_blank_data(df)
        self.check_cancel()
        self.update_progress(50, "正在创建PDF...", force_update=True)
        doc = SimpleDocTemplate(output_path, pagesize=landscape(A4))
        data = [df.columns.tolist()] + df.values.tolist()
        table = Table(data)
        style = TableStyle([
            ('BACKGROUND',(0,0),(-1,0),colors.grey),
            ('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke),
            ('ALIGN',(0,0),(-1,-1),'CENTER'),
            ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
            ('FONTSIZE',(0,0),(-1,0),12),
            ('BOTTOMPADDING',(0,0),(-1,0),12),
            ('BACKGROUND',(0,1),(-1,-1),colors.beige),
            ('GRID',(0,0),(-1,-1),1,colors.black),
            ('FONTSIZE',(0,1),(-1,-1),8),
        ])
        table.setStyle(style)
        doc.build([table])
        self.log(f"Excel转PDF完成: {output_path}")
    
    def convert_excel_to_ppt(self, input_path, output_path):
        try:
            from pptx import Presentation
            from pptx.util import Inches
        except ImportError:
            raise Exception("缺少python-pptx库，请安装: pip install python-pptx")
        self.update_progress(20, "正在读取Excel文件...", force_update=True)
        df = self.read_excel_file(input_path)
        if not self.detect_header(df):
            ext = os.path.splitext(input_path)[1].lower()
            if ext == '.csv': df = self.read_csv_without_header(input_path)
            else: df = self.read_excel_without_header(input_path)
        df = self.remove_blank_data(df)
        self.check_cancel()
        self.update_progress(50, "正在创建PPT...", force_update=True)
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = "数据展示"
        slide.placeholders[1].text = f"数据行数: {len(df)}"
        slide2 = prs.slides.add_slide(prs.slide_layouts[5])
        slide2.shapes.title.text = "数据详情"
        rows, cols = min(len(df)+1,20), min(len(df.columns),8)
        table_shape = slide2.shapes.add_table(rows, cols, Inches(0.5), Inches(1.5), Inches(9), Inches(5))
        table = table_shape.table
        for i, col in enumerate(df.columns[:cols]): table.cell(0,i).text = str(col)
        for i in range(1, rows):
            for j in range(cols):
                val = df.iloc[i-1,j]
                table.cell(i,j).text = '' if pd.isna(val) else str(val)
        prs.save(output_path)
        self.log(f"Excel转PPT完成: {output_path}")
    
    def convert_word_to_pdf(self, input_path, output_path):
        try:
            from docx2pdf import convert
            self.update_progress(50, "正在转换...", force_update=True)
            convert(input_path, output_path)
            self.log(f"Word转PDF完成: {output_path}")
        except ImportError:
            from docx import Document
            from reportlab.lib.pagesizes import A4
            from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
            from reportlab.lib.styles import getSampleStyleSheet
            self.update_progress(30, "正在读取Word文件...", force_update=True)
            doc = Document(input_path)
            self.check_cancel()
            self.update_progress(60, "正在创建PDF...", force_update=True)
            pdf = SimpleDocTemplate(output_path, pagesize=A4)
            styles = getSampleStyleSheet()
            elements = [Paragraph(p.text, styles['Normal']) for p in doc.paragraphs if p.text.strip()]
            pdf.build(elements)
            self.log(f"Word转PDF完成: {output_path}")
    
    def convert_ppt_to_pdf(self, input_path, output_path):
        try:
            import win32com.client
            self.update_progress(50, "正在转换...", force_update=True)
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            presentation = powerpoint.Presentations.Open(input_path)
            presentation.SaveAs(output_path, 32)
            presentation.Close(); powerpoint.Quit()
            self.log(f"PPT转PDF完成: {output_path}")
        except ImportError:
            from pptx import Presentation
            from reportlab.lib.pagesizes import A4, landscape
            from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
            from reportlab.lib.styles import getSampleStyleSheet
            self.update_progress(30, "正在读取PPT文件...", force_update=True)
            prs = Presentation(input_path)
            self.check_cancel()
            self.update_progress(60, "正在创建PDF...", force_update=True)
            pdf = SimpleDocTemplate(output_path, pagesize=landscape(A4))
            styles = getSampleStyleSheet()
            elements = []
            for slide in prs.slides:
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        for para in shape.text_frame.paragraphs:
                            if para.text.strip(): elements.append(Paragraph(para.text, styles['Normal']))
                elements.append(Spacer(1,30))
            pdf.build(elements)
            self.log(f"PPT转PDF完成: {output_path}")
    
    def start_split(self):
        if self.is_processing: return
        file_path = self.split_file_path.get()
        if not file_path: messagebox.showwarning("警告", "请选择要拆分的文件"); return
        threading.Thread(target=self.split_thread, args=(file_path,), daemon=True).start()
    
    def split_thread(self, file_path):
        try:
            self.start_indeterminate("正在拆分数据...")
            self.update_progress(10, "正在读取文件...", force_update=True)
            df = self.read_excel_file(file_path)
            if not self.detect_header(df):
                ext = os.path.splitext(file_path)[1].lower()
                if ext == '.csv': df = self.read_csv_without_header(file_path)
                else: df = self.read_excel_without_header(file_path)
            self.log("正在去除空白数据...")
            df = self.remove_blank_data(df)
            output_dir = self.split_output_dir.get()
            if not os.path.exists(output_dir): os.makedirs(output_dir)
            method = self.split_method.get()
            if method == "by_columns": self.split_by_columns(df, output_dir)
            elif method == "by_column_position": self.split_by_column_position(df, output_dir)
            elif method == "by_rows": self.split_by_rows(df, output_dir)
            elif method == "by_specific_rows": self.split_by_specific_rows(df, output_dir)
            self.update_progress(100, "拆分完成", force_update=True)
            self.log("拆分完成！")
            self.stop_indeterminate("拆分完成", success=True)
            messagebox.showinfo("成功", f"拆分完成！\n输出目录: {output_dir}")
        except Exception as e:
            if "取消" in str(e): self.stop_indeterminate("已取消", success=False)
            else:
                self.stop_indeterminate("拆分失败", success=False)
                self.log(f"拆分失败: {str(e)}")
                messagebox.showerror("错误", f"拆分失败: {str(e)}")
    
    def split_by_columns(self, df, output_dir):
        cols_str = self.split_columns.get().strip()
        if not cols_str: raise Exception("请输入拆分依据列")
        cols = [c.strip() for c in cols_str.split(',') if c.strip()]
        missing = [c for c in cols if c not in df.columns]
        if missing: raise Exception(f"列不存在: {missing}")
        groups = df.groupby(cols)
        for i, (name, group) in enumerate(groups):
            self.check_cancel()
            if isinstance(name, tuple): safe = '_'.join([re.sub(r'[\\/*?:"<>|]','_',str(n)) for n in name])
            else: safe = re.sub(r'[\\/*?:"<>|]','_',str(name))
            path = os.path.join(output_dir, f"{safe}.xlsx")
            self.save_excel_file(group, path)
            if i % max(1, len(groups)//20) == 0:
                self.update_progress(20 + ((i+1)/len(groups))*70)
        self.log(f"按列拆分完成！共生成 {len(groups)} 个文件")
    
    def split_by_column_position(self, df, output_dir):
        pos_type = self.column_position_type.get()
        if pos_type == "before_after":
            pos_str = self.split_position.get().strip()
            if not pos_str: raise Exception("请输入拆分位置列号")
            try:
                positions = sorted(set([int(p) for p in pos_str.split(',') if p.strip()]))
                valid = [p for p in positions if 1 <= p < len(df.columns)]
                if not valid: raise Exception("没有有效的拆分位置")
            except ValueError: raise Exception("无效的拆分位置格式")
            points = [0] + valid + [len(df.columns)]
            for i in range(len(points)-1):
                self.check_cancel()
                start, end = points[i], points[i+1]
                chunk = df.iloc[:, start:end]
                if i == 0: fname = f"列1-{end}.xlsx"
                elif i == len(points)-2: fname = f"列{start+1}-{len(df.columns)}.xlsx"
                else: fname = f"列{start+1}-{end}.xlsx"
                path = os.path.join(output_dir, fname)
                self.save_excel_file(chunk, path)
                self.update_progress(20 + ((i+1)/(len(points)-1))*70, force_update=True)
            self.log(f"按多列位置拆分完成！共生成 {len(points)-1} 个文件")
        else:
            cols_str = self.specific_column_numbers.get().strip()
            if not cols_str: raise Exception("请输入指定列号")
            try:
                idxs = []
                for part in cols_str.split(','):
                    part = part.strip()
                    if '-' in part:
                        s,e = part.split('-'); idxs.extend(range(int(s), int(e)+1))
                    else: idxs.append(int(part))
                idxs = sorted(set(idxs))
                valid = [i for i in idxs if 1 <= i <= len(df.columns)]
                if not valid: raise Exception("没有有效的列号")
            except ValueError: raise Exception("无效的列号格式")
            chunk = df.iloc[:, [i-1 for i in valid]]
            path = os.path.join(output_dir, f"指定列_{'_'.join(map(str,valid))}.xlsx")
            self.save_excel_file(chunk, path)
            self.log("按指定列提取完成！共生成 1 个文件")
    
    def split_by_rows(self, df, output_dir):
        try:
            rows_per_file = int(self.rows_per_file.get())
            if rows_per_file <= 0: raise ValueError
        except ValueError: raise Exception("每个文件行数必须大于0")
        total = len(df); files = (total + rows_per_file -1)//rows_per_file
        for i in range(files):
            self.check_cancel()
            start = i*rows_per_file; end = min((i+1)*rows_per_file, total)
            chunk = df.iloc[start:end]
            path = os.path.join(output_dir, f"part_{i+1:03d}.xlsx")
            self.save_excel_file(chunk, path)
            self.update_progress(20 + ((i+1)/files)*70)
        self.log(f"按行数拆分完成！共生成 {files} 个文件")
    
    def split_by_specific_rows(self, df, output_dir):
        rows_str = self.specific_rows.get().strip()
        if not rows_str: raise Exception("请输入拆分行号")
        try:
            rows = sorted([int(r) for r in rows_str.split(',') if r.strip()])
        except ValueError: raise Exception("无效的行号格式")
        if not rows: raise Exception("请输入有效的行号")
        points = [0] + rows + [len(df)]
        for i in range(len(points)-1):
            self.check_cancel()
            start, end = points[i], points[i+1]
            chunk = df.iloc[start:end]
            path = os.path.join(output_dir, f"part_{i+1:03d}.xlsx")
            self.save_excel_file(chunk, path)
            self.update_progress(20 + ((i+1)/(len(points)-1))*70)
        self.log(f"按特定行拆分完成！共生成 {len(points)-1} 个文件")
    
    def start_batch(self):
        if self.is_processing: return
        directory = self.batch_dir.get()
        if not directory: messagebox.showwarning("警告", "请选择要处理的文件夹"); return
        threading.Thread(target=self.batch_thread, args=(directory,), daemon=True).start()
    
    def batch_thread(self, directory):
        try:
            self.start_indeterminate("正在批量处理...")
            files = [f for f in os.listdir(directory) if f.endswith(('.xlsx','.xls','.csv'))]
            if not files:
                self.stop_indeterminate("未找到文件", success=False)
                messagebox.showwarning("警告", "目录中没有支持的文件"); return
            out_dir = os.path.join(directory, 'processed')
            if not os.path.exists(out_dir): os.makedirs(out_dir)
            op = self.batch_operation.get()
            for i, fname in enumerate(files):
                self.check_cancel()
                fpath = os.path.join(directory, fname)
                if i%5==0 or i==len(files)-1: self.log(f"处理文件 {i+1}/{len(files)}: {fname}")
                self.update_progress((i/len(files))*100)
                df = self.read_excel_file(fpath)
                if not self.detect_header(df):
                    ext = os.path.splitext(fpath)[1].lower()
                    if ext == '.csv': df = self.read_csv_without_header(fpath)
                    else: df = self.read_excel_without_header(fpath)
                if op == "clean":
                    df = self.remove_blank_data(df); df = df.drop_duplicates()
                elif op == "remove_empty": df = df.dropna()
                elif op == "formula_to_value":
                    for col in df.columns:
                        try: df[col] = pd.to_numeric(df[col], errors='ignore')
                        except: pass
                    self.log("公式已转换为纯数值（列头保持不变）")
                out_path = os.path.join(out_dir, f"processed_{fname}")
                self.save_excel_file(df, out_path)
            self.update_progress(100, "批量处理完成", force_update=True)
            self.log(f"批量处理完成！处理了 {len(files)} 个文件")
            self.stop_indeterminate("批量处理完成", success=True)
            messagebox.showinfo("成功", f"批量处理完成！\n处理了 {len(files)} 个文件")
        except Exception as e:
            if "取消" in str(e): self.stop_indeterminate("已取消", success=False)
            else:
                self.stop_indeterminate("处理失败", success=False)
                self.log(f"批量处理失败: {str(e)}")
                messagebox.showerror("错误", f"批量处理失败: {str(e)}")

def main():
    root = tk.Tk()
    app = DataProcessingGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
