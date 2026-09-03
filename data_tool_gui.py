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
        self.root.geometry("1200x800")
        
        # 设置样式
        style = ttk.Style()
        style.theme_use('clam')
        
        # 创建主Canvas和滚动条
        self.main_canvas = tk.Canvas(root, highlightthickness=0)
        self.main_scrollbar = ttk.Scrollbar(root, orient="vertical", command=self.main_canvas.yview)
        self.main_canvas.configure(yscrollcommand=self.main_scrollbar.set)
        
        # 布局主Canvas和滚动条
        self.main_canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.main_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # 创建主框架（放在Canvas中）
        self.main_frame = ttk.Frame(self.main_canvas, padding="10")
        self.canvas_window = self.main_canvas.create_window((0, 0), window=self.main_frame, anchor="nw")
        
        # 绑定配置事件
        self.main_frame.bind("<Configure>", self.on_frame_configure)
        self.main_canvas.bind("<Configure>", self.on_canvas_configure)
        
        # 绑定鼠标滚轮事件
        self.bind_mousewheel_recursive(root)
        
        # 创建标签页
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 创建各个功能标签页
        self.create_merge_tab()
        self.create_convert_tab()
        self.create_split_tab()
        self.create_batch_tab()
        
        # 创建进度条区域
        self.create_progress_area()
        
        # 日志区域
        self.create_log_area()
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.rowconfigure(0, weight=1)
        
        # 初始化处理状态
        self.is_processing = False
        self.last_progress_time = 0
        
        # 设置窗口最小大小
        self.root.minsize(800, 600)
        
        # 设置窗口居中
        self.center_window()
        
        # 保存所有可交互的控件
        self.interactive_widgets = []
        self.collect_interactive_widgets()
    
    def collect_interactive_widgets(self):
        """收集所有可交互的控件"""
        for widget in self.get_all_widgets(self.main_frame):
            if isinstance(widget, (ttk.Button, ttk.Entry, ttk.Checkbutton, 
                                  ttk.Radiobutton, ttk.Combobox, tk.Listbox,
                                  ttk.Scale, ttk.Spinbox)):
                if widget != self.cancel_button:
                    self.interactive_widgets.append(widget)
    
    def lock_interface(self):
        """锁定界面"""
        self.is_processing = True
        
        for tab_id in self.notebook.tabs():
            self.notebook.tab(tab_id, state='disabled')
        
        for widget in self.interactive_widgets:
            try:
                widget.config(state='disabled')
            except:
                pass
        
        try:
            self.main_scrollbar.grid_remove()
        except:
            pass
        
        self.unbind_mousewheel_recursive(self.root)
        self.cancel_button.config(state='normal')
    
    def unlock_interface(self):
        """解锁界面"""
        self.is_processing = False
        
        for tab_id in self.notebook.tabs():
            self.notebook.tab(tab_id, state='normal')
        
        for widget in self.interactive_widgets:
            try:
                widget.config(state='normal')
            except:
                pass
        
        try:
            self.main_scrollbar.grid()
        except:
            pass
        
        self.bind_mousewheel_recursive(self.root)
        self.cancel_button.config(state='disabled')
    
    def unbind_mousewheel_recursive(self, widget):
        """递归解绑鼠标滚轮事件"""
        try:
            widget.unbind("<MouseWheel>")
            widget.unbind("<Button-4>")
            widget.unbind("<Button-5>")
        except:
            pass
        
        for child in widget.winfo_children():
            self.unbind_mousewheel_recursive(child)
    
    def get_all_widgets(self, parent):
        """递归获取所有子组件"""
        widgets = []
        try:
            for child in parent.winfo_children():
                widgets.append(child)
                widgets.extend(self.get_all_widgets(child))
        except:
            pass
        return widgets
    
    def convert_dtypes_after_processing(self, df):
        """处理后将文本形式的数字转换为数字类型，但保留日期时间"""
        try:
            for col in df.columns:
                if df[col].dtype in ['int64', 'float64', 'datetime64[ns]']:
                    continue
                
                if self.is_datetime_column(df[col]):
                    continue
                
                try:
                    numeric_values = pd.to_numeric(df[col], errors='coerce')
                    non_null_count = df[col].notna().sum()
                    numeric_count = numeric_values.notna().sum()
                    
                    if numeric_count > non_null_count * 0.8:
                        if (numeric_values.dropna() == numeric_values.dropna().astype(int)).all():
                            df[col] = numeric_values.astype('Int64')
                        else:
                            df[col] = numeric_values
                except:
                    pass
            
            return df
        except Exception as e:
            self.log(f"数据类型转换失败: {str(e)}")
            return df
    
    def is_datetime_column(self, series):
        """检查列是否包含日期时间数据"""
        try:
            sample = series.dropna().head(100)
            if len(sample) == 0:
                return False
            
            if pd.api.types.is_datetime64_any_dtype(series):
                return True
            
            datetime_count = 0
            for value in sample:
                if isinstance(value, (datetime, pd.Timestamp)):
                    datetime_count += 1
                elif isinstance(value, str):
                    date_patterns = [
                        r'\d{4}-\d{2}-\d{2}',
                        r'\d{4}/\d{2}/\d{2}',
                        r'\d{2}-\d{2}-\d{4}',
                        r'\d{2}/\d{2}/\d{4}',
                        r'\d{4}年\d{1,2}月\d{1,2}日',
                    ]
                    for pattern in date_patterns:
                        if re.search(pattern, value):
                            datetime_count += 1
                            break
            
            return datetime_count > len(sample) * 0.5
        except:
            return False
    
    def read_csv_file_robust(self, file_path):
        """读取CSV文件"""
        self.log("读取CSV文件...")
        
        try:
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
            
            if not rows:
                raise Exception("CSV文件为空")
            
            processed_rows = []
            for row in rows:
                if len(row) < max_columns:
                    row = row + [''] * (max_columns - len(row))
                elif len(row) > max_columns:
                    row = row[:max_columns]
                processed_rows.append(row)
            
            if len(processed_rows) > 1:
                columns = processed_rows[0]
                data = processed_rows[1:]
                df = pd.DataFrame(data, columns=columns, dtype=object)
            else:
                df = pd.DataFrame(processed_rows, dtype=object)
            
            return df
            
        except Exception as e:
            self.log(f"CSV读取失败: {str(e)}")
            raise e
    
    def detect_file_encoding(self, file_path):
        """检测文件编码"""
        encodings = ['utf-8', 'utf-8-sig', 'gbk', 'gb2312', 'gb18030', 'big5', 'latin1']
        
        for encoding in encodings:
            try:
                with open(file_path, 'r', encoding=encoding) as f:
                    f.read(1024)
                return encoding
            except:
                continue
        
        return 'latin1'
    
    def detect_delimiter(self, sample_text):
        """检测CSV分隔符"""
        delimiters = [',', ';', '\t', '|']
        
        counts = {}
        for delimiter in delimiters:
            counts[delimiter] = sample_text.count(delimiter)
        
        if counts:
            max_delimiter = max(counts, key=counts.get)
            if counts[max_delimiter] > 0:
                return max_delimiter
        
        return ','
    
    def read_excel_file(self, file_path):
        """智能读取文件"""
        try:
            file_extension = os.path.splitext(file_path)[1].lower()
            
            if file_extension == '.csv':
                return self.read_csv_file_robust(file_path)
            elif file_extension == '.xlsx':
                return pd.read_excel(file_path, engine='openpyxl', dtype=object)
            elif file_extension == '.xls':
                try:
                    return pd.read_excel(file_path, engine='xlrd', dtype=object)
                except:
                    return pd.read_excel(file_path, engine='openpyxl', dtype=object)
            else:
                return pd.read_excel(file_path, dtype=object)
        
        except Exception as e:
            self.log(f"读取文件失败: {str(e)}")
            raise e
    
    def save_excel_file(self, df, file_path):
        """保存文件"""
        try:
            df = self.convert_dtypes_after_processing(df)
            
            file_extension = os.path.splitext(file_path)[1].lower()
            
            if file_extension == '.csv':
                df.to_csv(file_path, index=False, encoding='utf-8-sig')
                return True
            elif file_extension == '.xlsx':
                df.to_excel(file_path, index=False, engine='openpyxl')
                return True
            elif file_extension == '.xls':
                df.to_excel(file_path, index=False, engine='xlwt')
                return True
            else:
                df.to_excel(file_path, index=False, engine='openpyxl')
                return True
        
        except Exception as e:
            self.log(f"保存文件失败: {str(e)}")
            return False
    
    def on_frame_configure(self, event):
        self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))
    
    def on_canvas_configure(self, event):
        self.main_canvas.itemconfig(self.canvas_window, width=event.width)
    
    def bind_mousewheel_recursive(self, widget):
        """绑定鼠标滚轮事件"""
        widget.bind("<MouseWheel>", self.on_mousewheel_windows)
        widget.bind("<Button-4>", self.on_mousewheel_linux)
        widget.bind("<Button-5>", self.on_mousewheel_linux)
        
        for child in widget.winfo_children():
            self.bind_mousewheel_recursive(child)
    
    def on_mousewheel_windows(self, event):
        if self.is_processing:
            return "break"
        
        if event.delta > 0:
            self.main_canvas.yview_scroll(-3, "units")
        else:
            self.main_canvas.yview_scroll(3, "units")
        return "break"
    
    def on_mousewheel_linux(self, event):
        if self.is_processing:
            return "break"
        
        if event.num == 4:
            self.main_canvas.yview_scroll(-3, "units")
        elif event.num == 5:
            self.main_canvas.yview_scroll(3, "units")
        return "break"
    
    def center_window(self):
        self.root.update_idletasks()
        width = 1200
        height = 800
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def create_progress_area(self):
        progress_frame = ttk.LabelFrame(self.main_frame, text="处理进度", padding="5")
        progress_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, 
                                           maximum=100, length=1000, mode='determinate')
        self.progress_bar.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        self.progress_label = ttk.Label(progress_frame, text="0%")
        self.progress_label.grid(row=0, column=2, padx=10)
        
        self.status_label = ttk.Label(progress_frame, text="就绪", foreground="green")
        self.status_label.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=5)
        
        self.indeterminate_progress = ttk.Progressbar(progress_frame, mode='indeterminate',
                                                      length=1000)
        self.indeterminate_progress.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        self.indeterminate_progress.grid_remove()
        
        self.cancel_button = ttk.Button(progress_frame, text="取消", command=self.cancel_operation,
                                       state='disabled')
        self.cancel_button.grid(row=2, column=2, padx=10)
        
        self.cancel_flag = False
        self.last_progress_time = 0
    
    def create_log_area(self):
        log_frame = ttk.LabelFrame(self.main_frame, text="操作日志", padding="5")
        log_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=10)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=5, width=100)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        clear_btn = ttk.Button(log_frame, text="清除日志", command=self.clear_log)
        clear_btn.grid(row=1, column=0, pady=5)
    
    def clear_log(self):
        if self.is_processing:
            return
        self.log_text.delete(1.0, tk.END)
    
    def log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def update_progress(self, value, status_text=None, force_update=False):
        current_time = time.time()
        if not force_update and current_time - self.last_progress_time < 0.1:
            return
        
        self.last_progress_time = current_time
        self.progress_var.set(value)
        self.progress_label.config(text=f"{int(value)}%")
        if status_text:
            self.status_label.config(text=status_text)
        self.root.update_idletasks()
    
    def start_indeterminate(self, status_text="处理中..."):
        self.indeterminate_progress.grid()
        self.indeterminate_progress.start(10)
        self.status_label.config(text=status_text, foreground="blue")
        self.cancel_button.config(state='normal')
        self.cancel_flag = False
        
        self.lock_interface()
        
        self.root.update_idletasks()
    
    def stop_indeterminate(self, status_text="完成", success=True):
        self.indeterminate_progress.stop()
        self.indeterminate_progress.grid_remove()
        if success:
            self.status_label.config(text=status_text, foreground="green")
            self.progress_var.set(100)
            self.progress_label.config(text="100%")
        else:
            self.status_label.config(text=status_text, foreground="red")
        
        self.unlock_interface()
        
        self.root.update_idletasks()
    
    def cancel_operation(self):
        self.cancel_flag = True
        self.status_label.config(text="正在取消...", foreground="orange")
        self.cancel_button.config(state='disabled')
        self.log("用户请求取消操作")
    
    def check_cancel(self):
        if self.cancel_flag:
            raise Exception("操作已被用户取消")
    
    def create_merge_tab(self):
        """创建数据合并标签页 - 支持选择特定列"""
        merge_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(merge_frame, text="数据合并")
        
        # 文件选择区域
        file_frame = ttk.LabelFrame(merge_frame, text="选择要合并的文件（支持Excel和CSV）", padding="10")
        file_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        # 文件列表和滚动条
        file_list_frame = ttk.Frame(file_frame)
        file_list_frame.grid(row=0, column=0, columnspan=3, pady=5, sticky=(tk.W, tk.E))
        
        self.file_listbox = tk.Listbox(file_list_frame, height=4, width=80, selectmode=tk.MULTIPLE)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        file_scrollbar = ttk.Scrollbar(file_list_frame, orient="vertical", command=self.file_listbox.yview)
        file_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_listbox.configure(yscrollcommand=file_scrollbar.set)
        
        # 文件操作按钮
        button_frame = ttk.Frame(file_frame)
        button_frame.grid(row=1, column=0, columnspan=3, pady=5)
        
        ttk.Button(button_frame, text="添加文件", command=self.add_files).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="移除选中", command=self.remove_selected_files).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="清空列表", command=self.clear_file_list).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="预览列名", command=self.preview_columns).pack(side=tk.LEFT, padx=5)
        
        # 列选择区域
        column_frame = ttk.LabelFrame(merge_frame, text="列选择（留空表示选择所有列）", padding="10")
        column_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        # 按列名选择
        ttk.Label(column_frame, text="按列名选择:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.columns_by_name = ttk.Entry(column_frame, width=60)
        self.columns_by_name.grid(row=0, column=1, columnspan=2, sticky=tk.W, pady=5)
        ttk.Label(column_frame, text="(用逗号分隔列名，如: 姓名,年龄,工资)").grid(row=0, column=3, sticky=tk.W)
        
        # 按列坐标选择
        ttk.Label(column_frame, text="按列坐标选择:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.columns_by_index = ttk.Entry(column_frame, width=60)
        self.columns_by_index.grid(row=1, column=1, columnspan=2, sticky=tk.W, pady=5)
        ttk.Label(column_frame, text="(用逗号分隔列号，从1开始，如: 1,3,5 或 1-5)").grid(row=1, column=3, sticky=tk.W)
        
        # 合并选项
        option_frame = ttk.LabelFrame(merge_frame, text="合并选项", padding="10")
        option_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        # 合并方式
        ttk.Label(option_frame, text="合并方式:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.merge_type = tk.StringVar(value="vertical")
        ttk.Radiobutton(option_frame, text="垂直合并（按行追加）", variable=self.merge_type, 
                       value="vertical").grid(row=0, column=1, sticky=tk.W, padx=10)
        ttk.Radiobutton(option_frame, text="水平合并（按列拼接）", variable=self.merge_type, 
                       value="horizontal").grid(row=0, column=2, sticky=tk.W, padx=10)
        
        self.add_source_var = tk.BooleanVar()
        ttk.Checkbutton(option_frame, text="添加数据来源标识", 
                       variable=self.add_source_var).grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        self.remove_dup_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(option_frame, text="去除重复行", 
                       variable=self.remove_dup_var).grid(row=1, column=2, columnspan=2, sticky=tk.W, pady=5)
        
        # 输出设置
        output_frame = ttk.LabelFrame(merge_frame, text="输出设置", padding="10")
        output_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(output_frame, text="输出文件:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.output_path = tk.StringVar()
        ttk.Entry(output_frame, textvariable=self.output_path, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(output_frame, text="浏览", command=self.select_output_file).grid(row=0, column=2)
        
        # 执行按钮
        self.merge_button = ttk.Button(merge_frame, text="开始合并", command=self.start_merge, width=20)
        self.merge_button.grid(row=4, column=0, columnspan=3, pady=10)
    
    def create_convert_tab(self):
        """创建格式转换标签页"""
        convert_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(convert_frame, text="格式转换")
        
        # 转换类型选择
        type_frame = ttk.LabelFrame(convert_frame, text="转换类型", padding="10")
        type_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        self.convert_type = tk.StringVar(value="excel_to_word")
        
        # 第一行
        ttk.Radiobutton(type_frame, text="Excel/CSV转Word", variable=self.convert_type, 
                       value="excel_to_word", command=self.update_file_types).grid(row=0, column=0, padx=10, pady=5)
        ttk.Radiobutton(type_frame, text="Word转Excel", variable=self.convert_type, 
                       value="word_to_excel", command=self.update_file_types).grid(row=0, column=1, padx=10, pady=5)
        ttk.Radiobutton(type_frame, text="Excel/CSV转PDF", variable=self.convert_type, 
                       value="excel_to_pdf", command=self.update_file_types).grid(row=0, column=2, padx=10, pady=5)
        
        # 第二行
        ttk.Radiobutton(type_frame, text="Excel/CSV转PPT", variable=self.convert_type, 
                       value="excel_to_ppt", command=self.update_file_types).grid(row=1, column=0, padx=10, pady=5)
        ttk.Radiobutton(type_frame, text="Word转PDF", variable=self.convert_type, 
                       value="word_to_pdf", command=self.update_file_types).grid(row=1, column=1, padx=10, pady=5)
        ttk.Radiobutton(type_frame, text="PPT转PDF", variable=self.convert_type, 
                       value="ppt_to_pdf", command=self.update_file_types).grid(row=1, column=2, padx=10, pady=5)
        
        # 输入文件
        input_frame = ttk.LabelFrame(convert_frame, text="输入文件", padding="10")
        input_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        self.convert_input = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.convert_input, width=60).grid(row=0, column=0, padx=5)
        ttk.Button(input_frame, text="浏览", command=self.select_input_file).grid(row=0, column=1)
        
        # 输出文件
        output_frame = ttk.LabelFrame(convert_frame, text="输出文件", padding="10")
        output_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        self.convert_output = tk.StringVar()
        ttk.Entry(output_frame, textvariable=self.convert_output, width=60).grid(row=0, column=0, padx=5)
        ttk.Button(output_frame, text="浏览", command=self.select_convert_output).grid(row=0, column=1)
        
        # 执行按钮
        self.convert_button = ttk.Button(convert_frame, text="开始转换", command=self.start_convert, width=20)
        self.convert_button.grid(row=3, column=0, columnspan=3, pady=10)
    
    def create_split_tab(self):
        """创建数据拆分标签页 - 支持按多列或多行拆分"""
        split_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(split_frame, text="数据拆分")
        
        # 文件选择
        file_frame = ttk.LabelFrame(split_frame, text="选择文件", padding="10")
        file_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        self.split_file_path = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.split_file_path, width=60).grid(row=0, column=0, padx=5)
        ttk.Button(file_frame, text="浏览", command=self.select_split_file).grid(row=0, column=1)
        ttk.Button(file_frame, text="预览列名", command=self.preview_split_columns).grid(row=0, column=2, padx=5)
        
        # 拆分方式选择
        method_frame = ttk.LabelFrame(split_frame, text="拆分方式", padding="10")
        method_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        self.split_method = tk.StringVar(value="by_columns")
        ttk.Radiobutton(method_frame, text="按列值拆分", variable=self.split_method, 
                       value="by_columns", command=self.toggle_split_method).grid(row=0, column=0, padx=10)
        ttk.Radiobutton(method_frame, text="按行数拆分", variable=self.split_method, 
                       value="by_rows", command=self.toggle_split_method).grid(row=0, column=1, padx=10)
        ttk.Radiobutton(method_frame, text="按特定行拆分", variable=self.split_method, 
                       value="by_specific_rows", command=self.toggle_split_method).grid(row=0, column=2, padx=10)
        
        # 按列拆分设置
        self.column_split_frame = ttk.LabelFrame(split_frame, text="按列拆分设置", padding="10")
        self.column_split_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(self.column_split_frame, text="拆分依据列:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.split_columns = ttk.Entry(self.column_split_frame, width=50)
        self.split_columns.grid(row=0, column=1, sticky=tk.W, pady=5)
        ttk.Label(self.column_split_frame, text="(多个列用逗号分隔，如: 部门,地区)").grid(row=0, column=2, sticky=tk.W)
        
        # 按行数拆分设置
        self.row_split_frame = ttk.LabelFrame(split_frame, text="按行数拆分设置", padding="10")
        self.row_split_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(self.row_split_frame, text="每个文件行数:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.rows_per_file = ttk.Entry(self.row_split_frame, width=20)
        self.rows_per_file.grid(row=0, column=1, sticky=tk.W, pady=5)
        self.rows_per_file.insert(0, "1000")
        
        # 按特定行拆分设置
        self.specific_row_frame = ttk.LabelFrame(split_frame, text="按特定行拆分设置", padding="10")
        self.specific_row_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(self.specific_row_frame, text="拆分行号:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.specific_rows = ttk.Entry(self.specific_row_frame, width=50)
        self.specific_rows.grid(row=0, column=1, sticky=tk.W, pady=5)
        ttk.Label(self.specific_row_frame, text="(用逗号分隔行号，如: 100,200,300 表示在这些行处拆分)").grid(row=0, column=2, sticky=tk.W)
        
        # 输出设置
        output_frame = ttk.LabelFrame(split_frame, text="输出设置", padding="10")
        output_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(output_frame, text="输出目录:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.split_output_dir = tk.StringVar(value="./split_output")
        ttk.Entry(output_frame, textvariable=self.split_output_dir, width=40).grid(row=0, column=1, pady=5)
        ttk.Button(output_frame, text="浏览", command=self.select_split_output).grid(row=0, column=2, padx=5)
        
        # 执行按钮
        self.split_button = ttk.Button(split_frame, text="开始拆分", command=self.start_split, width=20)
        self.split_button.grid(row=6, column=0, columnspan=3, pady=10)
        
        # 初始隐藏行数拆分设置
        self.row_split_frame.grid_remove()
        self.specific_row_frame.grid_remove()
    
    def create_batch_tab(self):
        """创建批量处理标签页"""
        batch_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(batch_frame, text="批量处理")
        
        dir_frame = ttk.LabelFrame(batch_frame, text="选择文件夹", padding="10")
        dir_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        self.batch_dir = tk.StringVar()
        ttk.Entry(dir_frame, textvariable=self.batch_dir, width=60).grid(row=0, column=0, padx=5)
        ttk.Button(dir_frame, text="浏览", command=self.select_batch_dir).grid(row=0, column=1)
        
        option_frame = ttk.LabelFrame(batch_frame, text="处理选项", padding="10")
        option_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        self.batch_operation = tk.StringVar(value="clean")
        ttk.Radiobutton(option_frame, text="清理数据", variable=self.batch_operation, 
                       value="clean").grid(row=0, column=0, padx=10)
        ttk.Radiobutton(option_frame, text="去除空行", variable=self.batch_operation, 
                       value="remove_empty").grid(row=0, column=1, padx=10)
        
        self.batch_button = ttk.Button(batch_frame, text="开始批量处理", command=self.start_batch, width=20)
        self.batch_button.grid(row=2, column=0, columnspan=2, pady=10)
    
    def toggle_split_method(self):
        """切换拆分方式"""
        method = self.split_method.get()
        
        # 隐藏所有设置
        self.column_split_frame.grid_remove()
        self.row_split_frame.grid_remove()
        self.specific_row_frame.grid_remove()
        
        # 显示选中的设置
        if method == "by_columns":
            self.column_split_frame.grid()
        elif method == "by_rows":
            self.row_split_frame.grid()
        elif method == "by_specific_rows":
            self.specific_row_frame.grid()
    
    def preview_columns(self):
        """预览文件列名"""
        if self.is_processing:
            return
        
        files = list(self.file_listbox.get(0, tk.END))
        if not files:
            messagebox.showwarning("警告", "请先添加文件")
            return
        
        try:
            df = self.read_excel_file(files[0])
            columns = df.columns.tolist()
            
            # 创建预览窗口
            preview_window = tk.Toplevel(self.root)
            preview_window.title("列名预览")
            preview_window.geometry("400x500")
            
            # 显示列名
            columns_text = scrolledtext.ScrolledText(preview_window, width=50, height=20)
            columns_text.pack(padx=10, pady=10)
            
            for i, col in enumerate(columns, 1):
                columns_text.insert(tk.END, f"{i}. {col}\n")
            
            columns_text.config(state='disabled')
            
            ttk.Button(preview_window, text="关闭", command=preview_window.destroy).pack(pady=10)
            
        except Exception as e:
            messagebox.showerror("错误", f"读取文件失败: {str(e)}")
    
    def preview_split_columns(self):
        """预览拆分文件的列名"""
        if self.is_processing:
            return
        
        file_path = self.split_file_path.get()
        if not file_path:
            messagebox.showwarning("警告", "请先选择文件")
            return
        
        try:
            df = self.read_excel_file(file_path)
            columns = df.columns.tolist()
            
            preview_window = tk.Toplevel(self.root)
            preview_window.title("列名预览")
            preview_window.geometry("400x500")
            
            columns_text = scrolledtext.ScrolledText(preview_window, width=50, height=20)
            columns_text.pack(padx=10, pady=10)
            
            for i, col in enumerate(columns, 1):
                columns_text.insert(tk.END, f"{i}. {col}\n")
            
            columns_text.config(state='disabled')
            
            ttk.Button(preview_window, text="关闭", command=preview_window.destroy).pack(pady=10)
            
        except Exception as e:
            messagebox.showerror("错误", f"读取文件失败: {str(e)}")
    
    def parse_column_selection(self, columns_str, df_columns):
        """解析列选择字符串"""
        selected_columns = []
        
        # 按列名选择
        if self.columns_by_name.get().strip():
            column_names = [c.strip() for c in self.columns_by_name.get().split(',') if c.strip()]
            for col_name in column_names:
                if col_name in df_columns:
                    selected_columns.append(col_name)
                else:
                    self.log(f"警告: 列 '{col_name}' 不存在")
        
        # 按列坐标选择
        if self.columns_by_index.get().strip():
            index_str = self.columns_by_index.get().strip()
            indices = []
            
            for part in index_str.split(','):
                part = part.strip()
                if '-' in part:
                    # 范围选择，如 1-5
                    start, end = part.split('-')
                    indices.extend(range(int(start), int(end) + 1))
                else:
                    indices.append(int(part))
            
            for idx in indices:
                if 1 <= idx <= len(df_columns):
                    selected_columns.append(df_columns[idx - 1])
        
        # 去重并保持顺序
        selected_columns = list(dict.fromkeys(selected_columns))
        
        return selected_columns
    
    def add_files(self):
        if self.is_processing:
            return
        
        files = filedialog.askopenfilenames(
            title="选择文件",
            filetypes=[
                ("所有支持的文件", "*.xlsx *.xls *.csv"),
                ("Excel文件", "*.xlsx *.xls"),
                ("CSV文件", "*.csv"),
                ("所有文件", "*.*")
            ]
        )
        for file in files:
            if file not in self.file_listbox.get(0, tk.END):
                self.file_listbox.insert(tk.END, file)
                self.log(f"添加文件: {os.path.basename(file)}")
    
    def remove_selected_files(self):
        if self.is_processing:
            return
        
        selected = self.file_listbox.curselection()
        for index in reversed(selected):
            self.file_listbox.delete(index)
            self.log("移除文件")
    
    def clear_file_list(self):
        if self.is_processing:
            return
        
        self.file_listbox.delete(0, tk.END)
        self.log("清空文件列表")
    
    def select_output_file(self):
        if self.is_processing:
            return
        
        file_path = filedialog.asksaveasfilename(
            title="保存文件",
            defaultextension=".xlsx",
            filetypes=[
                ("Excel文件", "*.xlsx"),
                ("CSV文件", "*.csv"),
                ("所有文件", "*.*")
            ]
        )
        if file_path:
            self.output_path.set(file_path)
            self.log(f"输出文件设置为: {file_path}")
    
    def select_input_file(self):
        if self.is_processing:
            return
        
        convert_type = self.convert_type.get()
        
        if convert_type in ["excel_to_word", "excel_to_pdf", "excel_to_ppt"]:
            file_path = filedialog.askopenfilename(
                title="选择文件",
                filetypes=[
                    ("所有支持的文件", "*.xlsx *.xls *.csv"),
                    ("Excel文件", "*.xlsx *.xls"),
                    ("CSV文件", "*.csv")
                ]
            )
        elif convert_type in ["word_to_excel", "word_to_pdf"]:
            file_path = filedialog.askopenfilename(
                title="选择Word文件",
                filetypes=[("Word文件", "*.docx *.doc")]
            )
        elif convert_type == "ppt_to_pdf":
            file_path = filedialog.askopenfilename(
                title="选择PPT文件",
                filetypes=[("PPT文件", "*.pptx *.ppt")]
            )
        else:
            file_path = filedialog.askopenfilename(
                title="选择文件",
                filetypes=[("所有文件", "*.*")]
            )
        
        if file_path:
            self.convert_input.set(file_path)
    
    def select_convert_output(self):
        if self.is_processing:
            return
        
        convert_type = self.convert_type.get()
        
        if convert_type == "excel_to_word":
            file_path = filedialog.asksaveasfilename(
                title="保存Word文件",
                defaultextension=".docx",
                filetypes=[("Word文件", "*.docx")]
            )
        elif convert_type == "word_to_excel":
            file_path = filedialog.asksaveasfilename(
                title="保存文件",
                defaultextension=".xlsx",
                filetypes=[
                    ("Excel文件", "*.xlsx"),
                    ("CSV文件", "*.csv")
                ]
            )
        elif convert_type in ["excel_to_pdf", "word_to_pdf", "ppt_to_pdf"]:
            file_path = filedialog.asksaveasfilename(
                title="保存PDF文件",
                defaultextension=".pdf",
                filetypes=[("PDF文件", "*.pdf")]
            )
        elif convert_type == "excel_to_ppt":
            file_path = filedialog.asksaveasfilename(
                title="保存PPT文件",
                defaultextension=".pptx",
                filetypes=[("PPT文件", "*.pptx")]
            )
        else:
            file_path = filedialog.asksaveasfilename(
                title="保存文件",
                filetypes=[("所有文件", "*.*")]
            )
        
        if file_path:
            self.convert_output.set(file_path)
    
    def select_split_file(self):
        if self.is_processing:
            return
        
        file_path = filedialog.askopenfilename(
            title="选择文件",
            filetypes=[
                ("所有支持的文件", "*.xlsx *.xls *.csv"),
                ("Excel文件", "*.xlsx *.xls"),
                ("CSV文件", "*.csv")
            ]
        )
        if file_path:
            self.split_file_path.set(file_path)
    
    def select_split_output(self):
        if self.is_processing:
            return
        
        dir_path = filedialog.askdirectory(title="选择输出目录")
        if dir_path:
            self.split_output_dir.set(dir_path)
    
    def select_batch_dir(self):
        if self.is_processing:
            return
        
        dir_path = filedialog.askdirectory(title="选择文件夹")
        if dir_path:
            self.batch_dir.set(dir_path)
    
    def update_file_types(self):
        """更新文件类型"""
        pass
    
    # 处理功能方法
    def start_merge(self):
        if self.is_processing:
            return
        
        files = list(self.file_listbox.get(0, tk.END))
        if not files:
            messagebox.showwarning("警告", "请先添加要合并的文件")
            return
        
        output_path = self.output_path.get()
        if not output_path:
            messagebox.showwarning("警告", "请指定输出文件路径")
            return
        
        thread = threading.Thread(target=self.merge_thread, args=(files, output_path))
        thread.daemon = True
        thread.start()
    
    def merge_thread(self, files, output_path):
        try:
            self.start_indeterminate("正在合并文件...")
            self.log("开始合并文件...")
            
            dfs = []
            total_files = len(files)
            
            # 获取第一个文件的列名，用于列选择
            first_df = self.read_excel_file(files[0])
            all_columns = first_df.columns.tolist()
            
            # 解析列选择
            selected_columns = self.parse_column_selection(
                self.columns_by_name.get(), all_columns
            )
            
            if selected_columns:
                self.log(f"选择的列: {selected_columns}")
            else:
                self.log("选择所有列")
            
            for i, file in enumerate(files):
                self.check_cancel()
                
                if i == 0 or i == total_files - 1 or i % 5 == 0:
                    self.log(f"读取文件 {i+1}/{total_files}: {os.path.basename(file)}")
                
                self.update_progress((i / total_files) * 50, f"读取文件 {i+1}/{total_files}")
                
                df = self.read_excel_file(file)
                
                # 选择特定列
                if selected_columns:
                    # 检查列是否存在
                    available_columns = [col for col in selected_columns if col in df.columns]
                    if len(available_columns) != len(selected_columns):
                        missing = [col for col in selected_columns if col not in df.columns]
                        self.log(f"警告: 文件 {os.path.basename(file)} 缺少列: {missing}")
                    df = df[available_columns]
                
                # 添加数据来源列
                if self.add_source_var.get():
                    df['数据来源'] = os.path.basename(file)
                
                dfs.append(df)
            
            self.check_cancel()
            self.update_progress(60, "正在合并数据...", force_update=True)
            
            merge_type = self.merge_type.get()
            if merge_type == "vertical":
                merged_df = pd.concat(dfs, ignore_index=True)
            else:
                merged_df = pd.concat(dfs, axis=1)
            
            self.update_progress(80, "正在处理数据...", force_update=True)
            
            if self.remove_dup_var.get():
                self.check_cancel()
                merged_df = merged_df.drop_duplicates()
            
            self.check_cancel()
            self.update_progress(90, "正在保存结果...", force_update=True)
            
            if self.save_excel_file(merged_df, output_path):
                self.update_progress(100, "合并完成", force_update=True)
                self.log(f"合并完成！结果已保存到: {output_path}")
                self.log(f"合并后数据: {len(merged_df)}行, {len(merged_df.columns)}列")
                
                self.stop_indeterminate("合并完成", success=True)
                messagebox.showinfo("成功", f"合并完成！\n输出文件: {output_path}\n共处理 {len(merged_df)} 行数据")
            else:
                raise Exception("保存文件失败")
            
        except Exception as e:
            if "取消" in str(e):
                self.stop_indeterminate("已取消", success=False)
                self.log("操作已被用户取消")
            else:
                self.stop_indeterminate("处理失败", success=False)
                self.log(f"合并失败: {str(e)}")
                messagebox.showerror("错误", f"合并失败: {str(e)}")
    
    def start_convert(self):
        if self.is_processing:
            return
        
        input_path = self.convert_input.get()
        output_path = self.convert_output.get()
        
        if not input_path:
            messagebox.showwarning("警告", "请选择输入文件")
            return
        if not output_path:
            messagebox.showwarning("警告", "请指定输出文件路径")
            return
        
        thread = threading.Thread(target=self.convert_thread, args=(input_path, output_path))
        thread.daemon = True
        thread.start()
    
    def convert_thread(self, input_path, output_path):
        try:
            convert_type = self.convert_type.get()
            self.start_indeterminate("正在转换...")
            self.log(f"开始转换: {convert_type}")
            
            if convert_type == "excel_to_word":
                self.convert_excel_to_word(input_path, output_path)
            elif convert_type == "word_to_excel":
                self.convert_word_to_excel(input_path, output_path)
            elif convert_type == "excel_to_pdf":
                self.convert_excel_to_pdf(input_path, output_path)
            elif convert_type == "excel_to_ppt":
                self.convert_excel_to_ppt(input_path, output_path)
            elif convert_type == "word_to_pdf":
                self.convert_word_to_pdf(input_path, output_path)
            elif convert_type == "ppt_to_pdf":
                self.convert_ppt_to_pdf(input_path, output_path)
            
            self.update_progress(100, "转换完成", force_update=True)
            self.stop_indeterminate("转换完成", success=True)
            messagebox.showinfo("成功", f"转换完成！\n输出文件: {output_path}")
            
        except Exception as e:
            if "取消" in str(e):
                self.stop_indeterminate("已取消", success=False)
                self.log("操作已被用户取消")
            else:
                self.stop_indeterminate("转换失败", success=False)
                self.log(f"转换失败: {str(e)}")
                messagebox.showerror("错误", f"转换失败: {str(e)}")
    
    def convert_excel_to_word(self, input_path, output_path):
        """Excel转Word"""
        from docx import Document
        
        self.update_progress(20, "正在读取文件...", force_update=True)
        df = self.read_excel_file(input_path)
        
        self.check_cancel()
        self.update_progress(50, "正在创建Word文档...", force_update=True)
        doc = Document()
        doc.add_heading('数据转换结果', level=1)
        
        table = doc.add_table(rows=1, cols=len(df.columns))
        table.style = 'Light Grid Accent 1'
        
        header_cells = table.rows[0].cells
        for i, col in enumerate(df.columns):
            header_cells[i].text = str(col)
        
        total_rows = len(df)
        batch_size = max(1, total_rows // 10)
        
        for idx, (_, row) in enumerate(df.iterrows()):
            self.check_cancel()
            
            row_cells = table.add_row().cells
            for i, value in enumerate(row):
                if pd.isna(value):
                    row_cells[i].text = ''
                else:
                    row_cells[i].text = str(value)
            
            if idx % batch_size == 0:
                progress = 50 + (idx / total_rows) * 40
                self.update_progress(progress, f"正在转换数据 {idx}/{total_rows}")
        
        self.update_progress(90, "正在保存Word文件...", force_update=True)
        doc.save(output_path)
        self.log(f"Excel转Word完成: {output_path}")
    
    def convert_word_to_excel(self, input_path, output_path):
        """Word转Excel"""
        from docx import Document
        
        self.update_progress(30, "正在读取Word文件...", force_update=True)
        doc = Document(input_path)
        
        if not doc.tables:
            raise Exception("Word文档中没有表格")
        
        self.check_cancel()
        self.update_progress(60, "正在提取表格数据...", force_update=True)
        table = doc.tables[0]
        data = []
        for row in table.rows:
            row_data = [cell.text for cell in row.cells]
            data.append(row_data)
        
        self.update_progress(80, "正在转换为Excel...", force_update=True)
        df = pd.DataFrame(data[1:], columns=data[0])
        
        if self.save_excel_file(df, output_path):
            self.log(f"Word转Excel完成: {output_path}")
        else:
            raise Exception("保存文件失败")
    
    def convert_excel_to_pdf(self, input_path, output_path):
        """Excel转PDF"""
        try:
            from reportlab.lib.pagesizes import A4, landscape
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
            from reportlab.lib import colors
            
            self.update_progress(20, "正在读取Excel文件...", force_update=True)
            df = self.read_excel_file(input_path)
            
            self.check_cancel()
            self.update_progress(50, "正在创建PDF...", force_update=True)
            
            doc = SimpleDocTemplate(output_path, pagesize=landscape(A4))
            elements = []
            
            table_data = [df.columns.tolist()] + df.values.tolist()
            table = Table(table_data)
            
            style = TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 12),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('FONTSIZE', (0, 1), (-1, -1), 8),
            ])
            table.setStyle(style)
            
            elements.append(table)
            
            self.check_cancel()
            self.update_progress(80, "正在保存PDF...", force_update=True)
            
            doc.build(elements)
            self.log(f"Excel转PDF完成: {output_path}")
            
        except ImportError:
            self.log("缺少必要的库，请安装: pip install reportlab openpyxl")
            raise Exception("缺少必要的库: reportlab")
    
    def convert_excel_to_ppt(self, input_path, output_path):
        """Excel转PPT"""
        try:
            from pptx import Presentation
            from pptx.util import Inches, Pt
            
            self.update_progress(20, "正在读取Excel文件...", force_update=True)
            df = self.read_excel_file(input_path)
            
            self.check_cancel()
            self.update_progress(50, "正在创建PPT...", force_update=True)
            
            prs = Presentation()
            
            title_slide_layout = prs.slide_layouts[0]
            slide = prs.slides.add_slide(title_slide_layout)
            title = slide.shapes.title
            subtitle = slide.placeholders[1]
            title.text = "数据展示"
            subtitle.text = f"数据行数: {len(df)}"
            
            table_slide_layout = prs.slide_layouts[5]
            slide = prs.slides.add_slide(table_slide_layout)
            title = slide.shapes.title
            title.text = "数据详情"
            
            rows, cols = min(len(df) + 1, 20), min(len(df.columns), 8)
            table_shape = slide.shapes.add_table(rows, cols, 
                                                 Inches(0.5), Inches(1.5),
                                                 Inches(9), Inches(5))
            table = table_shape.table
            
            for i, col in enumerate(df.columns[:cols]):
                table.cell(0, i).text = str(col)
            
            for i in range(1, rows):
                for j in range(cols):
                    value = df.iloc[i-1, j]
                    table.cell(i, j).text = str(value)
            
            self.check_cancel()
            self.update_progress(80, "正在保存PPT...", force_update=True)
            
            prs.save(output_path)
            self.log(f"Excel转PPT完成: {output_path}")
            
        except ImportError:
            self.log("缺少必要的库，请安装: pip install python-pptx")
            raise Exception("缺少必要的库: python-pptx")
    
    def convert_word_to_pdf(self, input_path, output_path):
        """Word转PDF"""
        try:
            try:
                from docx2pdf import convert
                self.update_progress(50, "正在转换...", force_update=True)
                convert(input_path, output_path)
                self.log(f"Word转PDF完成: {output_path}")
                return
            except ImportError:
                pass
            
            from docx import Document
            from reportlab.lib.pagesizes import A4
            from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
            from reportlab.lib.styles import getSampleStyleSheet
            
            self.update_progress(30, "正在读取Word文件...", force_update=True)
            doc = Document(input_path)
            
            self.check_cancel()
            self.update_progress(60, "正在创建PDF...", force_update=True)
            
            pdf_doc = SimpleDocTemplate(output_path, pagesize=A4)
            styles = getSampleStyleSheet()
            elements = []
            
            for para in doc.paragraphs:
                if para.text.strip():
                    elements.append(Paragraph(para.text, styles['Normal']))
                    elements.append(Spacer(1, 12))
            
            self.update_progress(80, "正在保存PDF...", force_update=True)
            pdf_doc.build(elements)
            self.log(f"Word转PDF完成: {output_path}")
            
        except ImportError:
            self.log("缺少必要的库，请安装: pip install docx2pdf reportlab")
            raise Exception("缺少必要的库: docx2pdf 或 reportlab")
    
    def convert_ppt_to_pdf(self, input_path, output_path):
        """PPT转PDF"""
        try:
            try:
                import win32com.client
                
                self.update_progress(50, "正在转换...", force_update=True)
                
                powerpoint = win32com.client.Dispatch("PowerPoint.Application")
                presentation = powerpoint.Presentations.Open(input_path)
                presentation.SaveAs(output_path, 32)
                presentation.Close()
                powerpoint.Quit()
                
                self.log(f"PPT转PDF完成: {output_path}")
                return
                
            except ImportError:
                pass
            
            from pptx import Presentation
            from reportlab.lib.pagesizes import A4, landscape
            from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
            from reportlab.lib.styles import getSampleStyleSheet
            
            self.update_progress(30, "正在读取PPT文件...", force_update=True)
            prs = Presentation(input_path)
            
            self.check_cancel()
            self.update_progress(60, "正在创建PDF...", force_update=True)
            
            pdf_doc = SimpleDocTemplate(output_path, pagesize=landscape(A4))
            styles = getSampleStyleSheet()
            elements = []
            
            for slide in prs.slides:
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        for para in shape.text_frame.paragraphs:
                            if para.text.strip():
                                elements.append(Paragraph(para.text, styles['Normal']))
                                elements.append(Spacer(1, 12))
                
                elements.append(Spacer(1, 30))
            
            self.update_progress(80, "正在保存PDF...", force_update=True)
            pdf_doc.build(elements)
            self.log(f"PPT转PDF完成: {output_path}")
            
        except ImportError:
            self.log("缺少必要的库，请安装: pip install pywin32 python-pptx reportlab")
            raise Exception("缺少必要的库: pywin32 或 python-pptx")
    
    def start_split(self):
        if self.is_processing:
            return
        
        file_path = self.split_file_path.get()
        if not file_path:
            messagebox.showwarning("警告", "请选择要拆分的文件")
            return
        
        thread = threading.Thread(target=self.split_thread, args=(file_path,))
        thread.daemon = True
        thread.start()
    
    def split_thread(self, file_path):
        try:
            self.start_indeterminate("正在拆分数据...")
            
            self.update_progress(10, "正在读取文件...", force_update=True)
            df = self.read_excel_file(file_path)
            
            output_dir = self.split_output_dir.get()
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            split_method = self.split_method.get()
            
            if split_method == "by_columns":
                self.split_by_columns(df, output_dir)
            elif split_method == "by_rows":
                self.split_by_rows(df, output_dir)
            elif split_method == "by_specific_rows":
                self.split_by_specific_rows(df, output_dir)
            
            self.update_progress(100, "拆分完成", force_update=True)
            self.log("拆分完成！")
            
            self.stop_indeterminate("拆分完成", success=True)
            messagebox.showinfo("成功", f"拆分完成！\n输出目录: {output_dir}")
            
        except Exception as e:
            if "取消" in str(e):
                self.stop_indeterminate("已取消", success=False)
                self.log("操作已被用户取消")
            else:
                self.stop_indeterminate("拆分失败", success=False)
                self.log(f"拆分失败: {str(e)}")
                messagebox.showerror("错误", f"拆分失败: {str(e)}")
    
    def split_by_columns(self, df, output_dir):
        """按列值拆分"""
        columns_str = self.split_columns.get().strip()
        if not columns_str:
            raise Exception("请输入拆分依据列")
        
        split_columns = [c.strip() for c in columns_str.split(',') if c.strip()]
        
        # 检查列是否存在
        missing_columns = [col for col in split_columns if col not in df.columns]
        if missing_columns:
            raise Exception(f"列不存在: {missing_columns}")
        
        self.log(f"按列拆分: {split_columns}")
        
        # 按多列分组
        groups = df.groupby(split_columns)
        total_groups = len(groups)
        
        for i, (name, group) in enumerate(groups):
            self.check_cancel()
            
            # 生成文件名
            if isinstance(name, tuple):
                safe_name = '_'.join([re.sub(r'[\\/*?:"<>|]', '_', str(n)) for n in name])
            else:
                safe_name = re.sub(r'[\\/*?:"<>|]', '_', str(name))
            
            output_path = os.path.join(output_dir, f"{safe_name}.xlsx")
            
            if self.save_excel_file(group, output_path):
                if i % max(1, total_groups // 20) == 0:
                    progress = 20 + ((i + 1) / total_groups) * 70
                    self.update_progress(progress, f"正在保存: {safe_name}.xlsx ({i+1}/{total_groups})")
                    self.log(f"已保存: {safe_name}.xlsx ({len(group)}行)")
        
        self.log(f"按列拆分完成！共生成 {total_groups} 个文件")
    
    def split_by_rows(self, df, output_dir):
        """按行数拆分"""
        try:
            rows_per_file = int(self.rows_per_file.get())
            if rows_per_file <= 0:
                raise ValueError("每个文件行数必须大于0")
        except ValueError as e:
            raise Exception(f"无效的行数: {e}")
        
        total_rows = len(df)
        total_files = (total_rows + rows_per_file - 1) // rows_per_file
        
        self.log(f"按行数拆分: 每{rows_per_file}行一个文件，共{total_files}个文件")
        
        for i in range(total_files):
            self.check_cancel()
            
            start_idx = i * rows_per_file
            end_idx = min((i + 1) * rows_per_file, total_rows)
            
            chunk = df.iloc[start_idx:end_idx]
            
            output_path = os.path.join(output_dir, f"part_{i+1:03d}.xlsx")
            
            if self.save_excel_file(chunk, output_path):
                progress = 20 + ((i + 1) / total_files) * 70
                self.update_progress(progress, f"正在保存: part_{i+1:03d}.xlsx ({i+1}/{total_files})")
                self.log(f"已保存: part_{i+1:03d}.xlsx ({len(chunk)}行)")
        
        self.log(f"按行数拆分完成！共生成 {total_files} 个文件")
    
    def split_by_specific_rows(self, df, output_dir):
        """按特定行拆分"""
        rows_str = self.specific_rows.get().strip()
        if not rows_str:
            raise Exception("请输入拆分行号")
        
        try:
            split_rows = [int(r.strip()) for r in rows_str.split(',') if r.strip()]
            split_rows.sort()
        except ValueError:
            raise Exception("无效的行号格式")
        
        if not split_rows:
            raise Exception("请输入有效的行号")
        
        self.log(f"按特定行拆分: {split_rows}")
        
        # 添加起始和结束行
        all_split_points = [0] + split_rows + [len(df)]
        
        total_files = len(split_rows) + 1
        
        for i in range(total_files):
            self.check_cancel()
            
            start_idx = all_split_points[i]
            end_idx = all_split_points[i + 1]
            
            chunk = df.iloc[start_idx:end_idx]
            
            output_path = os.path.join(output_dir, f"part_{i+1:03d}.xlsx")
            
            if self.save_excel_file(chunk, output_path):
                progress = 20 + ((i + 1) / total_files) * 70
                self.update_progress(progress, f"正在保存: part_{i+1:03d}.xlsx ({i+1}/{total_files})")
                self.log(f"已保存: part_{i+1:03d}.xlsx ({len(chunk)}行)")
        
        self.log(f"按特定行拆分完成！共生成 {total_files} 个文件")
    
    def start_batch(self):
        if self.is_processing:
            return
        
        directory = self.batch_dir.get()
        if not directory:
            messagebox.showwarning("警告", "请选择要处理的文件夹")
            return
        
        thread = threading.Thread(target=self.batch_thread, args=(directory,))
        thread.daemon = True
        thread.start()
    
    def batch_thread(self, directory):
        try:
            self.start_indeterminate("正在批量处理...")
            
            excel_files = [f for f in os.listdir(directory) 
                          if f.endswith(('.xlsx', '.xls', '.csv'))]
            
            if not excel_files:
                self.log("目录中没有支持的文件")
                self.stop_indeterminate("未找到文件", success=False)
                messagebox.showwarning("警告", "目录中没有支持的文件（.xlsx, .xls, .csv）")
                return
            
            output_dir = os.path.join(directory, 'processed')
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            operation = self.batch_operation.get()
            total_files = len(excel_files)
            
            for i, file_name in enumerate(excel_files):
                self.check_cancel()
                file_path = os.path.join(directory, file_name)
                
                if i % 5 == 0 or i == total_files - 1:
                    self.log(f"处理文件 {i+1}/{total_files}: {file_name}")
                
                progress = (i / total_files) * 100
                self.update_progress(progress, f"正在处理 {i+1}/{total_files}")
                
                df = self.read_excel_file(file_path)
                
                if operation == "clean":
                    df = df.drop_duplicates()
                    df = df.fillna('')
                elif operation == "remove_empty":
                    df = df.dropna()
                
                output_path = os.path.join(output_dir, f"processed_{file_name}")
                self.save_excel_file(df, output_path)
            
            self.update_progress(100, "批量处理完成", force_update=True)
            self.log(f"批量处理完成！处理了 {total_files} 个文件")
            
            self.stop_indeterminate("批量处理完成", success=True)
            messagebox.showinfo("成功", f"批量处理完成！\n处理了 {total_files} 个文件\n输出目录: {output_dir}")
            
        except Exception as e:
            if "取消" in str(e):
                self.stop_indeterminate("已取消", success=False)
                self.log("操作已被用户取消")
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
