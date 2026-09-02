import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import os
import threading
from datetime import datetime
import re

class DataProcessingGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel/Word 数据处理工具")
        self.root.geometry("1000x700")
        
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
        self.create_clean_tab()
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
        
        # 设置窗口最小大小
        self.root.minsize(800, 600)
        
        # 设置窗口居中
        self.center_window()
    
    def on_frame_configure(self, event):
        """更新Canvas的滚动区域"""
        self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))
    
    def on_canvas_configure(self, event):
        """调整Canvas中窗口的宽度"""
        self.main_canvas.itemconfig(self.canvas_window, width=event.width)
    
    def bind_mousewheel_recursive(self, widget):
        """递归绑定鼠标滚轮事件到所有组件"""
        # 绑定鼠标滚轮事件
        widget.bind("<MouseWheel>", self.on_mousewheel_windows)
        widget.bind("<Button-4>", self.on_mousewheel_linux)
        widget.bind("<Button-5>", self.on_mousewheel_linux)
        
        # 递归绑定到所有子组件
        for child in widget.winfo_children():
            self.bind_mousewheel_recursive(child)
    
    def on_mousewheel_windows(self, event):
        """Windows鼠标滚轮事件"""
        if event.delta > 0:
            self.main_canvas.yview_scroll(-3, "units")
        else:
            self.main_canvas.yview_scroll(3, "units")
        return "break"  # 阻止事件传递
    
    def on_mousewheel_linux(self, event):
        """Linux鼠标滚轮事件"""
        if event.num == 4:
            self.main_canvas.yview_scroll(-3, "units")
        elif event.num == 5:
            self.main_canvas.yview_scroll(3, "units")
        return "break"  # 阻止事件传递
    
    def center_window(self):
        """窗口居中"""
        self.root.update_idletasks()
        width = 1000
        height = 700
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def create_progress_area(self):
        """创建进度条区域"""
        progress_frame = ttk.LabelFrame(self.main_frame, text="处理进度", padding="5")
        progress_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        
        # 进度条
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, 
                                           maximum=100, length=800, mode='determinate')
        self.progress_bar.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # 进度百分比标签
        self.progress_label = ttk.Label(progress_frame, text="0%")
        self.progress_label.grid(row=0, column=2, padx=10)
        
        # 状态标签
        self.status_label = ttk.Label(progress_frame, text="就绪", foreground="green")
        self.status_label.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=5)
        
        # 不确定进度条
        self.indeterminate_progress = ttk.Progressbar(progress_frame, mode='indeterminate',
                                                      length=800)
        self.indeterminate_progress.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        self.indeterminate_progress.grid_remove()
        
        # 取消按钮
        self.cancel_button = ttk.Button(progress_frame, text="取消", command=self.cancel_operation,
                                       state='disabled')
        self.cancel_button.grid(row=2, column=2, padx=10)
        
        self.cancel_flag = False
    
    def create_log_area(self):
        """创建日志显示区域"""
        log_frame = ttk.LabelFrame(self.main_frame, text="操作日志", padding="5")
        log_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=10)
        
        # 日志文本框
        self.log_text = scrolledtext.ScrolledText(log_frame, height=5, width=100)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        # 清除日志按钮
        clear_btn = ttk.Button(log_frame, text="清除日志", command=self.clear_log)
        clear_btn.grid(row=1, column=0, pady=5)
    
    def clear_log(self):
        """清除日志"""
        self.log_text.delete(1.0, tk.END)
    
    def log(self, message):
        """添加日志"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update()
    
    def update_progress(self, value, status_text=None):
        """更新进度条"""
        self.progress_var.set(value)
        self.progress_label.config(text=f"{int(value)}%")
        if status_text:
            self.status_label.config(text=status_text)
        self.root.update()
    
    def start_indeterminate(self, status_text="处理中..."):
        """启动不确定进度条"""
        self.indeterminate_progress.grid()
        self.indeterminate_progress.start(10)
        self.status_label.config(text=status_text, foreground="blue")
        self.cancel_button.config(state='normal')
        self.cancel_flag = False
        self.is_processing = True
        self.root.update()
    
    def stop_indeterminate(self, status_text="完成", success=True):
        """停止不确定进度条"""
        self.indeterminate_progress.stop()
        self.indeterminate_progress.grid_remove()
        if success:
            self.status_label.config(text=status_text, foreground="green")
            self.progress_var.set(100)
            self.progress_label.config(text="100%")
        else:
            self.status_label.config(text=status_text, foreground="red")
        self.cancel_button.config(state='disabled')
        self.is_processing = False
        self.root.update()
    
    def cancel_operation(self):
        """取消操作"""
        self.cancel_flag = True
        self.status_label.config(text="正在取消...", foreground="orange")
        self.log("用户请求取消操作")
    
    def check_cancel(self):
        """检查是否取消"""
        if self.cancel_flag:
            raise Exception("操作已被用户取消")
    
    def create_merge_tab(self):
        """创建数据合并标签页"""
        merge_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(merge_frame, text="数据合并")
        
        # 文件选择区域
        file_frame = ttk.LabelFrame(merge_frame, text="选择要合并的Excel文件", padding="10")
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
        
        # 合并选项
        option_frame = ttk.LabelFrame(merge_frame, text="合并选项", padding="10")
        option_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        # 合并方式
        ttk.Label(option_frame, text="合并方式:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.merge_type = tk.StringVar(value="vertical")
        ttk.Radiobutton(option_frame, text="垂直合并（按行追加）", variable=self.merge_type, 
                       value="vertical").grid(row=0, column=1, sticky=tk.W, padx=10)
        ttk.Radiobutton(option_frame, text="水平合并（按列拼接）", variable=self.merge_type, 
                       value="horizontal").grid(row=0, column=2, sticky=tk.W, padx=10)
        
        # 高级选项
        self.add_source_var = tk.BooleanVar()
        ttk.Checkbutton(option_frame, text="添加数据来源标识", 
                       variable=self.add_source_var).grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        self.remove_dup_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(option_frame, text="去除重复行", 
                       variable=self.remove_dup_var).grid(row=1, column=2, columnspan=2, sticky=tk.W, pady=5)
        
        # 输出设置
        output_frame = ttk.LabelFrame(merge_frame, text="输出设置", padding="10")
        output_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(output_frame, text="输出文件:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.output_path = tk.StringVar()
        ttk.Entry(output_frame, textvariable=self.output_path, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(output_frame, text="浏览", command=self.select_output_file).grid(row=0, column=2)
        
        # 执行按钮
        self.merge_button = ttk.Button(merge_frame, text="开始合并", command=self.start_merge, width=20)
        self.merge_button.grid(row=3, column=0, columnspan=3, pady=10)
    
    def create_clean_tab(self):
        """创建数据清理标签页"""
        clean_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(clean_frame, text="数据清理")
        
        # 文件选择
        file_frame = ttk.LabelFrame(clean_frame, text="选择文件", padding="10")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        self.clean_file_path = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.clean_file_path, width=60).grid(row=0, column=0, padx=5)
        ttk.Button(file_frame, text="浏览", command=self.select_clean_file).grid(row=0, column=1)
        
        # 清理选项
        option_frame = ttk.LabelFrame(clean_frame, text="清理选项", padding="10")
        option_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        self.remove_duplicates = tk.BooleanVar(value=True)
        ttk.Checkbutton(option_frame, text="去除重复行", 
                       variable=self.remove_duplicates).grid(row=0, column=0, sticky=tk.W, pady=5)
        
        self.fill_na = tk.BooleanVar(value=True)
        ttk.Checkbutton(option_frame, text="填充空值", 
                       variable=self.fill_na).grid(row=0, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(option_frame, text="填充值:").grid(row=0, column=2, sticky=tk.W, padx=10)
        self.fill_value = tk.StringVar(value="")
        ttk.Entry(option_frame, textvariable=self.fill_value, width=15).grid(row=0, column=3)
        
        self.strip_spaces = tk.BooleanVar(value=True)
        ttk.Checkbutton(option_frame, text="去除首尾空格", 
                       variable=self.strip_spaces).grid(row=1, column=0, sticky=tk.W, pady=5)
        
        # 执行按钮
        self.clean_button = ttk.Button(clean_frame, text="开始清理", command=self.start_clean, width=20)
        self.clean_button.grid(row=2, column=0, columnspan=2, pady=10)
    
    def create_convert_tab(self):
        """创建格式转换标签页"""
        convert_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(convert_frame, text="格式转换")
        
        # 转换类型选择
        type_frame = ttk.LabelFrame(convert_frame, text="转换类型", padding="10")
        type_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        self.convert_type = tk.StringVar(value="excel_to_word")
        ttk.Radiobutton(type_frame, text="Excel转Word", variable=self.convert_type, 
                       value="excel_to_word").grid(row=0, column=0, padx=10)
        ttk.Radiobutton(type_frame, text="Word转Excel", variable=self.convert_type, 
                       value="word_to_excel").grid(row=0, column=1, padx=10)
        
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
        """创建数据拆分标签页"""
        split_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(split_frame, text="数据拆分")
        
        # 文件选择
        file_frame = ttk.LabelFrame(split_frame, text="选择文件", padding="10")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        self.split_file_path = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.split_file_path, width=60).grid(row=0, column=0, padx=5)
        ttk.Button(file_frame, text="浏览", command=self.select_split_file).grid(row=0, column=1)
        
        # 拆分设置
        setting_frame = ttk.LabelFrame(split_frame, text="拆分设置", padding="10")
        setting_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(setting_frame, text="拆分依据列:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.split_column = ttk.Entry(setting_frame, width=30)
        self.split_column.grid(row=0, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(setting_frame, text="输出目录:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.split_output_dir = tk.StringVar(value="./split_output")
        ttk.Entry(setting_frame, textvariable=self.split_output_dir, width=40).grid(row=1, column=1, pady=5)
        ttk.Button(setting_frame, text="浏览", command=self.select_split_output).grid(row=1, column=2, padx=5)
        
        # 执行按钮
        self.split_button = ttk.Button(split_frame, text="开始拆分", command=self.start_split, width=20)
        self.split_button.grid(row=2, column=0, columnspan=2, pady=10)
    
    def create_batch_tab(self):
        """创建批量处理标签页"""
        batch_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(batch_frame, text="批量处理")
        
        # 目录选择
        dir_frame = ttk.LabelFrame(batch_frame, text="选择文件夹", padding="10")
        dir_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        self.batch_dir = tk.StringVar()
        ttk.Entry(dir_frame, textvariable=self.batch_dir, width=60).grid(row=0, column=0, padx=5)
        ttk.Button(dir_frame, text="浏览", command=self.select_batch_dir).grid(row=0, column=1)
        
        # 处理选项
        option_frame = ttk.LabelFrame(batch_frame, text="处理选项", padding="10")
        option_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        self.batch_operation = tk.StringVar(value="clean")
        ttk.Radiobutton(option_frame, text="清理数据", variable=self.batch_operation, 
                       value="clean").grid(row=0, column=0, padx=10)
        ttk.Radiobutton(option_frame, text="去除空行", variable=self.batch_operation, 
                       value="remove_empty").grid(row=0, column=1, padx=10)
        
        # 执行按钮
        self.batch_button = ttk.Button(batch_frame, text="开始批量处理", command=self.start_batch, width=20)
        self.batch_button.grid(row=2, column=0, columnspan=2, pady=10)
    
    # 文件选择方法
    def add_files(self):
        files = filedialog.askopenfilenames(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xls *.csv"), ("All files", "*.*")]
        )
        for file in files:
            if file not in self.file_listbox.get(0, tk.END):
                self.file_listbox.insert(tk.END, file)
                self.log(f"添加文件: {os.path.basename(file)}")
    
    def remove_selected_files(self):
        selected = self.file_listbox.curselection()
        for index in reversed(selected):
            self.file_listbox.delete(index)
            self.log("移除文件")
    
    def clear_file_list(self):
        self.file_listbox.delete(0, tk.END)
        self.log("清空文件列表")
    
    def select_output_file(self):
        file_path = filedialog.asksaveasfilename(
            title="保存文件",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
        )
        if file_path:
            self.output_path.set(file_path)
    
    def select_clean_file(self):
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xls *.csv")]
        )
        if file_path:
            self.clean_file_path.set(file_path)
    
    def select_input_file(self):
        if self.convert_type.get() == "excel_to_word":
            file_path = filedialog.askopenfilename(
                title="选择Excel文件",
                filetypes=[("Excel files", "*.xlsx *.xls")]
            )
        else:
            file_path = filedialog.askopenfilename(
                title="选择Word文件",
                filetypes=[("Word files", "*.docx")]
            )
        if file_path:
            self.convert_input.set(file_path)
    
    def select_convert_output(self):
        if self.convert_type.get() == "excel_to_word":
            file_path = filedialog.asksaveasfilename(
                title="保存Word文件",
                defaultextension=".docx",
                filetypes=[("Word files", "*.docx")]
            )
        else:
            file_path = filedialog.asksaveasfilename(
                title="保存Excel文件",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")]
            )
        if file_path:
            self.convert_output.set(file_path)
    
    def select_split_file(self):
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xls *.csv")]
        )
        if file_path:
            self.split_file_path.set(file_path)
    
    def select_split_output(self):
        dir_path = filedialog.askdirectory(title="选择输出目录")
        if dir_path:
            self.split_output_dir.set(dir_path)
    
    def select_batch_dir(self):
        dir_path = filedialog.askdirectory(title="选择文件夹")
        if dir_path:
            self.batch_dir.set(dir_path)
    
    # 处理功能方法
    def start_merge(self):
        if self.is_processing:
            messagebox.showwarning("警告", "正在处理中，请等待当前操作完成")
            return
            
        files = list(self.file_listbox.get(0, tk.END))
        if not files:
            messagebox.showwarning("警告", "请先添加要合并的文件")
            return
        
        output_path = self.output_path.get()
        if not output_path:
            messagebox.showwarning("警告", "请指定输出文件路径")
            return
        
        # 禁用按钮
        self.merge_button.config(state='disabled')
        # 在新线程中执行合并操作
        thread = threading.Thread(target=self.merge_thread, args=(files, output_path))
        thread.daemon = True
        thread.start()
    
    def merge_thread(self, files, output_path):
        try:
            self.start_indeterminate("正在合并文件...")
            self.log("开始合并文件...")
            
            # 重置进度
            self.progress_var.set(0)
            self.progress_label.config(text="0%")
            
            # 读取所有文件
            dfs = []
            total_files = len(files)
            for i, file in enumerate(files):
                self.check_cancel()
                self.log(f"读取文件 {i+1}/{total_files}: {os.path.basename(file)}")
                self.update_progress((i / total_files) * 50, f"读取文件 {i+1}/{total_files}")
                
                df = pd.read_excel(file)
                
                # 添加数据来源列
                if self.add_source_var.get():
                    df['数据来源'] = os.path.basename(file)
                
                dfs.append(df)
            
            self.check_cancel()
            self.update_progress(60, "正在合并数据...")
            self.log("正在合并数据...")
            
            # 合并数据
            merge_type = self.merge_type.get()
            if merge_type == "vertical":
                merged_df = pd.concat(dfs, ignore_index=True)
            else:
                merged_df = pd.concat(dfs, axis=1)
            
            self.update_progress(80, "正在处理数据...")
            
            # 去除重复行
            if self.remove_dup_var.get():
                self.check_cancel()
                original_count = len(merged_df)
                merged_df = merged_df.drop_duplicates()
                self.log(f"去除重复行: {original_count - len(merged_df)}行")
            
            self.check_cancel()
            self.update_progress(90, "正在保存结果...")
            self.log("正在保存结果...")
            
            # 保存结果
            merged_df.to_excel(output_path, index=False)
            
            self.update_progress(100, "合并完成")
            self.log(f"合并完成！结果已保存到: {output_path}")
            self.log(f"合并后数据: {len(merged_df)}行, {len(merged_df.columns)}列")
            
            self.stop_indeterminate("合并完成", success=True)
            messagebox.showinfo("成功", f"合并完成！\n输出文件: {output_path}\n共处理 {len(merged_df)} 行数据")
            
        except Exception as e:
            if "取消" in str(e):
                self.stop_indeterminate("已取消", success=False)
                self.log("操作已被用户取消")
            else:
                self.stop_indeterminate("处理失败", success=False)
                self.log(f"合并失败: {str(e)}")
                messagebox.showerror("错误", f"合并失败: {str(e)}")
        finally:
            # 恢复按钮
            self.root.after(0, lambda: self.merge_button.config(state='normal'))
    
    def start_clean(self):
        if self.is_processing:
            messagebox.showwarning("警告", "正在处理中，请等待当前操作完成")
            return
            
        file_path = self.clean_file_path.get()
        if not file_path:
            messagebox.showwarning("警告", "请选择要清理的文件")
            return
        
        self.clean_button.config(state='disabled')
        thread = threading.Thread(target=self.clean_thread, args=(file_path,))
        thread.daemon = True
        thread.start()
    
    def clean_thread(self, file_path):
        try:
            self.start_indeterminate("正在清理数据...")
            self.log(f"开始清理文件: {os.path.basename(file_path)}")
            
            # 读取文件
            self.update_progress(10, "正在读取文件...")
            df = pd.read_excel(file_path)
            original_shape = df.shape
            self.log(f"原始数据: {original_shape[0]}行, {original_shape[1]}列")
            
            # 去除重复行
            if self.remove_duplicates.get():
                self.check_cancel()
                self.update_progress(30, "正在去除重复行...")
                df = df.drop_duplicates()
                self.log(f"去除重复行后: {df.shape[0]}行")
            
            # 填充空值
            if self.fill_na.get():
                self.check_cancel()
                self.update_progress(50, "正在填充空值...")
                fill_value = self.fill_value.get()
                if fill_value == "":
                    fill_value = None
                df = df.fillna(fill_value)
                self.log("填充空值完成")
            
            # 去除首尾空格
            if self.strip_spaces.get():
                self.check_cancel()
                self.update_progress(70, "正在去除空格...")
                for col in df.columns:
                    if df[col].dtype == 'object':
                        df[col] = df[col].str.strip()
                self.log("去除首尾空格完成")
            
            # 保存清理后的文件
            self.check_cancel()
            self.update_progress(90, "正在保存文件...")
            output_path = file_path.replace('.xlsx', '_cleaned.xlsx').replace('.xls', '_cleaned.xlsx').replace('.csv', '_cleaned.csv')
            df.to_excel(output_path, index=False)
            
            self.update_progress(100, "清理完成")
            self.log(f"清理完成！")
            self.log(f"清理后数据: {df.shape[0]}行, {df.shape[1]}列")
            self.log(f"结果已保存到: {output_path}")
            
            self.stop_indeterminate("清理完成", success=True)
            messagebox.showinfo("成功", f"清理完成！\n输出文件: {output_path}\n清理后数据: {df.shape[0]}行, {df.shape[1]}列")
            
        except Exception as e:
            if "取消" in str(e):
                self.stop_indeterminate("已取消", success=False)
                self.log("操作已被用户取消")
            else:
                self.stop_indeterminate("处理失败", success=False)
                self.log(f"清理失败: {str(e)}")
                messagebox.showerror("错误", f"清理失败: {str(e)}")
        finally:
            self.root.after(0, lambda: self.clean_button.config(state='normal'))
    
    def start_convert(self):
        if self.is_processing:
            messagebox.showwarning("警告", "正在处理中，请等待当前操作完成")
            return
            
        input_path = self.convert_input.get()
        output_path = self.convert_output.get()
        
        if not input_path:
            messagebox.showwarning("警告", "请选择输入文件")
            return
        if not output_path:
            messagebox.showwarning("警告", "请指定输出文件路径")
            return
        
        self.convert_button.config(state='disabled')
        thread = threading.Thread(target=self.convert_thread, args=(input_path, output_path))
        thread.daemon = True
        thread.start()
    
    def convert_thread(self, input_path, output_path):
        try:
            convert_type = self.convert_type.get()
            self.start_indeterminate("正在转换...")
            self.log(f"开始转换: {convert_type}")
            
            if convert_type == "excel_to_word":
                from docx import Document
                
                # 读取Excel
                self.update_progress(20, "正在读取Excel文件...")
                df = pd.read_excel(input_path)
                self.log(f"读取Excel: {df.shape[0]}行, {df.shape[1]}列")
                
                # 创建Word文档
                self.check_cancel()
                self.update_progress(50, "正在创建Word文档...")
                doc = Document()
                doc.add_heading('Excel数据转换结果', level=1)
                
                # 添加表格
                table = doc.add_table(rows=1, cols=len(df.columns))
                table.style = 'Light Grid Accent 1'
                
                # 添加表头
                header_cells = table.rows[0].cells
                for i, col in enumerate(df.columns):
                    header_cells[i].text = str(col)
                
                # 添加数据
                total_rows = len(df)
                for idx, (_, row) in enumerate(df.iterrows()):
                    self.check_cancel()
                    if idx % 10 == 0:  # 每10行更新一次进度
                        progress = 50 + (idx / total_rows) * 40
                        self.update_progress(progress, f"正在转换数据 {idx}/{total_rows}")
                    
                    row_cells = table.add_row().cells
                    for i, value in enumerate(row):
                        row_cells[i].text = str(value)
                
                self.update_progress(90, "正在保存Word文件...")
                doc.save(output_path)
                self.log(f"Excel转Word完成: {output_path}")
                
            else:  # word_to_excel
                from docx import Document
                
                # 读取Word
                self.update_progress(30, "正在读取Word文件...")
                doc = Document(input_path)
                
                # 获取第一个表格
                if not doc.tables:
                    raise Exception("Word文档中没有表格")
                
                self.check_cancel()
                self.update_progress(60, "正在提取表格数据...")
                table = doc.tables[0]
                data = []
                for row in table.rows:
                    row_data = [cell.text for cell in row.cells]
                    data.append(row_data)
                
                # 转换为DataFrame
                self.update_progress(80, "正在转换为Excel...")
                df = pd.DataFrame(data[1:], columns=data[0])
                df.to_excel(output_path, index=False)
                self.log(f"Word转Excel完成: {output_path}")
            
            self.update_progress(100, "转换完成")
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
        finally:
            self.root.after(0, lambda: self.convert_button.config(state='normal'))
    
    def start_split(self):
        if self.is_processing:
            messagebox.showwarning("警告", "正在处理中，请等待当前操作完成")
            return
            
        file_path = self.split_file_path.get()
        column_name = self.split_column.get()
        output_dir = self.split_output_dir.get()
        
        if not file_path:
            messagebox.showwarning("警告", "请选择要拆分的文件")
            return
        if not column_name:
            messagebox.showwarning("警告", "请输入拆分依据列名")
            return
        
        self.split_button.config(state='disabled')
        thread = threading.Thread(target=self.split_thread, args=(file_path, column_name, output_dir))
        thread.daemon = True
        thread.start()
    
    def split_thread(self, file_path, column_name, output_dir):
        try:
            self.start_indeterminate("正在拆分数据...")
            self.log(f"开始拆分文件: {os.path.basename(file_path)}")
            
            # 读取文件
            self.update_progress(10, "正在读取文件...")
            df = pd.read_excel(file_path)
            
            if column_name not in df.columns:
                raise Exception(f"列 '{column_name}' 不存在")
            
            # 创建输出目录
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # 按列拆分
            groups = df.groupby(column_name)
            total_groups = len(groups)
            
            self.update_progress(20, f"正在拆分数据，共{total_groups}组...")
            
            for i, (name, group) in enumerate(groups):
                self.check_cancel()
                safe_name = re.sub(r'[\\/*?:"<>|]', '_', str(name))
                output_path = os.path.join(output_dir, f"{safe_name}.xlsx")
                group.to_excel(output_path, index=False)
                
                progress = 20 + ((i + 1) / total_groups) * 70
                self.update_progress(progress, f"正在保存: {safe_name}.xlsx ({i+1}/{total_groups})")
                self.log(f"已保存: {safe_name}.xlsx ({len(group)}行)")
            
            self.update_progress(100, "拆分完成")
            self.log(f"拆分完成！共生成 {total_groups} 个文件")
            
            self.stop_indeterminate("拆分完成", success=True)
            messagebox.showinfo("成功", f"拆分完成！\n生成 {total_groups} 个文件\n输出目录: {output_dir}")
            
        except Exception as e:
            if "取消" in str(e):
                self.stop_indeterminate("已取消", success=False)
                self.log("操作已被用户取消")
            else:
                self.stop_indeterminate("拆分失败", success=False)
                self.log(f"拆分失败: {str(e)}")
                messagebox.showerror("错误", f"拆分失败: {str(e)}")
        finally:
            self.root.after(0, lambda: self.split_button.config(state='normal'))
    
    def start_batch(self):
        if self.is_processing:
            messagebox.showwarning("警告", "正在处理中，请等待当前操作完成")
            return
            
        directory = self.batch_dir.get()
        if not directory:
            messagebox.showwarning("警告", "请选择要处理的文件夹")
            return
        
        self.batch_button.config(state='disabled')
        thread = threading.Thread(target=self.batch_thread, args=(directory,))
        thread.daemon = True
        thread.start()
    
    def batch_thread(self, directory):
        try:
            self.start_indeterminate("正在批量处理...")
            self.log(f"开始批量处理: {directory}")
            
            # 获取所有Excel文件
            excel_files = [f for f in os.listdir(directory) 
                          if f.endswith(('.xlsx', '.xls', '.csv'))]
            
            if not excel_files:
                self.log("目录中没有Excel文件")
                self.stop_indeterminate("未找到文件", success=False)
                messagebox.showwarning("警告", "目录中没有Excel文件")
                return
            
            # 创建输出目录
            output_dir = os.path.join(directory, 'processed')
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            operation = self.batch_operation.get()
            total_files = len(excel_files)
            
            for i, file_name in enumerate(excel_files):
                self.check_cancel()
                file_path = os.path.join(directory, file_name)
                self.log(f"处理文件 {i+1}/{total_files}: {file_name}")
                
                progress = (i / total_files) * 100
                self.update_progress(progress, f"正在处理 {i+1}/{total_files}: {file_name}")
                
                # 读取文件
                if file_name.endswith('.csv'):
                    df = pd.read_csv(file_path)
                else:
                    df = pd.read_excel(file_path)
                
                # 执行操作
                if operation == "clean":
                    df = df.drop_duplicates()
                    df = df.fillna('')
                elif operation == "remove_empty":
                    df = df.dropna()
                
                # 保存处理后的文件
                output_path = os.path.join(output_dir, f"processed_{file_name}")
                if file_name.endswith('.csv'):
                    df.to_csv(output_path, index=False)
                else:
                    df.to_excel(output_path, index=False)
                
                self.log(f"已处理: {file_name}")
            
            self.update_progress(100, "批量处理完成")
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
        finally:
            self.root.after(0, lambda: self.batch_button.config(state='normal'))

def main():
    root = tk.Tk()
    app = DataProcessingGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
