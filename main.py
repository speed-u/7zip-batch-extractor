import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import subprocess
import os
import threading
from queue import Queue
import time
import re
import json
import sys
try:
    import psutil
    HAS_PSUTIL = True
except ImportError:
    HAS_PSUTIL = False

class PasswordManager:
    def __init__(self):
        self.passwords = []
        if getattr(sys, 'frozen', False):
            # 如果是打包后的环境，使用 EXE 所在的目录
            base_dir = os.path.dirname(sys.executable)
        else:
            # 如果是普通的 Python 环境，使用脚本所在的目录
            base_dir = os.path.dirname(os.path.abspath(__file__))
        
        self.config_file = os.path.join(base_dir, "saved_passwords.json")
        # self.config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "saved_passwords.json")
        self.load_passwords()
    
    def load_passwords(self):
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.passwords = data.get('passwords', [])
        except:
            self.passwords = []
    
    def save_passwords(self):
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump({'passwords': self.passwords}, f, ensure_ascii=False, indent=2)
        except:
            pass
    
    def add_password(self, password):
        if password and password not in self.passwords:
            self.passwords.insert(0, password)
            if len(self.passwords) > 100:
                self.passwords = self.passwords[:100]
            self.save_passwords()
    
    def remove_password(self, password):
        if password in self.passwords:
            self.passwords.remove(password)
            self.save_passwords()
    
    def get_passwords(self):
        return self.passwords.copy()

class ExtractionTask:
    def __init__(self, archive_path, output_dir, password="", is_volume=False, display_name=""):
        self.archive_path = archive_path
        self.output_dir = output_dir
        self.password = password
        self.progress = 0
        self.speed = ""
        self.status = "等待中"
        self.process = None
        self.should_stop = False
        self.is_volume = is_volume
        self.display_name = display_name or os.path.basename(archive_path)

class PasswordDialog(simpledialog.Dialog):
    def __init__(self, parent, title, initial_value=""):
        self.initial_value = initial_value
        self.result_value = None
        super().__init__(parent, title)
    
    def body(self, master):
        ttk.Label(master, text="请输入密码:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.password_var = tk.StringVar(value=self.initial_value)
        self.password_entry = ttk.Entry(master, textvariable=self.password_var, width=40, show="*")
        self.password_entry.grid(row=0, column=1, padx=5, pady=5)
        self.show_password_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(master, text="显示密码", variable=self.show_password_var, 
                       command=self.toggle_password).grid(row=1, column=0, columnspan=2, pady=5)
        return self.password_entry
    
    def toggle_password(self):
        if self.show_password_var.get():
            self.password_entry.config(show="")
        else:
            self.password_entry.config(show="*")
    
    def apply(self):
        self.result_value = self.password_var.get()

class VolumeDetector:
    VOLUME_PATTERNS = [
        (r'\.7z\.(\d{3})$', '7z'),
        (r'\.(\d{3})$', '7z_simple'),
        (r'\.part(\d+)\.rar$', 'rar'),
        (r'\.z(\d+)$', 'zip'),
        (r'\.r(\d+)$', 'rar_old'),
    ]
    
    @classmethod
    def is_volume_file(cls, filepath):
        filename = os.path.basename(filepath).lower()
        for pattern, _ in cls.VOLUME_PATTERNS:
            if re.search(pattern, filename):
                return True
        return False
    
    @classmethod
    def get_first_volume(cls, filepath):
        filename = os.path.basename(filepath)
        filename_lower = filename.lower()
        dir_path = os.path.dirname(filepath)
        
        for pattern, vol_type in cls.VOLUME_PATTERNS:
            match = re.search(pattern, filename_lower)
            if match:
                if vol_type == '7z':
                    match_obj = re.search(r'\.7z\.(\d{3})$', filename_lower)
                    if match_obj:
                        first_vol = os.path.join(dir_path, filename[:match_obj.start()] + '.7z.001')
                        if os.path.exists(first_vol):
                            return first_vol, '7z'
                elif vol_type == '7z_simple':
                    match_obj = re.search(r'\.(\d{3})$', filename_lower)
                    if match_obj:
                        first_vol = os.path.join(dir_path, filename[:match_obj.start()] + '.001')
                        if os.path.exists(first_vol):
                            return first_vol, '7z'
                elif vol_type == 'rar':
                    match_obj = re.search(r'\.part(\d+)\.rar$', filename_lower)
                    if match_obj:
                        base_name = filename[:match_obj.start()]
                        first_vol = os.path.join(dir_path, base_name + '.part1.rar')
                        if os.path.exists(first_vol):
                            return first_vol, vol_type
                elif vol_type == 'zip':
                    match_obj = re.search(r'\.z(\d+)$', filename_lower)
                    if match_obj:
                        base_name = filename[:match_obj.start()]
                        first_vol = os.path.join(dir_path, base_name + '.zip')
                        if os.path.exists(first_vol):
                            return first_vol, vol_type
                elif vol_type == 'rar_old':
                    match_obj = re.search(r'\.r(\d+)$', filename_lower)
                    if match_obj:
                        base_name = filename[:match_obj.start()]
                        first_vol = os.path.join(dir_path, base_name + '.rar')
                        if os.path.exists(first_vol):
                            return first_vol, vol_type
        return filepath, None
    
    @classmethod
    def get_volume_display_name(cls, filepath, vol_type=None):
        filename = os.path.basename(filepath)
        if vol_type == '7z':
            match = re.search(r'\.7z\.(\d{3})$', filename, re.IGNORECASE)
            if match:
                return f"{filename[:match.start()]}.7z (分卷)"
            match = re.search(r'\.(\d{3})$', filename, re.IGNORECASE)
            if match:
                return f"{filename[:match.start()]}.7z (分卷)"
        elif vol_type == 'rar':
            match = re.search(r'\.part(\d+)\.rar$', filename, re.IGNORECASE)
            if match:
                base = filename[:match.start()]
                return f"{base}.rar (分卷)"
        elif vol_type == 'zip':
            if filename.lower().endswith('.zip'):
                base = filename[:-4]
                return f"{base}.zip (分卷)"
        return filename

class BatchExtractionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("批量解压工具")
        self.root.geometry("1050x750")
        
        self.file_list = []
        self.volume_map = {}
        self.tasks = {}
        self.is_extracting = False
        self.task_queue = Queue()
        self.lock = threading.Lock()
        self.completed_count = 0
        self.total_count = 0
        self.active_workers = 0
        self.max_workers = 50
        self.auto_adjust = False
        self.cpu_monitor_running = False
        self.worker_semaphore = None
        self.password_manager = PasswordManager()
        
        self.setup_ui()
        
    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(5, weight=1)
        
        title_label = ttk.Label(main_frame, text="批量解压工具 (使用7zip)", font=("微软雅黑", 14, "bold"))
        title_label.grid(row=0, column=0, pady=10)
        
        control_frame = ttk.Frame(main_frame)
        control_frame.grid(row=1, column=0, pady=5, sticky=tk.W)
        
        ttk.Button(control_frame, text="选择压缩包", command=self.select_files).grid(row=0, column=0, padx=5)
        ttk.Button(control_frame, text="清空列表", command=self.clear_list).grid(row=0, column=1, padx=5)
        
        ttk.Label(control_frame, text="并发解压数:").grid(row=0, column=2, padx=(20, 5))
        self.concurrent_var = tk.StringVar(value="3")
        self.concurrent_entry = ttk.Entry(control_frame, textvariable=self.concurrent_var, width=8)
        self.concurrent_entry.grid(row=0, column=3)
        
        self.auto_adjust_var = tk.BooleanVar(value=False)
        auto_cb = ttk.Checkbutton(control_frame, text="自动调整(根据CPU)", variable=self.auto_adjust_var)
        auto_cb.grid(row=0, column=4, padx=10)
        if not HAS_PSUTIL:
            auto_cb.config(state=tk.DISABLED)
        
        if HAS_PSUTIL:
            self.cpu_label = ttk.Label(control_frame, text="CPU: --%", foreground="blue")
            self.cpu_label.grid(row=0, column=5, padx=10)
        
        self.concurrent_label = ttk.Label(control_frame, text="当前并发: 0", foreground="purple")
        self.concurrent_label.grid(row=0, column=6, padx=10)
        
        settings_frame = ttk.Frame(main_frame)
        settings_frame.grid(row=2, column=0, pady=5, sticky=tk.W)
        
        ttk.Label(settings_frame, text="输出路径:").grid(row=0, column=0, padx=5)
        self.output_path_var = tk.StringVar(value="解压到原文件所在目录")
        self.output_path_entry = ttk.Entry(settings_frame, textvariable=self.output_path_var, width=40)
        self.output_path_entry.grid(row=0, column=1, padx=5)
        ttk.Button(settings_frame, text="浏览...", command=self.select_output_dir).grid(row=0, column=2, padx=5)
        ttk.Button(settings_frame, text="重置", command=self.reset_output_dir).grid(row=0, column=3, padx=5)
        
        ttk.Label(settings_frame, text="最大CPU%:").grid(row=0, column=4, padx=(20, 5))
        self.max_cpu_var = tk.StringVar(value="80")
        ttk.Entry(settings_frame, textvariable=self.max_cpu_var, width=5).grid(row=0, column=5)
        
        ttk.Label(settings_frame, text="最大线程:").grid(row=0, column=6, padx=(15, 5))
        self.max_workers_var = tk.StringVar(value="50")
        ttk.Entry(settings_frame, textvariable=self.max_workers_var, width=5).grid(row=0, column=7)
        
        self.delete_after_extract_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(settings_frame, text="解压后删除原文件", variable=self.delete_after_extract_var).grid(row=1, column=0, columnspan=2, padx=5, sticky=tk.W)
        
        self.create_subfolder_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(settings_frame, text="创建同名文件夹", variable=self.create_subfolder_var).grid(row=1, column=2, columnspan=2, padx=5, sticky=tk.W)
        
        ttk.Label(settings_frame, text="只解压后缀:").grid(row=1, column=4, padx=(20, 5))
        self.file_filter_var = tk.StringVar(value="")
        self.file_filter_entry = ttk.Entry(settings_frame, textvariable=self.file_filter_var, width=15)
        self.file_filter_entry.grid(row=1, column=5, padx=5)
        ttk.Label(settings_frame, text="(如: .jpg,.png)").grid(row=1, column=6, padx=5)
        
        password_frame = ttk.Frame(main_frame)
        password_frame.grid(row=3, column=0, pady=5, sticky=tk.W)
        
        ttk.Label(password_frame, text="全局密码:").grid(row=0, column=0, padx=5)
        self.global_password_var = tk.StringVar()
        self.global_password_entry = ttk.Entry(password_frame, textvariable=self.global_password_var, width=25, show="*")
        self.global_password_entry.grid(row=0, column=1, padx=5)
        
        self.show_global_password_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(password_frame, text="显示", variable=self.show_global_password_var,
                       command=self.toggle_global_password).grid(row=0, column=2, padx=5)
        
        saved_passwords = self.password_manager.get_passwords()
        self.saved_password_combo = ttk.Combobox(password_frame, values=saved_passwords, width=15)
        self.saved_password_combo.grid(row=0, column=3, padx=5)
        self.saved_password_combo.bind('<<ComboboxSelected>>', self.on_saved_password_selected)
        
        ttk.Button(password_frame, text="应用", command=self.apply_selected_password).grid(row=0, column=4, padx=2)
        ttk.Button(password_frame, text="管理密码", command=self.open_password_manager).grid(row=0, column=5, padx=5)
        
        self.auto_try_password_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(password_frame, text="自动尝试已保存密码", variable=self.auto_try_password_var).grid(row=0, column=6, padx=10)
        
        password_frame2 = ttk.Frame(main_frame)
        password_frame2.grid(row=4, column=0, pady=5, sticky=tk.W)
        
        ttk.Button(password_frame2, text="应用到选中文件", command=self.apply_password_to_selected).grid(row=0, column=0, padx=5)
        ttk.Button(password_frame2, text="应用到全部", command=self.apply_password_to_all).grid(row=0, column=1, padx=5)
        ttk.Button(password_frame2, text="设置选中文件密码", command=self.set_selected_password).grid(row=0, column=2, padx=5)
        
        tree_frame = ttk.Frame(main_frame)
        tree_frame.grid(row=5, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        
        columns = ("filename", "password", "progress", "speed", "status")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=18)
        self.tree.heading("filename", text="文件名")
        self.tree.heading("password", text="密码")
        self.tree.heading("progress", text="进度")
        self.tree.heading("speed", text="速度")
        self.tree.heading("status", text="状态")
        
        self.tree.column("filename", width=350)
        self.tree.column("password", width=120)
        self.tree.column("progress", width=80)
        self.tree.column("speed", width=130)
        self.tree.column("status", width=100)
        
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        scrollbar_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar_y.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.tree.config(yscrollcommand=scrollbar_y.set)
        
        self.tree.bind("<Double-1>", self.on_tree_double_click)
        
        self.status_label = ttk.Label(main_frame, text="就绪", foreground="green")
        self.status_label.grid(row=6, column=0, pady=5)
        
        total_progress_frame = ttk.Frame(main_frame)
        total_progress_frame.grid(row=7, column=0, sticky=(tk.W, tk.E), pady=5)
        total_progress_frame.columnconfigure(0, weight=1)
        
        ttk.Label(total_progress_frame, text="总进度:").grid(row=0, column=0, sticky=tk.W)
        self.total_progress = ttk.Progressbar(total_progress_frame, mode='determinate')
        self.total_progress.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        self.total_progress_label = ttk.Label(total_progress_frame, text="0/0")
        self.total_progress_label.grid(row=0, column=2)
        
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=8, column=0, pady=10)
        
        self.extract_button = ttk.Button(button_frame, text="开始解压", command=self.start_extraction)
        self.extract_button.grid(row=0, column=0, padx=5)
        
        self.stop_button = ttk.Button(button_frame, text="停止全部", command=self.stop_all, state=tk.DISABLED)
        self.stop_button.grid(row=0, column=1, padx=5)
        
        ttk.Button(button_frame, text="打开输出目录", command=self.open_output_dir).grid(row=0, column=2, padx=5)
        
    def toggle_global_password(self):
        if self.show_global_password_var.get():
            self.global_password_entry.config(show="")
        else:
            self.global_password_entry.config(show="*")
    
    def select_output_dir(self):
        directory = filedialog.askdirectory(title="选择输出目录")
        if directory:
            self.output_path_var.set(directory)
    
    def reset_output_dir(self):
        self.output_path_var.set("解压到原文件所在目录")
    
    def on_saved_password_selected(self, event):
        selected = self.saved_password_combo.get()
        if selected:
            self.global_password_var.set(selected)
    
    def apply_selected_password(self):
        password = self.saved_password_combo.get()
        if password:
            self.global_password_var.set(password)
            self.apply_password_to_all()
    
    def open_password_manager(self):
        manager_window = tk.Toplevel(self.root)
        manager_window.title("密码管理")
        manager_window.geometry("400x350")
        manager_window.transient(self.root)
        manager_window.grab_set()
        
        ttk.Label(manager_window, text="已保存的密码:", font=("微软雅黑", 10, "bold")).pack(pady=10)
        
        list_frame = ttk.Frame(manager_window)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        password_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, font=("Consolas", 10))
        password_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=password_listbox.yview)
        
        for pwd in self.password_manager.get_passwords():
            password_listbox.insert(tk.END, pwd)
        
        def add_password():
            dialog = PasswordDialog(manager_window, "添加密码")
            if dialog.result_value:
                self.password_manager.add_password(dialog.result_value)
                password_listbox.insert(0, dialog.result_value)
                self.refresh_saved_passwords()
        
        def delete_password():
            selection = password_listbox.curselection()
            if selection:
                password = password_listbox.get(selection[0])
                self.password_manager.remove_password(password)
                password_listbox.delete(selection[0])
                self.refresh_saved_passwords()
        
        def use_password():
            selection = password_listbox.curselection()
            if selection:
                password = password_listbox.get(selection[0])
                self.global_password_var.set(password)
                manager_window.destroy()
        
        btn_frame = ttk.Frame(manager_window)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="添加密码", command=add_password).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="删除选中", command=delete_password).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="使用选中", command=use_password).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="关闭", command=manager_window.destroy).pack(side=tk.LEFT, padx=5)
    
    def refresh_saved_passwords(self):
        passwords = self.password_manager.get_passwords()
        self.saved_password_combo['values'] = passwords
        
    def on_tree_double_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region == "cell":
            column = self.tree.identify_column(event.x)
            if column == "#2":
                item = self.tree.identify_row(event.y)
                if item:
                    self.set_single_password(item)
        
    def set_single_password(self, item_id):
        current_values = self.tree.item(item_id, "values")
        current_password = current_values[1] if len(current_values) > 1 else ""
        
        dialog = PasswordDialog(self.root, "设置密码", current_password)
        if dialog.result_value is not None:
            new_values = list(current_values)
            new_values[1] = dialog.result_value if dialog.result_value else ""
            self.tree.item(item_id, values=tuple(new_values))
            if item_id in self.tasks:
                self.tasks[item_id].password = dialog.result_value
    
    def set_selected_password(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("提示", "请先选择文件")
            return
        
        dialog = PasswordDialog(self.root, "设置密码")
        if dialog.result_value is not None:
            for item_id in selected:
                current_values = self.tree.item(item_id, "values")
                new_values = list(current_values)
                new_values[1] = dialog.result_value if dialog.result_value else ""
                self.tree.item(item_id, values=tuple(new_values))
                if item_id in self.tasks:
                    self.tasks[item_id].password = dialog.result_value
        
    def apply_password_to_selected(self):
        password = self.global_password_var.get()
        selected = self.tree.selection()
        if not selected:
            messagebox.showinfo("提示", "请先选择文件")
            return
        
        for item_id in selected:
            current_values = self.tree.item(item_id, "values")
            new_values = list(current_values)
            new_values[1] = password if password else ""
            self.tree.item(item_id, values=tuple(new_values))
            if item_id in self.tasks:
                self.tasks[item_id].password = password
        
    def apply_password_to_all(self):
        password = self.global_password_var.get()
        for item_id in self.tree.get_children():
            current_values = self.tree.item(item_id, "values")
            new_values = list(current_values)
            new_values[1] = password if password else ""
            self.tree.item(item_id, values=tuple(new_values))
            if item_id in self.tasks:
                self.tasks[item_id].password = password
    
    def select_files(self):
        files = filedialog.askopenfilenames(
            title="选择压缩包",
            filetypes=[
                ("所有支持的格式", "*.zip *.rar *.7z *.tar *.gz *.bz2 *.xz *.iso *.001 *.z01"),
                ("ZIP文件", "*.zip"),
                ("RAR文件", "*.rar"),
                ("7Z文件", "*.7z"),
                ("分卷文件", "*.001 *.z01 *.part1.rar"),
                ("所有文件", "*.*")
            ]
        )
        if files:
            added_count = 0
            for f in files:
                if f not in self.file_list:
                    first_vol, vol_type = VolumeDetector.get_first_volume(f)
                    
                    if first_vol != f:
                        if first_vol in self.volume_map:
                            continue
                        self.volume_map[first_vol] = f
                        actual_file = first_vol
                        display_name = VolumeDetector.get_volume_display_name(first_vol, vol_type)
                    else:
                        if VolumeDetector.is_volume_file(f):
                            if f in self.volume_map:
                                continue
                            first_vol, vol_type = VolumeDetector.get_first_volume(f)
                            if first_vol != f:
                                if first_vol in self.volume_map:
                                    continue
                                self.volume_map[first_vol] = f
                                actual_file = first_vol
                                display_name = VolumeDetector.get_volume_display_name(first_vol, vol_type)
                            else:
                                actual_file = f
                                display_name = VolumeDetector.get_volume_display_name(f, vol_type) if vol_type else os.path.basename(f)
                        else:
                            actual_file = f
                            display_name = os.path.basename(f)
                    
                    if actual_file not in self.file_list:
                        self.file_list.append(actual_file)
                        is_volume = vol_type is not None
                        self.tree.insert("", tk.END, iid=actual_file, values=(display_name, "", "0%", "", "等待中" + (" [分卷]" if is_volume else "")))
                        self.tasks[actual_file] = ExtractionTask(actual_file, "", "", is_volume, display_name)
                        added_count += 1
            
            if added_count > 0:
                self.update_status(f"已选择 {len(self.file_list)} 个文件")
            
    def clear_list(self):
        self.file_list.clear()
        self.tasks.clear()
        self.volume_map.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.update_status("列表已清空")
        
    def update_status(self, message, color="black"):
        self.status_label.config(text=message, foreground=color)
        
    def update_tree_item(self, archive_path, progress=None, speed=None, status=None, password=None):
        if archive_path not in self.tree.get_children():
            return
        current = self.tree.item(archive_path, "values")
        new_values = list(current)
        if password is not None:
            new_values[1] = password
        if progress is not None:
            new_values[2] = progress
        if speed is not None:
            new_values[3] = speed
        if status is not None:
            base_status = "等待中 [分卷]" if archive_path in self.tasks and self.tasks[archive_path].is_volume else "等待中"
            if status != base_status:
                new_values[4] = status
        self.tree.item(archive_path, values=tuple(new_values))
        
    def get_7zip_path(self):
        possible_paths = [
            "7z",
            "7za",
            r"C:\Program Files\7-Zip\7z.exe",
            r"C:\Program Files (x86)\7-Zip\7z.exe",
        ]
        for path in possible_paths:
            try:
                subprocess.run([path], capture_output=True, timeout=5)
                return path
            except (FileNotFoundError, subprocess.TimeoutExpired):
                continue
        return None
    
    def get_output_dir(self, archive_path):
        filename = os.path.basename(archive_path)
        
        name_lower = filename.lower()
        match_7z_vol = re.search(r'\.7z\.(\d{3})$', name_lower)
        if match_7z_vol:
            archive_name = filename[:match_7z_vol.start()]
        elif name_lower.endswith('.001'):
            archive_name = filename[:-4]
        elif re.search(r'\.part\d+\.rar$', name_lower):
            archive_name = re.sub(r'\.part\d+\.rar$', '', filename, flags=re.IGNORECASE)
        elif name_lower.endswith('.zip') and archive_path in self.tasks and self.tasks[archive_path].is_volume:
            archive_name = filename[:-4]
        else:
            archive_name = os.path.splitext(filename)[0]
        
        custom_path = self.output_path_var.get()
        if custom_path and custom_path != "解压到原文件所在目录":
            base_output = custom_path
        else:
            base_output = os.path.dirname(archive_path)
        
        if self.create_subfolder_var.get():
            output_dir = os.path.join(base_output, archive_name)
        else:
            output_dir = base_output
        
        return output_dir
    
    def extract_single(self, archive_path):
        seven_zip = self.get_7zip_path()
        if not seven_zip:
            self.root.after(0, self.update_tree_item, archive_path, None, None, "错误: 未找到7-Zip")
            return False
        
        output_dir = self.get_output_dir(archive_path)
        task = self.tasks.get(archive_path)
        if task:
            task.status = "解压中"
            task.output_dir = output_dir
        
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        passwords_to_try = []
        if task and task.password:
            passwords_to_try.append(task.password)
        
        if self.auto_try_password_var.get():
            for pwd in self.password_manager.get_passwords():
                if pwd not in passwords_to_try:
                    passwords_to_try.append(pwd)
        
        if not passwords_to_try:
            passwords_to_try.append("")
        
        for password in passwords_to_try:
            if task and task.should_stop:
                self.root.after(0, self.update_tree_item, archive_path, None, None, "已停止")
                return False
            
            result = self.try_extract(archive_path, output_dir, task, password, seven_zip)
            
            if result == "success":
                if password:
                    self.password_manager.add_password(password)
                if self.delete_after_extract_var.get():
                    self.delete_archive_files(archive_path)
                return True
            elif result == "stopped":
                return False
            elif result == "wrong_password":
                if password == passwords_to_try[-1]:
                    self.root.after(0, self.update_tree_item, archive_path, None, None, "失败: 密码错误")
                    return False
                continue
            else:
                if password == passwords_to_try[-1]:
                    return False
                continue
        
        return False
    
    def delete_archive_files(self, archive_path):
        try:
            files_to_delete = [archive_path]
            
            dir_path = os.path.dirname(archive_path)
            filename = os.path.basename(archive_path)
            filename_lower = filename.lower()
            
            match_7z = re.search(r'\.7z\.(\d{3})$', filename_lower)
            if match_7z:
                base_name = filename[:match_7z.start()]
                for i in range(1, 1000):
                    vol_file = os.path.join(dir_path, f"{base_name}.7z.{i:03d}")
                    if os.path.exists(vol_file):
                        files_to_delete.append(vol_file)
                    else:
                        break
            
            match_001 = filename_lower.endswith('.001')
            if match_001:
                base_name = filename[:-4]
                for i in range(1, 1000):
                    vol_file = os.path.join(dir_path, f"{base_name}.{i:03d}")
                    if os.path.exists(vol_file):
                        files_to_delete.append(vol_file)
                    else:
                        break
            
            match_rar = re.search(r'\.part(\d+)\.rar$', filename_lower)
            if match_rar:
                base_name = filename[:match_rar.start()]
                for i in range(1, 100):
                    vol_file = os.path.join(dir_path, f"{base_name}.part{i}.rar")
                    if os.path.exists(vol_file):
                        files_to_delete.append(vol_file)
                    else:
                        break
            
            match_z = re.search(r'\.z(\d+)$', filename_lower)
            if match_z:
                base_name = filename[:match_z.start()]
                zip_file = os.path.join(dir_path, f"{base_name}.zip")
                if os.path.exists(zip_file):
                    files_to_delete.append(zip_file)
                for i in range(1, 100):
                    vol_file = os.path.join(dir_path, f"{base_name}.z{i:02d}")
                    if os.path.exists(vol_file):
                        files_to_delete.append(vol_file)
                    else:
                        break
            
            match_r = re.search(r'\.r(\d+)$', filename_lower)
            if match_r:
                base_name = filename[:match_r.start()]
                rar_file = os.path.join(dir_path, f"{base_name}.rar")
                if os.path.exists(rar_file):
                    files_to_delete.append(rar_file)
                for i in range(0, 100):
                    vol_file = os.path.join(dir_path, f"{base_name}.r{i:02d}")
                    if os.path.exists(vol_file):
                        files_to_delete.append(vol_file)
                    else:
                        break
            
            files_to_delete = list(set(files_to_delete))
            
            for f in files_to_delete:
                try:
                    if os.path.exists(f):
                        os.remove(f)
                except:
                    pass
                    
        except:
            pass
    
    def try_extract(self, archive_path, output_dir, task, password, seven_zip):
        try:
            cmd = [
                seven_zip, "x",
                "-y",
                f"-o{output_dir}",
                archive_path,
            ]
            
            file_filter = self.file_filter_var.get().strip()
            if file_filter:
                extensions = [ext.strip() for ext in file_filter.split(',') if ext.strip()]
                for ext in extensions:
                    if not ext.startswith('*'):
                        ext = '*' + ext
                    cmd.append(ext)
                cmd.append("-r")
            
            cmd.append("-bsp1")
            
            if password:
                cmd.append(f"-p{password}")
            
            task.process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding='utf-8',
                errors='replace',
                bufsize=1
            )
            
            progress_pattern = re.compile(r'(\d+)%')
            speed_pattern = re.compile(r'(\d+(?:\.\d+)?\s*[KMGT]?B/s)')
            
            while True:
                if task and task.should_stop:
                    task.process.terminate()
                    return "stopped"
                    
                line = task.process.stdout.readline()
                if not line and task.process.poll() is not None:
                    break
                    
                if line:
                    progress_match = progress_pattern.search(line)
                    if progress_match:
                        progress = progress_match.group(1) + "%"
                        self.root.after(0, self.update_tree_item, archive_path, progress, None, "解压中")
                        if task:
                            task.progress = int(progress_match.group(1))
                    
                    speed_match = speed_pattern.search(line)
                    if speed_match:
                        speed = speed_match.group(1)
                        self.root.after(0, self.update_tree_item, archive_path, None, speed, None)
                        if task:
                            task.speed = speed
            
            return_code = task.process.wait()
            
            if return_code == 0:
                self.root.after(0, self.update_tree_item, archive_path, "100%", "", "完成")
                return "success"
            else:
                stderr = task.process.stderr.read()
                if "Wrong password" in stderr or "密码错误" in stderr or "Data Error" in stderr:
                    return "wrong_password"
                elif "Can not open" in stderr:
                    self.root.after(0, self.update_tree_item, archive_path, None, None, "失败: 无法打开文件")
                    return "error"
                elif "Unexpected end" in stderr:
                    self.root.after(0, self.update_tree_item, archive_path, None, None, "失败: 分卷文件不完整")
                    return "error"
                else:
                    error_msg = stderr[:60] if stderr else "解压失败"
                    self.root.after(0, self.update_tree_item, archive_path, None, None, f"失败: {error_msg}")
                    return "error"
                
        except Exception as e:
            self.root.after(0, self.update_tree_item, archive_path, None, None, f"错误: {str(e)[:50]}")
            return "error"
    
    def cpu_monitor(self):
        while self.cpu_monitor_running and self.is_extracting:
            try:
                cpu_percent = psutil.cpu_percent(interval=1.0)
                self.root.after(0, self.update_cpu_display, cpu_percent)
                
                if self.auto_adjust_var.get():
                    try:
                        max_cpu = int(self.max_cpu_var.get())
                        if max_cpu < 10:
                            max_cpu = 10
                        elif max_cpu > 100:
                            max_cpu = 100
                    except:
                        max_cpu = 80
                    
                    try:
                        max_workers = int(self.max_workers_var.get())
                        if max_workers < 1:
                            max_workers = 1
                    except:
                        max_workers = 50
                    
                    if cpu_percent < max_cpu and self.active_workers < max_workers:
                        with self.lock:
                            if self.active_workers < max_workers and not self.task_queue.empty():
                                t = threading.Thread(target=self.worker, daemon=True)
                                t.start()
            except:
                pass
    
    def update_cpu_display(self, cpu_percent):
        if hasattr(self, 'cpu_label'):
            color = "green" if cpu_percent < 60 else "orange" if cpu_percent < 80 else "red"
            self.cpu_label.config(text=f"CPU: {cpu_percent:.0f}%", foreground=color)
    
    def update_concurrent_display(self):
        self.concurrent_label.config(text=f"当前并发: {self.active_workers}")
    
    def worker(self):
        with self.lock:
            self.active_workers += 1
            self.root.after(0, self.update_concurrent_display)
        
        while True:
            try:
                archive_path = self.task_queue.get(timeout=0.5)
            except:
                break
            
            task = self.tasks.get(archive_path)
            if task and task.should_stop:
                continue
                
            self.extract_single(archive_path)
            
            with self.lock:
                self.completed_count += 1
                progress = (self.completed_count / self.total_count) * 100
                self.root.after(0, self.update_total_progress, progress, self.completed_count, self.total_count)
            
            self.task_queue.task_done()
        
        with self.lock:
            self.active_workers -= 1
            self.root.after(0, self.update_concurrent_display)
    
    def update_total_progress(self, progress, completed, total):
        self.total_progress.config(value=progress)
        self.total_progress_label.config(text=f"{completed}/{total}")
        if completed == total:
            self.cpu_monitor_running = False
            self.update_status(f"全部完成！共处理 {completed} 个文件", "green")
            self.is_extracting = False
            self.extract_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)
            if hasattr(self, 'cpu_label'):
                self.cpu_label.config(text="CPU: --%", foreground="blue")
            self.concurrent_label.config(text="当前并发: 0")
            messagebox.showinfo("完成", f"解压完成！共处理 {completed} 个文件")
    
    def start_extraction(self):
        if not self.file_list:
            messagebox.showwarning("警告", "请先选择压缩包")
            return
        
        try:
            max_concurrent = int(self.concurrent_var.get())
            if max_concurrent < 1:
                max_concurrent = 1
        except ValueError:
            max_concurrent = 3
        
        self.completed_count = 0
        self.total_count = len(self.file_list)
        self.is_extracting = True
        self.active_workers = 0
        
        for archive in self.file_list:
            output_dir = self.get_output_dir(archive)
            values = self.tree.item(archive, "values")
            password = values[1] if len(values) > 1 else ""
            if archive in self.tasks:
                self.tasks[archive].password = password
                self.tasks[archive].output_dir = output_dir
            self.task_queue.put(archive)
        
        self.total_progress.config(value=0)
        self.total_progress_label.config(text=f"0/{self.total_count}")
        self.update_status(f"正在解压... (并发数: {max_concurrent})", "blue")
        
        self.extract_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        
        if HAS_PSUTIL:
            self.cpu_monitor_running = True
            monitor_thread = threading.Thread(target=self.cpu_monitor, daemon=True)
            monitor_thread.start()
        
        for i in range(max_concurrent):
            t = threading.Thread(target=self.worker, daemon=True)
            t.start()
        
    def stop_all(self):
        self.cpu_monitor_running = False
        self.is_extracting = False
        self.active_workers = 0
        for archive_path, task in self.tasks.items():
            task.should_stop = True
            if task.process:
                try:
                    task.process.terminate()
                except:
                    pass
        self.update_status("正在停止所有任务...", "orange")
        if hasattr(self, 'cpu_label'):
            self.cpu_label.config(text="CPU: --%", foreground="blue")
        self.concurrent_label.config(text="当前并发: 0")
        
    def open_output_dir(self):
        if self.file_list:
            output_dir = self.get_output_dir(self.file_list[0])
            if os.path.exists(output_dir):
                os.startfile(output_dir)
            else:
                messagebox.showinfo("提示", "输出目录不存在，请先解压文件")
        else:
            messagebox.showinfo("提示", "请先选择并解压文件")

def main():
    root = tk.Tk()
    app = BatchExtractionApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
