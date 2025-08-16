import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import shutil
import ctypes
from ctypes import wintypes

# 系统API配置
shell32 = ctypes.WinDLL("shell32", use_last_error=True)
SHERB_NOCONFIRMATION = 0x00000001
SHERB_NOPROGRESSUI = 0x00000002
SHERB_NOSOUND = 0x00000004

# 回收站信息结构体
class SHQUERYRBINFO(ctypes.Structure):
    _fields_ = [("cbSize", wintypes.DWORD),
                ("i64Size", ctypes.c_longlong),
                ("i64NumItems", ctypes.c_longlong)]

class WindowsCleaner:
    def __init__(self, root):
        # 基础配置
        self.root = root
        self.root.title("Windows Cleaner")
        self.root.geometry("390x440")
        self.root.resizable(False, False)
        
        # 程序信息
        self.version = "1.2"
        self.app_name = "Windows Cleaner"
        self.developer = "刘浩宇"
        self.rights_info = "本软件受著作权保护，未经授权禁止商用"
        
        # 状态变量
        self.total_files = 0
        self.deleted_files = 0
        self.is_cleaning = False
        self.clean_window = None  # 清理进度窗口
        self.path_window = None  # 路径管理窗口
        
        # 浏览器检测
        self.browser_installed = {
            "Chrome": self.check_chrome_installed(),
            "Edge": self.check_edge_installed()
        }

        # 清理路径配置
        self.cache_paths = {
            "系统临时文件": [os.environ.get("TEMP", ""),
                           os.path.join(os.environ.get("WINDIR", ""), "Temp")],
            "Chrome浏览器缓存": [os.path.join(os.path.expanduser("~"), 
                                           "AppData", "Local", "Google", "Chrome", 
                                           "User Data", "Default", "Cache")],
            "Edge浏览器缓存": [os.path.join(os.path.expanduser("~"), 
                                          "AppData", "Local", "Microsoft", "Edge", 
                                          "User Data", "Default", "Cache")],
            "回收站": "recycle_bin",
            "下载文件夹": [os.path.join(os.path.expanduser("~"), "Downloads")],
            "自定义路径": []
        }

        # 创建菜单栏
        self.create_menu()
        # 初始化UI
        self.setup_ui_styles()
        self.init_ui()

    def create_menu(self):
        """创建菜单栏"""
        menubar = tk.Menu(self.root)
        
        # 添加三个顶级菜单选项
        menubar.add_command(label="一键清理", command=self.start_cleaning)
        menubar.add_command(label="关于此程序", command=self.show_about)
        menubar.add_command(label="关闭程序", command=self.root.quit)
        
        self.root.config(menu=menubar)

    def check_chrome_installed(self):
        """检测Chrome是否安装"""
        paths = [
            os.path.join(os.environ.get("ProgramFiles", ""), "Google", "Chrome", "Application", "chrome.exe"),
            os.path.join(os.environ.get("ProgramFiles(x86)", ""), "Google", "Chrome", "Application", "chrome.exe"),
            os.path.join(os.path.expanduser("~"), "AppData", "Local", "Google", "Chrome", "Application", "chrome.exe")
        ]
        return any(os.path.exists(p) for p in paths)

    def check_edge_installed(self):
        """检测Edge是否安装"""
        paths = [
            os.path.join(os.environ.get("ProgramFiles", ""), "Microsoft", "Edge", "Application", "msedge.exe"),
            os.path.join(os.environ.get("ProgramFiles(x86)", ""), "Microsoft", "Edge", "Application", "msedge.exe"),
            os.path.join(os.path.expanduser("~"), "AppData", "Local", "Microsoft", "Edge", "Application", "msedge.exe")
        ]
        return any(os.path.exists(p) for p in paths)

    def setup_ui_styles(self):
        """设置UI样式"""
        style = ttk.Style()
        style.configure(".", font=("Microsoft YaHei", 10))
        style.configure("Title.TLabel", font=("Microsoft YaHei", 16, "bold"), foreground="#2c3e50")
        style.configure("TLabelFrame", font=("Microsoft YaHei", 11, "bold"), foreground="#34495e")
        style.configure("TProgressbar", thickness=10)
        style.configure("PathList.TLabel", font=("Microsoft YaHei", 9), foreground="#34495e")

    def init_ui(self):
        # 顶部标题区
        header_frame = ttk.Frame(self.root, padding=10)
        header_frame.pack(fill=tk.X, padx=10)
        ttk.Label(header_frame, text=self.app_name, style="Title.TLabel").pack(anchor=tk.CENTER)
        ttk.Label(header_frame, text="系统冗余文件清理工具 | 版本：1.2", foreground="#7f8c8d").pack(anchor=tk.CENTER, pady=5)

        # 清理项选择区
        options_frame = ttk.LabelFrame(self.root, text="清理选项", padding=15)
        options_frame.pack(fill=tk.X, padx=20, pady=10)

        self.check_vars = {}
        for name in self.cache_paths.keys():
            # 不显示未安装的浏览器选项
            if (name == "Chrome浏览器缓存" and not self.browser_installed["Chrome"]) or \
               (name == "Edge浏览器缓存" and not self.browser_installed["Edge"]):
                continue
                
            # 默认勾选状态：仅系统临时文件默认勾选
            default_value = True if name == "系统临时文件" else False
            var = tk.BooleanVar(value=default_value)
            self.check_vars[name] = var
            
            frame = ttk.Frame(options_frame)
            frame.pack(anchor=tk.W, pady=6, fill=tk.X)
            
            if name == "自定义路径":
                ttk.Checkbutton(frame, text=name, variable=var).pack(side=tk.LEFT, padx=2)
                ttk.Button(frame, text="设置", command=self.open_path_management).pack(side=tk.RIGHT, padx=5)
            else:
                ttk.Checkbutton(frame, text=f"● {name}", variable=var).pack(anchor=tk.W, padx=2)

        # 底部信息
        ttk.Label(self.root, text=f"© 2025 {self.app_name} | 开发者：{self.developer}", 
                 font=("Microsoft YaHei", 9), foreground="#95a5a6").pack(anchor=tk.CENTER, pady=20)

    def create_clean_window(self):
        """创建清理进度窗口"""
        # 关闭已有窗口
        if self.clean_window and isinstance(self.clean_window, tk.Toplevel) and self.clean_window.winfo_exists():
            self.clean_window.destroy()
            
        # 创建新窗口
        self.clean_window = tk.Toplevel(self.root)
        self.clean_window.title("清理进度")
        self.clean_window.geometry("500x150")
        self.clean_window.resizable(False, False)
        self.clean_window.transient(self.root)  # 依附于主窗口
        self.clean_window.grab_set()  # 模态窗口
        
        # 进度条
        progress_frame = ttk.Frame(self.clean_window, padding=15)
        progress_frame.pack(fill=tk.X, padx=20)
        self.progress_var = tk.DoubleVar()
        ttk.Progressbar(progress_frame, variable=self.progress_var, length=400).pack(fill=tk.X, pady=5)
        
        # 状态标签
        self.status_label = ttk.Label(progress_frame, text="准备清理...", foreground="#87CEEB", font=("Microsoft YaHei", 10))
        self.status_label.pack(anchor=tk.W, padx=2, pady=5)

    def show_about(self):
        """显示关于窗口"""
        about_window = tk.Toplevel(self.root)
        about_window.title("关于")
        about_window.geometry("300x200")
        about_window.resizable(False, False)
        about_window.transient(self.root)
        about_window.grab_set()
        
        # 关于内容
        ttk.Label(about_window, text=self.app_name, font=("Microsoft YaHei", 14, "bold")).pack(pady=10)
        ttk.Label(about_window, text=f"版本：{self.version}").pack(pady=5)
        ttk.Label(about_window, text=f"开发者：{self.developer}").pack(pady=5)
        ttk.Label(about_window, text=self.rights_info, font=("Microsoft YaHei", 9), foreground="#7f8c8d").pack(pady=10, padx=10)
        
        # 关闭按钮
        ttk.Button(about_window, text="关闭", command=about_window.destroy).pack(pady=10)

    def open_path_management(self):
        """打开路径管理窗口"""
        # 关闭已有窗口
        if self.path_window and isinstance(self.path_window, tk.Toplevel) and self.path_window.winfo_exists():
            self.path_window.destroy()
            
        # 创建新窗口
        self.path_window = tk.Toplevel(self.root)
        self.path_window.title("自定义路径管理")
        self.path_window.geometry("500x300")
        self.path_window.resizable(False, False)
        self.path_window.transient(self.root)  # 依附于主窗口
        self.path_window.grab_set()  # 模态窗口
        
        # 标题
        ttk.Label(self.path_window, text="已添加的自定义路径", font=("Microsoft YaHei", 11, "bold")).pack(pady=10)
        
        # 路径列表框架
        list_frame = ttk.Frame(self.path_window)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)
        
        # 滚动条
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 路径列表
        self.path_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, height=10, 
                                      font=("Microsoft YaHei", 9), selectbackground="#87CEEB")
        self.path_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.path_listbox.yview)
        
        # 刷新路径列表
        self.refresh_path_list()
        
        # 按钮框架
        btn_frame = ttk.Frame(self.path_window)
        btn_frame.pack(fill=tk.X, padx=20, pady=15)
        
        # 添加按钮
        ttk.Button(btn_frame, text="添加路径...", command=self.add_custom_path).pack(side=tk.LEFT, padx=5)
        
        # 删除按钮
        ttk.Button(btn_frame, text="删除选中", command=self.remove_selected_path).pack(side=tk.LEFT, padx=5)
        
        # 清空按钮
        ttk.Button(btn_frame, text="清空列表", command=self.clear_all_paths).pack(side=tk.LEFT, padx=5)
        
        # 关闭按钮
        ttk.Button(btn_frame, text="关闭", command=self.path_window.destroy).pack(side=tk.RIGHT, padx=5)

    def refresh_path_list(self):
        """刷新路径列表"""
        self.path_listbox.delete(0, tk.END)
        for path in self.cache_paths["自定义路径"]:
            self.path_listbox.insert(tk.END, path)

    def add_custom_path(self):
        """添加自定义路径"""
        try:
            path = filedialog.askdirectory(title="选择要清理的文件夹")
            if path:
                # 去除路径末尾的斜杠，统一格式
                path = path.rstrip(os.sep)
                if path not in self.cache_paths["自定义路径"]:
                    self.cache_paths["自定义路径"].append(path)
                    messagebox.showinfo("提示", f"已添加自定义清理路径：\n{path}")
                    self.refresh_path_list()
                else:
                    messagebox.showinfo("提示", "该路径已添加过")
        except Exception as e:
            messagebox.showerror("错误", f"添加路径失败：{str(e)}")

    def remove_selected_path(self):
        """删除选中的路径"""
        selected_indices = self.path_listbox.curselection()
        if not selected_indices:
            messagebox.showinfo("提示", "请先选择要删除的路径")
            return
            
        # 从后往前删除，避免索引错乱
        for i in sorted(selected_indices, reverse=True):
            del self.cache_paths["自定义路径"][i]
            
        self.refresh_path_list()
        messagebox.showinfo("提示", "已删除选中的路径")

    def clear_all_paths(self):
        """清空所有自定义路径"""
        if not self.cache_paths["自定义路径"]:
            messagebox.showinfo("提示", "路径列表已为空")
            return
            
        if messagebox.askyesno("确认", "确定要清空所有自定义路径吗？"):
            self.cache_paths["自定义路径"].clear()
            self.refresh_path_list()
            messagebox.showinfo("提示", "已清空所有自定义路径")

    def update_clean_status(self, text):
        """更新清理窗口状态"""
        if self.clean_window and self.clean_window.winfo_exists():
            self.status_label.config(text=text)
            self.clean_window.update_idletasks()  # 实时刷新

    def force_delete(self, path):
        """强制删除文件/文件夹"""
        try:
            if os.path.isfile(path):
                os.chmod(path, 0o777)
                os.remove(path)
                return True
            elif os.path.isdir(path):
                shutil.rmtree(path, ignore_errors=False, 
                             onerror=lambda e,f,p: os.chmod(f, 0o777))
                return True
        except:
            try:
                # 管理员权限删除
                ctypes.windll.shell32.ShellExecuteW(
                    None, "runas", "cmd.exe", 
                    f"/c del /f /q /s \"{path}\"", None, 0
                )
                return not os.path.exists(path)
            except:
                return False

    def clean_path(self, path, name):
        """清理指定路径"""
        if not os.path.exists(path):
            self.update_clean_status(f"路径不存在：{os.path.basename(path)}")
            return

        # 统计文件总数
        file_count = 0
        for root, _, files in os.walk(path):
            file_count += len(files)
        
        if file_count == 0:
            self.update_clean_status(f"{name} 中没有可清理的文件")
            return

        self.total_files += file_count
        current_deleted = 0
        self.update_clean_status(f"正在清理 {name}（共 {file_count} 个文件）")
        
        for root, _, files in os.walk(path):
            for file in files:
                file_path = os.path.join(root, file)
                if self.force_delete(file_path):
                    current_deleted += 1
                    self.deleted_files += 1
                    if current_deleted % 20 == 0:
                        self.update_clean_status(f"清理 {name}：{current_deleted}/{file_count} 个文件")
    def clean_recycle_bin(self):
        """清理回收站"""
        rbinfo = SHQUERYRBINFO()
        rbinfo.cbSize = ctypes.sizeof(SHQUERYRBINFO)
        shell32.SHQueryRecycleBinW(None, ctypes.byref(rbinfo))
        item_count = rbinfo.i64NumItems

        if item_count == 0:
            self.update_clean_status("回收站为空，无需清理")
            return True

        self.total_files += item_count
        self.update_clean_status(f"正在清理回收站（共 {item_count} 个项目）")
        
        try:
            result = shell32.SHEmptyRecycleBinW(
                None, None,
                SHERB_NOCONFIRMATION | SHERB_NOPROGRESSUI | SHERB_NOSOUND
            )
            if result == 0:
                self.deleted_files += item_count
                self.update_clean_status(f"回收站清理完成（删除 {item_count} 个项目）")
                return True
        except:
            pass
            
        self.update_clean_status("回收站清理失败，建议手动清空")
        return False

    def start_cleaning(self):
        """开始清理流程"""
        if self.is_cleaning:
            messagebox.showinfo("提示", "清理正在进行中，请稍候...")
            return
            
        # 检查危险清理项
        selected_items = [name for name, var in self.check_vars.items() if var.get()]
        if not selected_items:
            messagebox.showinfo("提示", "请至少选择一项清理内容")
            return
            
        warning_items = ["回收站", "下载文件夹", "自定义路径"]
        has_warning = any(item in selected_items for item in warning_items)
        
        if has_warning and not messagebox.askyesno(
            "警告", "您选择了高风险清理项，可能会删除重要文件。是否继续？"):
            return
            
        # 创建清理窗口
        self.create_clean_window()
        
        # 重置状态
        self.is_cleaning = True
        self.total_files = 0
        self.deleted_files = 0
        self.progress_var.set(0)
        
        # 启动清理线程
        threading.Thread(target=self.perform_cleaning, daemon=True).start()

    def perform_cleaning(self):
        """执行清理操作"""
        try:
            selected_items = [name for name, var in self.check_vars.items() if var.get()]
            total_items = len(selected_items)
            
            for idx, name in enumerate(selected_items):
                self.update_clean_status(f"正在准备清理：{name}")
                
                if name == "回收站":
                    self.clean_recycle_bin()
                else:
                    for path in self.cache_paths[name]:
                        self.clean_path(path, name)
                
                # 更新进度
                self.progress_var.set(((idx + 1) / total_items) * 100)

            # 清理完成
            self.update_clean_status(f"清理完成！共处理 {self.total_files} 个项目")
            # 延迟关闭窗口
            self.root.after(2000, self.clean_window.destroy)
            # 显示总结信息
            messagebox.showinfo("完成", f"清理完成！\n共处理 {self.total_files} 个项目\n成功删除 {self.deleted_files} 个项目")
            
        except Exception as e:
            self.update_clean_status(f"清理失败：{str(e)}")
            messagebox.showerror("错误", f"清理过程发生错误：{str(e)}")
        finally:
            self.is_cleaning = False
            # 确保窗口关闭
            if self.clean_window and self.clean_window.winfo_exists():
                self.root.after(3000, self.clean_window.destroy)

if __name__ == "__main__":
    root = tk.Tk()
    app = WindowsCleaner(root)
    # 管理员权限提示
    try:
        if not ctypes.windll.shell32.IsUserAnAdmin():
            messagebox.showinfo("权限提示", "建议以管理员身份运行以获得更好效果")
    except:
        pass
    root.mainloop()