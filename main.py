# Author: Sean
# Email: fishis@126.com

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
from smb.SMBConnection import SMBConnection
from nmb.NetBIOS import NetBIOS
import socket
import os
import platform
import tempfile
import subprocess

import os
import platform
import tempfile
import subprocess
import sys

import json

class SMBBrowserApp:
    def __init__(self, root):
        self.root = root
        self.root.title("云铠 SMB 浏览器 1.0")
        self.root.geometry("800x600")
        
        # Style
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        self.conn = None
        self.current_share = None
        self.current_path = ""
        self.file_list = []
        self.config_file = "config.json"
        
        # For thread safety in UI updates
        self.lock = threading.Lock()
        
        self.setup_ui()
        self.setup_menu()
        self.load_config()

    def setup_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="帮助", menu=help_menu)
        help_menu.add_command(label="关于", command=self.show_about)

    def show_about(self):
        # Create a custom TopLevel window
        about_window = tk.Toplevel(self.root)
        about_window.title("关于")
        about_window.geometry("320x480")
        about_window.resizable(False, False)
        about_window.configure(bg='white')  # White background
        
        # Center the window
        self.root.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 160
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 240
        about_window.geometry(f"+{x}+{y}")

        # Main container with white background
        container = tk.Frame(about_window, bg='white', padx=20, pady=20)
        container.pack(fill=tk.BOTH, expand=True)

        # Title
        tk.Label(container, text="云铠 SMB 浏览器", font=("Helvetica", 18, "bold"), bg='white', fg='#333333').pack(pady=(10, 5))
        tk.Label(container, text="v1.0", font=("Helvetica", 12), bg='white', fg='#888888').pack(pady=(0, 20))
        
        # Info card
        info_frame = tk.Frame(container, bg='#f5f5f5', padx=15, pady=10) # Light gray bg for info
        info_frame.pack(fill=tk.X, pady=(0, 20))
        
        tk.Label(info_frame, text="作者：Sean", font=("Helvetica", 11), bg='#f5f5f5', fg='#333333').pack(anchor='w')
        tk.Label(info_frame, text="邮箱：fishis@126.com", font=("Helvetica", 11), bg='#f5f5f5', fg='#333333').pack(anchor='w')
        
        # QR Code Image Section
        try:
            # Use resource_path helper for PyInstaller compatibility
            def resource_path(relative_path):
                """ Get absolute path to resource, works for dev and for PyInstaller """
                try:
                    # PyInstaller creates a temp folder and stores path in _MEIPASS
                    base_path = sys._MEIPASS
                except Exception:
                    base_path = os.path.abspath(".")
                return os.path.join(base_path, relative_path)

            img_path = resource_path("wechat_qr.png")
            
            if os.path.exists(img_path):
                # Load image
                original_image = tk.PhotoImage(file=img_path)
                
                # Auto-resize logic (simple subsample)
                w = original_image.width()
                h = original_image.height()
                
                # Target width ~200px
                target_w = 200
                if w > target_w:
                    factor = int(w / target_w)
                    if factor < 1: factor = 1
                    about_window.image = original_image.subsample(factor, factor)
                else:
                    about_window.image = original_image
                
                img_label = tk.Label(container, image=about_window.image, bg='white')
                img_label.pack(pady=5)
                
                tk.Label(container, text="扫一扫添加微信", font=("Helvetica", 10), bg='white', fg='#666666').pack(pady=(5, 0))
            else:
                 tk.Label(container, text="(未找到名片图片)", bg='white', fg='#999').pack(pady=20)
                
        except Exception as e:
            print(f"Error loading image: {e}")
            tk.Label(container, text="(图片加载失败: 格式不支持?)", bg='white', fg='red').pack(pady=20)
            
        # Close button styled
        tk.Button(container, text="关闭", command=about_window.destroy, 
                  highlightbackground='white').pack(side=tk.BOTTOM, pady=(10, 0))

    def setup_ui(self):
        # Top Frame: Connection Settings
        top_frame = ttk.Frame(self.root, padding="10")
        top_frame.pack(fill=tk.X)
        
        # Grid layout for top frame
        ttk.Label(top_frame, text="服务器地址:").grid(row=0, column=0, padx=5, sticky=tk.W)
        self.server_ip = tk.StringVar()
        ttk.Entry(top_frame, textvariable=self.server_ip, width=15).grid(row=0, column=1, padx=5)
        
        ttk.Label(top_frame, text="端口:").grid(row=0, column=2, padx=5, sticky=tk.W)
        self.port = tk.StringVar(value="445")
        ttk.Entry(top_frame, textvariable=self.port, width=6).grid(row=0, column=3, padx=5)

        ttk.Label(top_frame, text="用户:").grid(row=0, column=4, padx=5, sticky=tk.W)
        self.username = tk.StringVar(value="guest")
        ttk.Entry(top_frame, textvariable=self.username, width=12).grid(row=0, column=5, padx=5)
        
        ttk.Label(top_frame, text="密码:").grid(row=0, column=6, padx=5, sticky=tk.W)
        self.password = tk.StringVar()
        ttk.Entry(top_frame, textvariable=self.password, show="*", width=12).grid(row=0, column=7, padx=5)
        
        self.connect_btn = ttk.Button(top_frame, text="连接", command=self.start_connect_thread)
        self.connect_btn.grid(row=0, column=8, padx=10)

        # Middle Frame: File Browser
        mid_frame = ttk.Frame(self.root, padding="10")
        mid_frame.pack(fill=tk.BOTH, expand=True)

        # Toolbar
        toolbar = ttk.Frame(mid_frame)
        toolbar.pack(fill=tk.X, pady=(0, 5))
        
        self.back_btn = ttk.Button(toolbar, text="返回", state=tk.DISABLED, command=self.go_back)
        self.back_btn.pack(side=tk.LEFT)
        
        self.path_label = ttk.Label(toolbar, text="未连接", anchor="w")
        self.path_label.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)

        # Treeview for files
        columns = ("Size", "Type")
        self.tree = ttk.Treeview(mid_frame, columns=columns, show="tree headings")
        self.tree.heading("#0", text="名称", anchor="w")
        self.tree.heading("Size", text="大小")
        self.tree.heading("Type", text="类型")
        
        self.tree.column("#0", width=400)
        self.tree.column("Size", width=100)
        self.tree.column("Type", width=100)
        
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(mid_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscroll=scrollbar.set)
        
        self.tree.bind("<Double-1>", self.on_double_click)

        # Bottom Frame: Actions
        bottom_frame = ttk.Frame(self.root, padding="10")
        bottom_frame.pack(fill=tk.X)
        
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(bottom_frame, textvariable=self.status_var).pack(side=tk.LEFT)
        
        self.download_btn = ttk.Button(bottom_frame, text="下载选中文件", state=tk.DISABLED, command=self.download_file)
        self.download_btn.pack(side=tk.RIGHT)

    def load_config(self):
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                    self.server_ip.set(config.get("ip", ""))
                    self.port.set(config.get("port", "445"))
                    self.username.set(config.get("user", "guest"))
                    # Password is NOT saved for security, or can be if requested (keeping it safe for now)
        except Exception as e:
            print(f"Failed to load config: {e}")

    def save_config(self):
        try:
            config = {
                "ip": self.server_ip.get(),
                "port": self.port.get(),
                "user": self.username.get()
            }
            with open(self.config_file, 'w') as f:
                json.dump(config, f)
        except Exception as e:
            print(f"Failed to save config: {e}")

    def update_status(self, msg):
        self.root.after(0, lambda: self.status_var.set(msg))

    def show_error(self, title, msg):
        self.root.after(0, lambda: messagebox.showerror(title, msg))

    def start_connect_thread(self):
        self.connect_btn.config(state=tk.DISABLED)
        self.update_status("正在连接...")
        threading.Thread(target=self.connect, daemon=True).start()

    def connect(self):
        addr_input = self.server_ip.get().strip()
        user_port_str = self.port.get().strip()
        user = self.username.get().strip()
        password = self.password.get().strip()
        
        if not addr_input:
            self.show_error("错误", "请输入服务器地址")
            self.update_status("就绪")
            self.root.after(0, lambda: self.connect_btn.config(state=tk.NORMAL))
            return

        # Resolve hostname to IP if needed
        real_ip = addr_input
        try:
            # Check if it's already an IP
            socket.inet_aton(addr_input)
        except socket.error:
            # Not an IP, try to resolve
            try:
                self.update_status(f"正在解析主机名 {addr_input}...")
                real_ip = socket.gethostbyname(addr_input)
                self.update_status(f"主机名已解析: {addr_input} -> {real_ip}")
            except Exception as e:
                 print(f"DNS Resolution failed: {e}")
                 # Continue anyway, pysmb/socket might handle it or fail later
                 pass

        ports_to_try = []
        if user_port_str:
            p = int(user_port_str)
            ports_to_try.append(p)
            # If user left it as default 445, also try 139 as fallback
            if p == 445:
                ports_to_try.append(139)
        else:
            ports_to_try = [445, 139]

        client_name = socket.gethostname()
        
        # Try to resolve NetBIOS name once
        remote_name = "*SMBSERVER"
        try:
            # Use the resolved IP for NetBIOS query
            self.update_status(f"正在解析 NetBIOS 名称 {real_ip}...")
            nb = NetBIOS()
            resolved = nb.queryIPForName(real_ip, port=137, timeout=2)
            nb.close()
            if resolved:
                remote_name = resolved[0]
        except:
            pass # Ignore resolution errors, use *SMBSERVER

        success = False
        last_error = None

        for port in ports_to_try:
            try:
                self.update_status(f"正在尝试连接 {real_ip}:{port}...")
                
                is_direct = (port == 445)
                self.conn = SMBConnection(
                    user, 
                    password, 
                    client_name, 
                    remote_name, 
                    use_ntlm_v2=True,
                    sign_options=2,
                    is_direct_tcp=is_direct
                )
                
                if self.conn.connect(real_ip, port, timeout=5):
                    success = True
                    self.update_status(f"已连接到 {real_ip} 端口 {port}")
                    
                    # Update the UI port to show what actually worked
                    self.root.after(0, lambda p=port: self.port.set(str(p)))
                    break
                else:
                    last_error = "认证失败"
            except Exception as e:
                last_error = str(e)
                print(f"Failed on port {port}: {e}")
                self.conn = None
        
        if success:
            # Save successful connection details
            self.save_config()
            try:
                self.update_status("正在列出共享...")
                shares = self.conn.listShares()
                self.root.after(0, lambda: self.show_shares(shares))
                
                # Determine protocol version for display
                protocol_ver = "SMB2/3" if self.conn.isUsingSMB2 else "SMB1"
                self.update_status(f"已连接到 {real_ip} (协议: {protocol_ver})")
            except Exception as e:
                self.show_error("列出共享错误", str(e))
                self.update_status("已连接 (获取列表失败)")
        else:
            msg = last_error if last_error else "连接失败"
            self.show_error("连接错误", f"无法连接到 {real_ip}.\n已尝试端口: {ports_to_try}\n错误: {msg}")
            self.update_status("连接失败")
            self.conn = None
        
        self.root.after(0, lambda: self.connect_btn.config(state=tk.NORMAL))

    def show_shares(self, shares):
        self.current_share = None
        self.current_path = ""
        self.path_label.config(text=f"\\\\{self.server_ip.get()}")
        self.back_btn.config(state=tk.DISABLED)
        self.download_btn.config(state=tk.DISABLED)
        
        # Clear tree
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        for share in shares:
            # Special shares often usually start with $
            if not share.isSpecial:
                self.tree.insert("", "end", text=share.name, values=("共享文件夹", "文件夹"), iid=share.name)
        
    def on_double_click(self, event):
        item_id = self.tree.selection()[0]
        item = self.tree.item(item_id)
        name = item['text']
        i_type = item['values'][1]
        
        if i_type == "文件夹" or self.current_share is None:
            self.enter_directory(name)
        else:
            self.open_file(name)

    def open_file(self, filename):
        threading.Thread(target=self.perform_file_open, args=(filename,), daemon=True).start()

    def perform_file_open(self, filename):
        try:
            self.update_status(f"正在准备预览 {filename}...")
            
            path_to_file = filename
            if self.current_path:
                path_to_file = f"{self.current_path}/{filename}"
            
            # 使用临时目录的子目录作为缓存
            temp_dir = tempfile.gettempdir()
            cache_dir = os.path.join(temp_dir, "smb_browser_cache")
            if not os.path.exists(cache_dir):
                os.makedirs(cache_dir, exist_ok=True)
            
            save_path = os.path.join(cache_dir, filename)
            
            # 下载文件
            with open(save_path, 'wb') as f:
                self.conn.retrieveFile(self.current_share, path_to_file, f)
            
            self.update_status(f"正在打开 {filename}...")
            
            # 根据系统打开文件
            system_name = platform.system()
            if system_name == 'Darwin':       # macOS
                subprocess.run(['open', save_path])
            elif system_name == 'Windows':    # Windows
                os.startfile(save_path)
            else:                             # Linux/Unix
                subprocess.run(['xdg-open', save_path])
                
            self.update_status(f"已打开 {filename}")
            
        except Exception as e:
            self.show_error("预览错误", f"无法打开文件: {str(e)}")
            self.update_status("预览失败")

    def enter_directory(self, name):
        if self.current_share is None:
            self.current_share = name
            self.current_path = ""
        else:
            if self.current_path:
                 self.current_path = f"{self.current_path}/{name}"
            else:
                 self.current_path = name

        self.update_status(f"正在列出 {self.current_path}...")
        threading.Thread(target=self.list_files, daemon=True).start()

    def list_files(self):
        try:
            files = self.conn.listPath(self.current_share, self.current_path)
            self.root.after(0, lambda: self.update_file_list(files))
        except Exception as e:
            self.show_error("列出文件错误", str(e))
            # Revert path change if failed
            # (Simple logic: just don't update list for now)
            self.update_status("列出文件失败")

    def update_file_list(self, files):
        # Update breadcrumb
        display_path = f"\\\\{self.server_ip.get()}\\{self.current_share}"
        if self.current_path:
             path_backslashes = self.current_path.replace('/', '\\')
             display_path += f"\\{path_backslashes}"
        self.path_label.config(text=display_path)
        
        self.back_btn.config(state=tk.NORMAL)
        self.download_btn.config(state=tk.NORMAL)
        
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        for f in files:
            if f.filename in ['.', '..']:
                continue
            
            ftype = "文件夹" if f.isDirectory else "文件"
            size = f"{f.file_size / 1024:.1f} KB" if not f.isDirectory else ""
            
            # Simple icon differentiation by type text
            self.tree.insert("", "end", text=f.filename, values=(size, ftype))

    def go_back(self):
        if not self.current_share:
            return
            
        if not self.current_path:
            # Go back to share list
            self.update_status("正在返回共享列表...")
            threading.Thread(target=self.refresh_shares, daemon=True).start()
        else:
            # Go up one directory
            if '/' in self.current_path:
                self.current_path = self.current_path.rsplit('/', 1)[0]
            else:
                self.current_path = ""
            
            self.update_status(f"正在列出 {self.current_path}...")
            threading.Thread(target=self.list_files, daemon=True).start()

    def refresh_shares(self):
        try:
            shares = self.conn.listShares()
            self.root.after(0, lambda: self.show_shares(shares))
        except Exception as e:
             self.show_error("连接错误", str(e))

    def download_file(self):
        selected_items = self.tree.selection()
        if not selected_items:
            return

        # 收集所有选中的文件（忽略文件夹）
        files_to_download = []
        has_folder = False
        for iid in selected_items:
            item = self.tree.item(iid)
            # values 是一个列表，第二个元素是类型
            if item['values'][1] == "文件夹":
                has_folder = True
            else:
                files_to_download.append(item['text'])
        
        if not files_to_download:
            msg = "未选择文件。"
            if has_folder:
                msg += " (不支持下载文件夹)"
            messagebox.showinfo("提示", msg)
            return

        # 如果只包含一个文件，行为和以前一样（可以选择保存名）
        if len(files_to_download) == 1:
            filename = files_to_download[0]
            save_path = filedialog.asksaveasfilename(initialfile=filename)
            if save_path:
                threading.Thread(target=self.perform_download_single, args=(filename, save_path), daemon=True).start()
        else:
            # 如果包含多个文件，选择目录批量下载
            if has_folder:
                # 提示用户只有文件会被下载
                if not messagebox.askyesno("提示", "选中的项目中包含文件夹，文件夹将被忽略。是否继续下载其余文件？"):
                    return
            
            target_dir = filedialog.askdirectory()
            if target_dir:
                threading.Thread(target=self.perform_download_batch, args=(files_to_download, target_dir), daemon=True).start()

    def perform_download_single(self, filename, save_path):
        try:
            self.update_status(f"正在下载 {filename}...")
            path_to_file = filename
            if self.current_path:
                path_to_file = f"{self.current_path}/{filename}"
                
            with open(save_path, 'wb') as f:
                self.conn.retrieveFile(self.current_share, path_to_file, f)
            
            self.update_status(f"已下载 {filename}")
            self.root.after(0, lambda: messagebox.showinfo("成功", f"文件已下载: {filename}"))
        except Exception as e:
            self.show_error("下载错误", str(e))
            self.update_status("下载失败")

    def perform_download_batch(self, files, target_dir):
        success_count = 0
        total = len(files)
        errors = []

        for index, filename in enumerate(files, 1):
            try:
                self.update_status(f"正在下载 ({index}/{total}): {filename}...")
                path_to_file = filename
                if self.current_path:
                    path_to_file = f"{self.current_path}/{filename}"
                
                save_path = os.path.join(target_dir, filename)
                
                with open(save_path, 'wb') as f:
                    self.conn.retrieveFile(self.current_share, path_to_file, f)
                
                success_count += 1
            except Exception as e:
                errors.append(f"{filename}: {str(e)}")
                print(f"Error downloading {filename}: {e}")

        # 最终状态
        status_msg = f"批量下载完成。成功: {success_count}/{total}"
        if errors:
             status_msg += f" 失败: {len(errors)}"
        
        self.update_status(status_msg)
        
        report = f"下载完成。\n成功: {success_count}\n失败: {len(errors)}"
        if errors:
            report += "\n\n错误详情 (前5个):\n" + "\n".join(errors[:5])
        
        self.root.after(0, lambda: messagebox.showinfo("下载报告", report))

if __name__ == "__main__":
    root = tk.Tk()
    app = SMBBrowserApp(root)
    root.mainloop()
