import sys
import os
import re
import time
import win32com.client
import pythoncom
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QTextEdit, QSystemTrayIcon, QMenu)
from PyQt5.QtCore import QTimer, Qt

class OutlookMaster(QWidget):
    def __init__(self):
        super().__init__()
        # --- 初始配置 ---
        self.share_dir = r'\\10.1.93.32\DT_HU_RDteam_F\视频\Z\ZOUQIU\paican'
        self.target_sender = 'zhoubaozhen1@huawei.com'
        self.scan_ms = 2000 
        self.target_item = None # 搜索缓存
        
        self.init_ui()
        self.init_tray()
        
        # 定时器：轮询信号文件
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.poll_signal)
        self.timer.start(self.scan_ms)

    def init_ui(self):
        self.setWindowTitle("Outlook 同步助手 Pro")
        self.resize(500, 450)
        self.setWindowFlags(Qt.WindowStaysOnTopHint) # 悬浮置顶
        
        layout = QVBoxLayout()
        
        # 路径配置
        layout.addWidget(QLabel("📂 共享保存目录:"))
        self.edit_path = QLineEdit(self.share_dir)
        layout.addWidget(self.edit_path)

        # 抓件人配置
        layout.addWidget(QLabel("📧 监控发件人 (建议使用中文名或工号):"))
        self.edit_sender = QLineEdit(self.target_sender)
        layout.addWidget(self.edit_sender)

        # 频率配置
        h_layout = QHBoxLayout()
        h_layout.addWidget(QLabel("⏱ 扫描频率(ms):"))
        self.edit_freq = QLineEdit(str(self.scan_ms))
        h_layout.addWidget(self.edit_freq)
        
        self.btn_apply = QPushButton("保存并应用配置")
        self.btn_apply.setStyleSheet("background: #0078d4; color: white; font-weight: bold; padding: 6px;")
        self.btn_apply.clicked.connect(self.apply_settings)
        h_layout.addWidget(self.btn_apply)
        layout.addLayout(h_layout)

        # 日志区
        layout.addWidget(QLabel("📝 运行日志:"))
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setStyleSheet("background: #1e1e1e; color: #00ff00;")
        layout.addWidget(self.log_area)

        self.setLayout(layout)
        self.add_log("系统就绪，请确保 Outlook 已打开。")

    def init_tray(self):
        self.tray = QSystemTrayIcon(self)
        self.tray.setIcon(self.style().standardIcon(21)) 
        menu = QMenu()
        menu.addAction("显示主界面", self.showNormal)
        menu.addAction("彻底退出", QApplication.instance().quit)
        self.tray.setContextMenu(menu)
        self.tray.activated.connect(lambda r: self.showNormal() if r == QSystemTrayIcon.Trigger else None)
        self.tray.show()

    def closeEvent(self, event):
        self.hide()
        self.tray.showMessage("同步助手", "已转入后台运行", QSystemTrayIcon.Information, 1500)
        event.ignore()

    def add_log(self, text):
        self.log_area.append(f"[{time.strftime('%H:%M:%S')}] {text}")

    def apply_settings(self):
        self.share_dir = self.edit_path.text().strip()
        self.target_sender = self.edit_sender.text().strip()
        try:
            self.scan_ms = int(self.edit_freq.text())
            self.timer.start(self.scan_ms)
            self.add_log(f"✅ 配置已生效：正在监控 '{self.target_sender}'")
            self.generate_html_index() 
        except:
            self.add_log("❌ 频率必须为数字")

    def poll_signal(self):
        sig_file = os.path.join(self.share_dir, "sync_request.txt")
        if os.path.exists(sig_file):
            self.add_log("📩 收到同步请求，开始执行...")
            self.run_sync()
            try: os.remove(sig_file)
            except: pass

    def run_sync(self):
        """强力兼容版同步逻辑"""
        try:
            pythoncom.CoInitialize()
            outlook = None
            try:
                outlook = win32com.client.GetActiveObject("Outlook.Application")
            except:
                try:
                    outlook = win32com.client.DispatchEx("Outlook.Application")
                except Exception as e:
                    self.add_log(f"❌ 无法连接到 Outlook: {str(e)}")
                    return

            ns = outlook.GetNamespace("MAPI")
            # 强制挂载当前会话
            ns.Logon("", "", False, True) 
            
            target = self.target_sender.lower().strip()
            self.add_log(f"🔎 正在全盘搜索关键字: {target}")
            
            self.target_item = None
            # 暴力扫描所有文件夹
            for folder in ns.Folders:
                self._recursive_search(folder, target)
                if self.target_item: break

            if self.target_item:
                subj = self.target_item.Subject
                r_time = self.target_item.ReceivedTime.strftime("%Y%m%d_%H%M%S")
                # 正则清洗，兼容 3.9
                clean_subj = re.sub(r'[\\/:*?<>|]', '_', subj)[:50]
                fname = f"{r_time}_{clean_subj}.html"
                fpath = os.path.join(self.share_dir, fname)
                
                if not os.path.exists(fpath):
                    self.target_item.SaveAs(fpath, 4) # olHTML
                    self.add_log(f"✅ 抓取成功: {subj}")
                    # 处理附件
                    if self.target_item.Attachments.Count > 0:
                        att_p = os.path.join(self.share_dir, f"{r_time}_附件")
                        if not os.path.exists(att_p): os.makedirs(att_p)
                        for i in range(1, self.target_item.Attachments.Count + 1):
                            self.target_item.Attachments.Item(i).SaveAsFile(os.path.join(att_p, self.target_item.Attachments.Item(i).FileName))
                else:
                    self.add_log(f"ℹ️ 邮件已同步过: {subj}")
            else:
                self.add_log(f"❌ 未发现包含 '{target}' 的邮件")
            
            self.generate_html_index()
        except Exception as e:
            self.add_log(f"❌ 运行异常: {str(e)}")
        finally:
            pythoncom.CoUninitialize()

    def _recursive_search(self, folder, target):
        """递归搜索辅助函数"""
        if self.target_item: return
        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)
            
            count = 0
            for item in items:
                if count > 30: break # 每个文件夹扫最新30封
                count += 1
                if item.Class == 43:
                    s_addr = str(getattr(item, 'SenderEmailAddress', '')).lower()
                    s_name = str(getattr(item, 'SenderName', '')).lower()
                    if target in s_addr or target in s_name:
                        self.target_item = item
                        return
            
            for sub in folder.Folders:
                self._recursive_search(sub, target)
                if self.target_item: return
        except:
            pass

    def generate_html_index(self):
        if not os.path.exists(self.share_dir): return
        files = [f for f in os.listdir(self.share_dir) if f.endswith('.html') and f != 'index.html']
        files.sort(reverse=True)
        html = f"""
        <html><head><meta charset='utf-8'><title>邮件预览</title>
        <style>
            body {{ font-family: sans-serif; padding: 20px; background: #f3f2f1; }}
            .header {{ background: #0078d4; color: white; padding: 15px; border-radius: 5px; margin-bottom: 20px; }}
            .card {{ background: white; padding: 15px; margin: 10px 0; border-radius: 5px; border-left: 5px solid #0078d4; display: flex; align-items:center; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }}
            .btn {{ background: white; color: #0078d4; padding: 10px 20px; border:none; cursor:pointer; border-radius: 4px; font-weight: bold; }}
            a {{ flex:1; text-decoration: none; color: #333; font-weight: 600; margin-left: 15px; }}
            a:hover {{ color: #0078d4; }}
        </style>
        <script>
            function sync() {{
                const b = new Blob(['s'], {{type:'text/plain'}});
                const a = document.createElement('a');
                a.href = window.URL.createObjectURL(b);
                a.download = 'sync_request.txt';
                a.click();
                alert('信号已发出，请点击保存到共享文件夹并等待刷新。');
                setTimeout(()=>location.reload(), 5000);
            }}
        </script>
        </head><body>
            <div class="header">
                <h2 style="margin:0;">📩 邮件管理列表</h2>
                <p>监控发件人: {self.target_sender}</p>
                <button class="btn" onclick="sync()">🔄 网页端一键同步</button>
            </div>
        """
        for f in files:
            html += f'<div class="card"><small style="color:#888">{f[:15]}</small><a href="{f}" target="_blank">{f[16:-5]}</a></div>'
        html += "</body></html>"
        with open(os.path.join(self.share_dir, "index.html"), "w", encoding="utf-8") as f:
            f.write(html)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    QApplication.setQuitOnLastWindowClosed(False)
    m = OutlookMaster()
    m.show()
    sys.exit(app.exec_())
