import sys
import os
import re
import time
import win32com.client
import pythoncom
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QTextEdit, QSystemTrayIcon, QMenu)
from PyQt5.QtCore import QTimer, Qt
from PyQt5.QtGui import QIcon, QFont

class OutlookMaster(QWidget):
    def __init__(self):
        super().__init__()
        # --- 默认配置 ---
        self.share_dir = r'\\10.1.93.32\DT_HU_RDteam_F\视频\Z\ZOUQIU\paican'
        self.target_sender = 'zhoubaozhen1@huawei.com'
        self.scan_ms = 2000 
        
        self.init_ui()
        self.init_tray()
        
        # 定时器：轮询信号文件
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.poll_signal)
        self.timer.start(self.scan_ms)

    def init_ui(self):
        self.setWindowTitle("Outlook 同步管理端")
        self.resize(500, 400)
        self.setWindowFlags(Qt.WindowStaysOnTopHint) # 开启悬浮置顶
        
        layout = QVBoxLayout()
        
        # 路径配置
        layout.addWidget(QLabel("📂 共享保存目录 (SMB Path):"))
        self.edit_path = QLineEdit(self.share_dir)
        layout.addWidget(self.edit_path)

        # 抓件人配置
        layout.addWidget(QLabel("📧 监控发件人 (Sender Email):"))
        self.edit_sender = QLineEdit(self.target_sender)
        layout.addWidget(self.edit_sender)

        # 频率配置
        h_layout = QHBoxLayout()
        h_layout.addWidget(QLabel("⏱ 扫描频率(ms):"))
        self.edit_freq = QLineEdit(str(self.scan_ms))
        h_layout.addWidget(self.edit_freq)
        
        self.btn_apply = QPushButton("保存并应用配置")
        self.btn_apply.setStyleSheet("background: #0078d4; color: white; font-weight: bold;")
        self.btn_apply.clicked.connect(self.apply_settings)
        h_layout.addWidget(self.btn_apply)
        layout.addLayout(h_layout)

        # 日志区
        layout.addWidget(QLabel("📝 实时运行日志:"))
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setStyleSheet("background: #252525; color: #61FF61;")
        layout.addWidget(self.log_area)

        self.setLayout(layout)
        self.add_log("系统就绪，置顶监控已启动...")

    def init_tray(self):
        self.tray = QSystemTrayIcon(self)
        self.tray.setIcon(self.style().standardIcon(21)) # 默认邮件图标
        
        menu = QMenu()
        menu.addAction("打开主界面", self.showNormal)
        menu.addAction("彻底退出", QApplication.instance().quit)
        
        self.tray.setContextMenu(menu)
        self.tray.activated.connect(lambda r: self.showNormal() if r == QSystemTrayIcon.Trigger else None)
        self.tray.show()

    def closeEvent(self, event):
        """关闭窗口时隐藏到托盘"""
        self.hide()
        self.tray.showMessage("同步助手", "已转入后台运行", QSystemTrayIcon.Information, 1500)
        event.ignore()

    def add_log(self, text):
        self.log_area.append(f"[{time.strftime('%H:%M:%S')}] {text}")

    def apply_settings(self):
        self.share_dir = self.edit_path.text().strip()
        self.target_sender = self.edit_sender.text().strip()
        self.scan_ms = int(self.edit_freq.text())
        self.timer.start(self.scan_ms)
        self.add_log(f"✅ 配置更新：正在监控 {self.target_sender}")
        self.generate_html_index() # 立即刷新网页

    def poll_signal(self):
        sig_file = os.path.join(self.share_dir, "sync_request.txt")
        if os.path.exists(sig_file):
            self.add_log("📩 收到信号，执行同步...")
            self.run_sync()
            try: os.remove(sig_file)
            except: pass

    def run_sync(self):
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            ns = outlook.GetNamespace("MAPI")
            items = ns.GetDefaultFolder(6).Items
            items.Sort("[ReceivedTime]", True)
            
            for item in items:
                if item.Class == 43 and self.target_sender.lower() in item.SenderEmailAddress.lower():
                    subj = item.Subject
                    r_time = item.ReceivedTime.strftime("%Y%m%d_%H%M%S")
                    fname = f"{r_time}_{re.sub(r'[\\/:*?<>|]', '_', subj)[:50]}.html"
                    fpath = os.path.join(self.share_dir, fname)
                    
                    if not os.path.exists(fpath):
                        item.SaveAs(fpath, 4)
                        self.add_log(f"✅ 已存: {subj}")
                        if item.Attachments.Count > 0:
                            att_p = os.path.join(self.share_dir, f"{r_time}_附件")
                            if not os.path.exists(att_p): os.makedirs(att_p)
                            for i in range(1, item.Attachments.Count + 1):
                                item.Attachments.Item(i).SaveAsFile(os.path.join(att_p, item.Attachments.Item(i).FileName))
                    else:
                        self.add_log("ℹ️ 邮件重复，跳过。")
                    break
            self.generate_html_index()
        except Exception as e:
            self.add_log(f"❌ 错误: {e}")

    def generate_html_index(self):
        """生成供同事查看的 index.html"""
        files = [f for f in os.listdir(self.share_dir) if f.endswith('.html') and f != 'index.html']
        files.sort(reverse=True)
        html = f"""
        <html><head><meta charset='utf-8'><title>预览</title>
        <style>
            body {{ font-family: sans-serif; padding: 20px; background: #f3f2f1; }}
            .card {{ background: white; padding: 15px; margin: 10px 0; border-radius: 5px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); display: flex; align-items:center; }}
            .btn {{ background: #0078d4; color: white; padding: 10px; border:none; cursor:pointer; border-radius: 4px; }}
            a {{ flex:1; text-decoration: none; color: #0078d4; font-weight: bold; }}
        </style>
        <script>
            function sync() {{
                const a = document.createElement('a');
                a.href = window.URL.createObjectURL(new Blob(['s'], {{type:'text/plain'}}));
                a.download = 'sync_request.txt';
                a.click();
                alert('信号已发出，请保存到共享目录并等待刷新');
                setTimeout(()=>location.reload(), 5000);
            }}
        </script>
        </head><body>
            <h2>📩 邮件列表 (来自: {self.target_sender})</h2>
            <button class="btn" onclick="sync()">🔄 一键同步最新邮件</button><hr>
        """
        for f in files:
            html += f'<div class="card"><span>[{f[:15]}] </span><a href="{f}" target="_blank">{f[16:-5]}</a></div>'
        html += "</body></html>"
        with open(os.path.join(self.share_dir, "index.html"), "w", encoding="utf-8") as f:
            f.write(html)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    QApplication.setQuitOnLastWindowClosed(False)
    m = OutlookMaster()
    m.show()
    sys.exit(app.exec_())
