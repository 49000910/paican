import sys
import os
import re
import time
import win32com.client
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
        
        self.init_ui()
        self.init_tray()
        
        # 定时器：轮询信号文件
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.poll_signal)
        self.timer.start(self.scan_ms)

    def init_ui(self):
        self.setWindowTitle("Outlook 同步管理端 Pro")
        self.resize(500, 450)
        self.setWindowFlags(Qt.WindowStaysOnTopHint) # 悬浮置顶
        
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
        self.btn_apply.setStyleSheet("background: #0078d4; color: white; font-weight: bold; padding: 5px;")
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
        self.add_log("系统已就绪，正在实时监控信号...")

    def init_tray(self):
        self.tray = QSystemTrayIcon(self)
        # 使用系统默认图标
        self.tray.setIcon(self.style().standardIcon(21)) 
        
        menu = QMenu()
        menu.addAction("打开主界面", self.showNormal)
        menu.addAction("彻底退出", QApplication.instance().quit)
        
        self.tray.setContextMenu(menu)
        self.tray.activated.connect(self.on_tray_activated)
        self.tray.show()

    def on_tray_activated(self, reason):
        if reason == QSystemTrayIcon.Trigger:
            self.showNormal()

    def closeEvent(self, event):
        """点击关闭按钮时隐藏到系统托盘"""
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
            self.add_log(f"✅ 配置更新：正在监控 {self.target_sender}")
            self.generate_html_index() 
        except:
            self.add_log("❌ 频率请输入数字")

    def poll_signal(self):
        sig_file = os.path.join(self.share_dir, "sync_request.txt")
        if os.path.exists(sig_file):
            self.add_log("📩 收到信号，开始抓取邮件...")
            self.run_sync()
            try:
                os.remove(sig_file)
                self.add_log("🗑️ 信号处理完毕。")
            except Exception as e:
                self.add_log(f"⚠ 无法删除信号文件: {e}")

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
                    
                    # 修复：将正则提取出 f-string，兼容 Python 3.9
                    clean_subj = re.sub(r'[\\/:*?<>|]', '_', subj)[:50]
                    fname = f"{r_time}_{clean_subj}.html"
                    fpath = os.path.join(self.share_dir, fname)
                    
                    if not os.path.exists(fpath):
                        item.SaveAs(fpath, 4)
                        self.add_log(f"✅ 已存新邮件: {subj}")
                        if item.Attachments.Count > 0:
                            att_p = os.path.join(self.share_dir, f"{r_time}_附件")
                            if not os.path.exists(att_p): os.makedirs(att_p)
                            for i in range(1, item.Attachments.Count + 1):
                                att = item.Attachments.Item(i)
                                att.SaveAsFile(os.path.join(att_p, att.FileName))
                    else:
                        self.add_log("ℹ 邮件已存在，跳过。")
                    break
            self.generate_html_index()
        except Exception as e:
            self.add_log(f"❌ Outlook 错误: {e}")

    def generate_html_index(self):
        """同步本地配置到共享盘的预览页面"""
        if not os.path.exists(self.share_dir): return
        files = [f for f in os.listdir(self.share_dir) if f.endswith('.html') and f != 'index.html']
        files.sort(reverse=True)
        
        html = f"""
        <html><head><meta charset='utf-8'><title>邮件预览</title>
        <style>
            body {{ font-family: 'Segoe UI', sans-serif; padding: 20px; background: #f3f2f1; }}
            .header {{ background: #0078d4; color: white; padding: 15px; border-radius: 5px; margin-bottom: 20px; }}
            .card {{ background: white; padding: 15px; margin: 10px 0; border-radius: 5px; border-left: 5px solid #0078d4; display: flex; align-items:center; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }}
            .btn {{ background: white; color: #0078d4; padding: 10px 20px; border:none; cursor:pointer; border-radius: 4px; font-weight: bold; }}
            a {{ flex:1; text-decoration: none; color: #333; font-weight: 600; margin-left: 15px; }}
            a:hover {{ color: #0078d4; }}
        </style>
        <script>
            function sync() {{
                const blob = new Blob(['sync'], {{type:'text/plain'}});
                const a = document.createElement('a');
                a.href = window.URL.createObjectURL(blob);
                a.download = 'sync_request.txt';
                a.click();
                alert('同步请求已发出，请保存到共享目录并等待几秒后刷新页面。');
                setTimeout(()=>location.reload(), 5000);
            }}
        </script>
        </head><body>
            <div class="header">
                <h2 style="margin:0;">📩 邮件同步列表</h2>
                <p style="margin:5px 0 15px;">当前监控: {self.target_sender}</p>
                <button class="btn" onclick="sync()">🔄 网页端一键抓取最新</button>
            </div>
        """
        for f in files:
            html += f'<div class="card"><small style="color:#888">{f[:15]}</small><a href="{f}" target="_blank">{f[16:-5]}</a></div>'
        html += "</body></html>"
        
        with open(os.path.join(self.share_dir, "index.html"), "w", encoding="utf-8") as f:
            f.write(html)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    # 保证即使主窗口关闭，进程也在托盘运行
    QApplication.setQuitOnLastWindowClosed(False)
    m = OutlookMaster()
    m.show()
    sys.exit(app.exec_())
