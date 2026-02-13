import sys, os, re, time, subprocess
import win32com.client
import pythoncom
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QTextEdit, QSystemTrayIcon, QMenu)
from PyQt5.QtCore import QTimer, Qt
from PyQt5.QtGui import QIcon

class OutlookMasterPro(QWidget):
    def __init__(self):
        super().__init__()
        # --- 初始默认参数 ---
        self.share_dir = r'\\10.1.93.32\DT_HU_RDteam_F\视频\Z\ZOUQIU\paican'
        self.target_kw = 'EDFA' 
        self.copyright_text = "© 2024-2026 RD Team | 视频组技术支持"
        self.scan_ms = 2000 
        
        self.init_ui()
        self.init_tray()
        
        # 启动定时器监听信号
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.poll_signal)
        self.timer.start(self.scan_ms)
        
        # 启动时先刷新一次网页
        self.update_web_ui()

    def init_ui(self):
        self.setWindowTitle("Outlook 同步管理端 - 后台版")
        self.resize(550, 450)
        self.setWindowFlags(Qt.WindowStaysOnTopHint) # 悬浮置顶
        
        layout = QVBoxLayout()
        
        # 配置区
        layout.addWidget(QLabel("📂 共享保存目录:"))
        self.edit_path = QLineEdit(self.share_dir)
        layout.addWidget(self.edit_path)

        layout.addWidget(QLabel("📧 监控关键词 (发件人/标题):"))
        self.edit_kw = QLineEdit(self.target_kw)
        layout.addWidget(self.edit_kw)

        layout.addWidget(QLabel("📝 网页底部版权信息:"))
        self.edit_copy = QLineEdit(self.copyright_text)
        layout.addWidget(self.edit_copy)

        h_layout = QHBoxLayout()
        h_layout.addWidget(QLabel("⏱ 扫描频率(ms):"))
        self.edit_freq = QLineEdit(str(self.scan_ms))
        h_layout.addWidget(self.edit_freq)
        
        self.btn_apply = QPushButton("🚀 保存并同步网页")
        self.btn_apply.setStyleSheet("background: #0078d4; color: white; font-weight: bold; padding: 8px;")
        self.btn_apply.clicked.connect(self.apply_settings)
        h_layout.addWidget(self.btn_apply)
        layout.addLayout(h_layout)

        # 日志区
        layout.addWidget(QLabel("📊 运行日志:"))
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setStyleSheet("background: #1e1e1e; color: #00ff00;")
        layout.addWidget(self.log_area)

        self.setLayout(layout)
        self.add_log("系统已就绪。点击关闭按钮将转入后台运行。")

    def init_tray(self):
        self.tray = QSystemTrayIcon(self)
        self.tray.setIcon(self.style().standardIcon(21)) # 邮件图标
        menu = QMenu()
        menu.addAction("打开管理界面", self.showNormal)
        menu.addAction("彻底退出程序", QApplication.instance().quit)
        self.tray.setContextMenu(menu)
        self.tray.activated.connect(lambda r: self.showNormal() if r == QSystemTrayIcon.Trigger else None)
        self.tray.show()

    def closeEvent(self, event):
        """实现点击 X 最小化到托盘"""
        self.hide()
        self.tray.showMessage("RD同步助手", "程序已转入后台运行", QSystemTrayIcon.Information, 1000)
        event.ignore()

    def add_log(self, text):
        self.log_area.append(f"[{time.strftime('%H:%M:%S')}] {text}")

    def apply_settings(self):
        self.share_dir = self.edit_path.text().strip()
        self.target_kw = self.edit_kw.text().strip()
        self.copyright_text = self.edit_copy.text().strip()
        try:
            self.scan_ms = int(self.edit_freq.text())
            self.timer.start(self.scan_ms)
            self.update_web_ui()
            self.add_log(f"✅ 配置已更新并重绘网页。当前监控: {self.target_kw}")
        except:
            self.add_log("❌ 频率格式错误")

    def poll_signal(self):
        sig_file = os.path.join(self.share_dir, "sync_request.txt")
        if os.path.exists(sig_file):
            self.add_log("📩 收到信号，唤醒 PowerShell 执行 MHT 抓取...")
            self.run_ps_logic()
            try: os.remove(sig_file)
            except: pass

    def run_ps_logic(self):
        """调用 PowerShell 解决乱码并以原名保存 MHT"""
        ps_cmd = f"""
        $KW = "{self.target_kw}"; $DIR = "{self.share_dir}"
        $ol = New-Object -ComObject Outlook.Application
        $ns = $ol.GetNamespace("MAPI")
        $fnd = $false
        foreach($f in $ns.Folders){{
            $ib = $f.Folders | ?{{$_.Name -match "收件箱|Inbox"}}
            if($ib){{
                $items = $ib.Items; $items.Sort("[ReceivedTime]", $true)
                for($i=1; $i -le [Math]::Min($items.Count, 30); $i++){{
                    $it = $items.Item($i)
                    if($it.Subject -like "*$KW*" -or $it.SenderName -like "*$KW*"){{
                        $fnd = $true
                        $n = ($it.Subject -replace '[\\\\/:*?"<>|]', '_')
                        if($n.Length -gt 50){{ $n = $n.Substring(0,50) }}
                        $p = Join-Path $DIR "$n.mht"
                        $it.SaveAs($p, 10); Start-Sleep -Seconds 2
                        break
                    }}
                }}
            }}
            if($fnd){{break}}
        }}
        """
        try:
            subprocess.run(["powershell", "-Command", ps_cmd], capture_output=True, text=True, creationflags=0x08000000)
            self.add_log("✅ 同步动作完成。")
            self.update_web_ui()
        except Exception as e:
            self.add_log(f"❌ 运行失败: {e}")

    def update_web_ui(self):
        """根据 UI 实时设置生成网页内容"""
        if not os.path.exists(self.share_dir): return
        # 自动生成 .bat 同步工具
        bat_tool = os.path.join(self.share_dir, "点击同步.bat")
        with open(bat_tool, "w", encoding="gbk") as f:
            f.write(f'@echo off\necho %date% %time% > "%~dp0sync_request.txt"\necho 指令已发送！\ntimeout /t 2 > nul\nexit')

        files = [f for f in os.listdir(self.share_dir) if f.endswith('.mht')]
        files.sort(key=lambda x: os.path.getmtime(os.path.join(self.share_dir, x)), reverse=True)
        
        index_path = os.path.join(self.share_dir, "index.html")
        html = f"""
        <!DOCTYPE html>
        <html lang="zh">
        <head>
            <meta charset="utf-8">
            <meta http-equiv="refresh" content="60">
            <title>RD Team 邮件看板</title>
            <style>
                :root {{ --main: #0078d4; --bg: #f3f2f1; }}
                body {{ font-family: 'Segoe UI', sans-serif; background: var(--bg); margin: 0; padding: 20px; }}
                .container {{ max-width: 900px; margin: 0 auto; background: white; border-radius: 12px; box-shadow: 0 8px 30px rgba(0,0,0,0.1); overflow: hidden; }}
                .header {{ background: var(--main); color: white; padding: 30px; display: flex; justify-content: space-between; align-items: center; }}
                .btn-sync {{ background: white; color: var(--main); padding: 12px 25px; border-radius: 25px; text-decoration: none; font-weight: bold; box-shadow: 0 4px 10px rgba(0,0,0,0.2); transition: 0.3s; }}
                .btn-sync:hover {{ transform: scale(1.05); background: #f0f0f0; }}
                .list {{ padding: 20px; min-height: 350px; }}
                .item {{ display: flex; align-items: center; padding: 18px; border-bottom: 1px solid #eee; text-decoration: none; color: #323130; transition: 0.2s; border-radius: 8px; }}
                .item:hover {{ background: #f9f9f9; transform: translateX(10px); }}
                .item-title {{ flex: 1; font-weight: 600; font-size: 16px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
                .footer {{ background: #fafafa; text-align: center; padding: 20px; color: #888; font-size: 13px; border-top: 1px solid #eee; }}
                .tag {{ background: #dff6dd; color: #107c10; padding: 3px 10px; border-radius: 12px; font-size: 11px; margin-left: 15px; font-weight: bold; }}
            </style>
        </head>
        <body>
            <div class="container">
                <div class="header">
                    <div>
                        <h2 style="margin:0; font-size: 24px;">📫 RD Team 邮件看板</h2>
                        <small>实时监控：{self.target_kw}</small>
                    </div>
                    <a href="点击同步.bat" class="btn-sync">🔄 一键抓取最新</a>
                </div>
                <div class="list">
        """
        for i, f in enumerate(files):
            mtime = time.strftime("%Y-%m-%d %H:%M", time.localtime(os.path.getmtime(os.path.join(self.share_dir, f))))
            tag = '<span class="tag">NEW</span>' if i == 0 else ""
            html += f"""
                <a href="{f}" target="_blank" class="item">
                    <span style="font-size:24px; margin-right:20px;">📄</span>
                    <span class="item-title">{f.replace('.mht','')} {tag}</span>
                    <span style="color:#999;">{mtime}</span>
                </a>
            """
        html += f"""
                </div>
                <div class="footer">{self.copyright_text}</div>
            </div>
            <div style="text-align:center; color:#bbb; margin-top:25px; font-size:12px;">
                💡 建议使用 Chrome 或 Edge 浏览器访问。本页面每 60 秒自动更新。
            </div>
        </body>
        </html>
        """
        with open(index_path, "w", encoding="utf-8") as f_out:
            f_out.write(html)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    QApplication.setQuitOnLastWindowClosed(False)
    m = OutlookMasterPro()
    m.show()
    sys.exit(app.exec_())
