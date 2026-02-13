import sys, os, re, time, subprocess, tempfile, urllib.parse
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QTextEdit, QSystemTrayIcon, QMenu)
from PyQt5.QtCore import QTimer, Qt

class OutlookMHTMaster(QWidget):
    def __init__(self):
        super().__init__()
        # --- 建议：先在 Windows 映射 Z 盘指向 \\10.1.93.32\DT_HU_RDteam_F ---
        # 然后将下方路径改为 r'Z:\视频\Z\ZOUQIU\paican'
        self.share_dir = r'\\10.1.93.32\DT_HU_RDteam_F\视频\Z\ZOUQIU\paican'
        self.target_kw = 'EDFA' 
        self.copyright_text = "© 2024-2026 RD Team | 视频组技术支持"
        self.interval_min = 10     
        self.start_hour = 8        
        self.end_hour = 18         
        
        self.tmp_log = os.path.join(tempfile.gettempdir(), "outlook_sync_res.txt")
        self.init_ui()
        self.init_tray()
        
        self.sync_timer = QTimer(self)
        self.sync_timer.timeout.connect(self.run_cycle)
        QTimer.singleShot(2000, self.run_cycle)

    def init_ui(self):
        self.setWindowTitle("RD 邮件全自动看板 (V4.0 预览修复版)")
        self.resize(580, 550)
        self.setWindowFlags(Qt.WindowStaysOnTopHint)
        layout = QVBoxLayout()
        
        layout.addWidget(QLabel("📂 保存目录 (推荐使用映射后的盘符如 Z:\...):"))
        self.edit_path = QLineEdit(self.share_dir); layout.addWidget(self.edit_path)

        layout.addWidget(QLabel("📧 监控关键词:"))
        self.edit_kw = QLineEdit(self.target_kw); layout.addWidget(self.edit_kw)

        h_time_layout = QHBoxLayout()
        h_time_layout.addWidget(QLabel("⏰ 活跃时段:")); self.edit_start = QLineEdit(str(self.start_hour)); h_time_layout.addWidget(self.edit_start)
        h_time_layout.addWidget(QLabel("至")); self.edit_end = QLineEdit(str(self.end_hour)); h_time_layout.addWidget(self.edit_end)
        layout.addLayout(h_time_layout)

        h_freq_layout = QHBoxLayout()
        h_freq_layout.addWidget(QLabel("⏱ 频率(分):")); self.edit_freq = QLineEdit(str(self.interval_min)); h_freq_layout.addWidget(self.edit_freq)
        layout.addLayout(h_freq_layout)

        layout.addWidget(QLabel("📝 底部版权:"))
        self.edit_copy = QLineEdit(self.copyright_text); layout.addWidget(self.edit_copy)

        self.btn_apply = QPushButton("🚀 保存配置并启动")
        self.btn_apply.setStyleSheet("background: #0078d4; color: white; font-weight: bold; padding: 12px;")
        self.btn_apply.clicked.connect(self.apply_settings)
        layout.addWidget(self.btn_apply)

        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setStyleSheet("background: #1e1e1e; color: #00ff00; font-family: 'Consolas';")
        layout.addWidget(self.log_area)
        self.setLayout(layout)

    def init_tray(self):
        self.tray = QSystemTrayIcon(self)
        self.tray.setIcon(self.style().standardIcon(21)) 
        menu = QMenu()
        menu.addAction("主界面", self.showNormal)
        menu.addAction("退出", QApplication.instance().quit)
        self.tray.setContextMenu(menu)
        self.tray.show()

    def add_log(self, text):
        safe_text = str(text).replace('\x00', '')
        self.log_area.append(f"[{time.strftime('%H:%M:%S')}] {safe_text}")

    def apply_settings(self):
        self.share_dir, self.target_kw = self.edit_path.text().strip(), self.edit_kw.text().strip()
        try:
            self.interval_min, self.start_hour, self.end_hour = int(self.edit_freq.text()), int(self.edit_start.text()), int(self.edit_end.text())
            self.add_log("✅ 配置已重载")
            self.run_cycle()
        except: self.add_log("❌ 输入有误")

    def run_cycle(self):
        now_hour = int(time.strftime("%H"))
        if not (self.start_hour <= now_hour < self.end_hour):
            self.add_log(f"💤 静默中 ({now_hour}点)")
            self.sync_timer.start(30 * 60000); return
        self.run_shell_mht_logic()
        self.sync_timer.start(self.interval_min * 60000)

    def run_shell_mht_logic(self):
        """Shell 逻辑：文件中转方案"""
        if os.path.exists(self.tmp_log): 
            try: os.remove(self.tmp_log)
            except: pass

        ps_dir = self.share_dir.replace('\\', '\\\\').replace('"', '""')
        ps_kw = self.target_kw.replace('"', '""')
        ps_tmp = self.tmp_log.replace('\\', '\\\\')
        
        ps_cmd = f"""
        $ErrorActionPreference = "Stop"
        try {{
            if (!(Test-Path "{ps_dir}")) {{ New-Item -ItemType Directory -Path "{ps_dir}" -Force | Out-Null }}
            $ol = New-Object -ComObject Outlook.Application
            $ns = $ol.GetNamespace("MAPI")
            $limit = (Get-Date).AddDays(-3)
            $it = $ns.GetDefaultFolder(6).Items | Where-Object {{ 
                $_.ReceivedTime -gt $limit -and ($_.Subject -like "*{ps_kw}*" -or $_.SenderName -like "*{ps_kw}*") 
            }} | Sort-Object ReceivedTime -Descending | Select-Object -First 1
            if ($it) {{
                $name = $it.Subject -replace '[\\x00-\\x1f\\\\/:*?"<>|]', '_'
                $path = Join-Path "{ps_dir}" "$($name.Trim()).mht"
                if (!(Test-Path $path)) {{
                    $it.SaveAs($path, 10)
                    "SUCCESS|$name" | Out-File "{ps_tmp}" -Encoding utf8
                }} else {{ "EXISTS" | Out-File "{ps_tmp}" -Encoding utf8 }}
            }} else {{ "NOTFOUND" | Out-File "{ps_tmp}" -Encoding utf8 }}
        }} catch {{ "ERROR|$($_.Exception.Message -replace '[\\x00-\\x1f]', '')" | Out-File "{ps_tmp}" -Encoding utf8
        }} finally {{ if ($ol) {{ [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ol) | Out-Null }} }}
        """
        try:
            subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_cmd], creationflags=0x08000000, timeout=60)
            if os.path.exists(self.tmp_log):
                with open(self.tmp_log, 'rb') as f:
                    content = f.read().decode('utf-8', errors='ignore').replace('\x00', '').strip()
                if "SUCCESS|" in content: self.add_log(f"✅ 同步: {content.split('|')[-1]}")
                elif "EXISTS" in content: self.add_log("ℹ️ 已存在")
                elif "NOTFOUND" in content: self.add_log("❓ 未匹配")
                elif "ERROR|" in content: self.add_log(f"❌ 报错: {content.split('|')[-1]}")
            self.generate_html_index()
        except Exception as e: self.add_log(f"❌ 异常: {str(e)}")

    def generate_html_index(self):
        """生成支持本地安全预览的看板"""
        if not os.path.exists(self.share_dir): return
        files = [f for f in os.listdir(self.share_dir) if f.endswith('.mht')]
        files.sort(key=lambda x: os.path.getmtime(os.path.join(self.share_dir, x)), reverse=True)
        
        html = f"""
        <!DOCTYPE html><html><head><meta charset="utf-8">
        <style>
            body {{ margin: 0; display: flex; height: 100vh; font-family: 'Segoe UI'; overflow: hidden; }}
            .sidebar {{ width: 320px; background: #f5f5f5; border-right: 1px solid #ddd; display: flex; flex-direction: column; }}
            .header {{ background: #0078d4; color: white; padding: 15px; font-weight: bold; }}
            .list {{ flex: 1; overflow-y: auto; }}
            .item {{ display: block; padding: 12px; border-bottom: 1px solid #eee; cursor: pointer; }}
            .item:hover {{ background: #eef; }}
            .item.active {{ background: #d0e7ff; border-left: 5px solid #0078d4; }}
            .time {{ font-size: 11px; color: #888; display: block; }}
            .preview {{ flex: 1; background: #fff; position: relative; }}
            iframe {{ width: 100%; height: 100%; border: none; }}
            .footer {{ font-size: 11px; padding: 10px; text-align: center; color: #999; border-top: 1px solid #ddd; }}
        </style>
        <script>
            function view(el, url) {{
                document.querySelectorAll('.item').forEach(i => i.classList.remove('active'));
                el.classList.add('active');
                document.getElementById('f').src = url;
            }}
        </script></head>
        <body>
            <div class="sidebar">
                <div class="header">📫 RD 邮件看板<br><small style="font-weight:normal">关键词: {self.target_kw}</small></div>
                <div class="list">
        """
        for i, f in enumerate(files):
            mt = time.strftime("%m-%d %H:%M", time.localtime(os.path.getmtime(os.path.join(self.share_dir, f))))
            # 核心锁：对路径进行安全编码
            safe_f = urllib.parse.quote(f)
            html += f'<div class="item" onclick="view(this, \'{safe_f}\')"><b>{f[:35]}...</b><span class="time">🕒 {mt}</span></div>'
        
        html += f"""
                </div><div class="footer">{self.copyright_text}</div>
            </div>
            <div class="preview"><iframe id="f"></iframe></div>
        </body></html>
        """
        with open(os.path.join(self.share_dir, "index.html"), "w", encoding="utf-8") as f: f.write(html)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = OutlookMHTMaster(); ex.show()
    sys.exit(app.exec_())
