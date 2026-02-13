import sys, os, re, time, subprocess, tempfile
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QTextEdit, QSystemTrayIcon, QMenu)
from PyQt5.QtCore import QTimer, Qt

class OutlookMHTMaster(QWidget):
    def __init__(self):
        super().__init__()
        # --- 默认配置 ---
        self.share_dir = r'\\10.1.93.32\DT_HU_RDteam_F\视频\Z\ZOUQIU\paican'
        self.target_kw = 'EDFA' 
        self.copyright_text = "© 2024-2026 RD Team | 视频组技术支持"
        self.interval_min = 10     
        self.start_hour = 8        
        self.end_hour = 18         
        
        # 临时文件路径：用于接收 Shell 运行结果
        self.tmp_log = os.path.join(tempfile.gettempdir(), "outlook_sync_res.txt")
        
        self.init_ui()
        self.init_tray()
        
        self.sync_timer = QTimer(self)
        self.sync_timer.timeout.connect(self.run_cycle)
        QTimer.singleShot(2000, self.run_cycle)

    def init_ui(self):
        self.setWindowTitle("RD 邮件全自动看板 (V3.0 文件中转版)")
        self.resize(580, 550)
        self.setWindowFlags(Qt.WindowStaysOnTopHint)
        layout = QVBoxLayout()
        
        layout.addWidget(QLabel("📂 共享保存目录 (MHT 模式):"))
        self.edit_path = QLineEdit(self.share_dir); layout.addWidget(self.edit_path)

        layout.addWidget(QLabel("📧 监控关键词 (发件人/标题):"))
        self.edit_kw = QLineEdit(self.target_kw); layout.addWidget(self.edit_kw)

        h_time_layout = QHBoxLayout()
        h_time_layout.addWidget(QLabel("⏰ 活跃时段:")); self.edit_start = QLineEdit(str(self.start_hour)); h_time_layout.addWidget(self.edit_start)
        h_time_layout.addWidget(QLabel("至")); self.edit_end = QLineEdit(str(self.end_hour)); h_time_layout.addWidget(self.edit_end)
        layout.addLayout(h_time_layout)

        h_freq_layout = QHBoxLayout()
        h_freq_layout.addWidget(QLabel("⏱ 频率(分):")); self.edit_freq = QLineEdit(str(self.interval_min)); h_freq_layout.addWidget(self.edit_freq)
        layout.addLayout(h_freq_layout)

        layout.addWidget(QLabel("📝 网页底部版权修改:"))
        self.edit_copy = QLineEdit(self.copyright_text); layout.addWidget(self.edit_copy)

        self.btn_apply = QPushButton("🚀 保存并立即强制同步")
        self.btn_apply.setStyleSheet("background: #0078d4; color: white; font-weight: bold; padding: 12px; border-radius: 4px;")
        self.btn_apply.clicked.connect(self.apply_settings)
        layout.addWidget(self.btn_apply)

        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setStyleSheet("background: #1e1e1e; color: #00ff00; font-family: 'Consolas'; font-size: 12px;")
        layout.addWidget(self.log_area)
        self.setLayout(layout)

    def init_tray(self):
        self.tray = QSystemTrayIcon(self)
        self.tray.setIcon(self.style().standardIcon(21)) 
        menu = QMenu()
        menu.addAction("显示主界面", self.showNormal)
        menu.addAction("完全退出", QApplication.instance().quit)
        self.tray.setContextMenu(menu)
        self.tray.show()

    def closeEvent(self, event):
        self.hide(); event.ignore()

    def add_log(self, text):
        # Python 层过滤，确保界面不因空字符崩溃
        safe_text = str(text).replace('\x00', '[NULL]')
        self.log_area.append(f"[{time.strftime('%H:%M:%S')}] {safe_text}")

    def apply_settings(self):
        self.share_dir, self.target_kw = self.edit_path.text().strip(), self.edit_kw.text().strip()
        self.copyright_text = self.edit_copy.text().strip()
        try:
            self.interval_min, self.start_hour, self.end_hour = int(self.edit_freq.text()), int(self.edit_start.text()), int(self.edit_end.text())
            self.add_log("✅ 配置已重载。")
            self.run_cycle()
        except: self.add_log("❌ 输入的数值格式有误。")

    def run_cycle(self):
        now_hour = int(time.strftime("%H"))
        if not (self.start_hour <= now_hour < self.end_hour):
            self.add_log(f"💤 当前 {now_hour} 点，处于静默时段。")
            self.sync_timer.start(30 * 60000)
            return
        self.run_shell_mht_logic()
        self.sync_timer.start(self.interval_min * 60000)

    def run_shell_mht_logic(self):
        """核心修复：通过临时文件接收 Shell 结果，彻底避开 Null Character"""
        if os.path.exists(self.tmp_log): 
            try: os.remove(self.tmp_log)
            except: pass

        ps_dir = self.share_dir.replace('\\', '\\\\').replace('"', '""')
        ps_kw = self.target_kw.replace('"', '""')
        ps_tmp = self.tmp_log.replace('\\', '\\\\')
        
        # 封装 Shell 逻辑：直接将运行状态 Out-File 到临时文件
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
                $name = $name.Trim()
                if ($name.Length -gt 55) {{ $name = $name.Substring(0,55) }}
                $path = Join-Path "{ps_dir}" "$name.mht"
                
                if (!(Test-Path $path)) {{
                    $it.SaveAs($path, 10)
                    "SUCCESS|$name" | Out-File "{ps_tmp}" -Encoding utf8
                }} else {{ "EXISTS" | Out-File "{ps_tmp}" -Encoding utf8 }}
            }} else {{ "NOTFOUND" | Out-File "{ps_tmp}" -Encoding utf8 }}
        }} catch {{
            "ERROR|$($_.Exception.Message -replace '[\\x00-\\x1f]', '')" | Out-File "{ps_tmp}" -Encoding utf8
        }} finally {{
            if ($ol) {{ [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ol) | Out-Null }}
        }}
        """
        try:
            # 执行时不捕获 stdout，直接让它静默运行
            subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_cmd], 
                           creationflags=0x08000000, timeout=60)
            
            # 读取结果文件（二进制读取 + ignore 模式，万无一失）
            if os.path.exists(self.tmp_log):
                with open(self.tmp_log, 'rb') as f:
                    content = f.read().decode('utf-8', errors='ignore').replace('\x00', '').strip()
                
                if "SUCCESS|" in content:
                    self.add_log(f"✅ 同步成功: {content.split('|')[-1]}")
                elif "EXISTS" in content:
                    self.add_log("ℹ️ 邮件已同步过，无需更新。")
                elif "NOTFOUND" in content:
                    self.add_log(f"❓ 未找到关键词 '{self.target_kw}' 的新邮件。")
                elif "ERROR|" in content:
                    self.add_log(f"❌ Shell 报错: {content.split('|')[-1]}")
            else:
                self.add_log("⚠️ 警告：Shell 进程未生成结果文件。")
            
            self.generate_html_index()
        except Exception as e:
            self.add_log(f"❌ 系统异常: {str(e)}")

    def generate_html_index(self):
        """生成支持左右预览的 HTML 页面"""
        if not os.path.exists(self.share_dir): return
        files = [f for f in os.listdir(self.share_dir) if f.endswith('.mht')]
        files.sort(key=lambda x: os.path.getmtime(os.path.join(self.share_dir, x)), reverse=True)
        
        html = f"""
        <!DOCTYPE html><html><head><meta charset="utf-8">
        <style>
            body {{ margin: 0; display: flex; height: 100vh; font-family: 'Segoe UI', Tahoma; overflow: hidden; }}
            .sidebar {{ width: 340px; background: #fafafa; border-right: 1px solid #ddd; display: flex; flex-direction: column; }}
            .header {{ background: #0078d4; color: white; padding: 15px; }}
            .list {{ flex: 1; overflow-y: auto; }}
            .item {{ display: block; padding: 12px 15px; border-bottom: 1px solid #eee; cursor: pointer; text-decoration: none; color: #333; }}
            .item:hover {{ background: #f0f7ff; }}
            .item.active {{ background: #e1f0fe; border-left: 4px solid #0078d4; }}
            .time {{ font-size: 11px; color: #999; display: block; margin-top: 4px; }}
            .preview {{ flex: 1; background: #fff; position: relative; }}
            iframe {{ width: 100%; height: 100%; border: none; }}
            .placeholder {{ position: absolute; top: 45%; width: 100%; text-align: center; color: #ccc; }}
            .footer {{ padding: 10px; background: #eee; font-size: 11px; text-align: center; color: #666; }}
        </style>
        <script>
            function view(el, url) {{
                document.querySelectorAll('.item').forEach(i => i.classList.remove('active'));
                el.classList.add('active');
                document.getElementById('f').src = url;
                document.getElementById('p').style.display = 'none';
            }}
        </script></head>
        <body>
            <div class="sidebar">
                <div class="header"><b>📫 RD Team 邮件看板</b><br><small>关键词: {self.target_kw}</small></div>
                <div class="list">
        """
        for i, f in enumerate(files):
            mt = time.strftime("%Y-%m-%d %H:%M", time.localtime(os.path.getmtime(os.path.join(self.share_dir, f))))
            html += f'<div class="item" onclick="view(this, \'{f}\')"><b>{f[:42]}...</b><span class="time">📅 {mt}</span></div>'
        
        html += f"""
                </div><div class="footer">{self.copyright_text}</div>
            </div>
            <div class="preview">
                <div id="p" class="placeholder"><h2>📂 请选择邮件查看内容</h2></div>
                <iframe id="f"></iframe>
            </div>
        </body></html>
        """
        try:
            with open(os.path.join(self.share_dir, "index.html"), "w", encoding="utf-8") as f: f.write(html)
        except: pass

if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = OutlookMHTMaster()
    ex.show()
    sys.exit(app.exec_())
