import sys, os, re, time, subprocess, tempfile, urllib.parse
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QTextEdit, QSystemTrayIcon, QMenu)
from PyQt5.QtCore import QTimer, Qt

class OutlookMHTMaster(QWidget):
    def __init__(self):
        super().__init__()
        # --- 基础配置 ---
        self.share_dir = r'\\10.1.93.32\DT_HU_RDteam_F\视频\Z\ZOUQIU\paican'
        self.target_kw = 'EDFA' 
        self.interval_min = 10     
        self.keep_days = 7
        self.copyright_text = "© 2024-2026 RD Team | 视频组技术支持"
        
        self.tmp_log = os.path.join(tempfile.gettempdir(), "outlook_sync_res.txt")
        self.last_sync_time = "尚未同步"
        
        self.init_ui()
        self.init_tray()
        self.sync_timer = QTimer(self)
        self.sync_timer.timeout.connect(self.run_cycle)
        QTimer.singleShot(2000, self.run_cycle)

    def init_ui(self):
        self.setWindowTitle("RD 邮件同步助手 V10.0 (EP任务令版)")
        self.resize(600, 620)
        layout = QVBoxLayout()
        
        form = QVBoxLayout()
        self.edit_path = QLineEdit(self.share_dir); form.addWidget(QLabel("📂 共享网络路径 (UNC):")); form.addWidget(self.edit_path)
        
        h_box1 = QHBoxLayout()
        self.edit_kw = QLineEdit(self.target_kw); h_box1.addWidget(QLabel("📧 关键词:")); h_box1.addWidget(self.edit_kw)
        self.edit_freq = QLineEdit(str(self.interval_min)); h_box1.addWidget(QLabel("⏱ 频率(分):")); h_box1.addWidget(self.edit_freq)
        form.addLayout(h_box1)
        
        h_box2 = QHBoxLayout()
        self.edit_days = QLineEdit(str(self.keep_days)); h_box2.addWidget(QLabel("♻️ 清理天数:")); h_box2.addWidget(self.edit_days)
        self.edit_copy = QLineEdit(self.copyright_text); h_box2.addWidget(QLabel("📝 版权信息:")); h_box2.addWidget(self.edit_copy)
        form.addLayout(h_box2)
        
        layout.addLayout(form)
        self.btn_apply = QPushButton("🚀 保存配置并强制同步")
        self.btn_apply.setStyleSheet("background: #107c10; color: white; font-weight: bold; padding: 12px; border-radius: 4px;")
        self.btn_apply.clicked.connect(self.apply_settings); layout.addWidget(self.btn_apply)

        self.log_area = QTextEdit(); self.log_area.setReadOnly(True)
        self.log_area.setStyleSheet("background: #1e1e1e; color: #00ff00; font-family: 'Consolas';")
        layout.addWidget(self.log_area); self.setLayout(layout)

    def init_tray(self):
        self.tray = QSystemTrayIcon(self); self.tray.setIcon(self.style().standardIcon(21))
        menu = QMenu(); menu.addAction("打开", self.showNormal); menu.addAction("退出", QApplication.instance().quit)
        self.tray.setContextMenu(menu); self.tray.show()

    def add_log(self, text):
        t_prefix = time.strftime('%H:%M:%S')
        # 兼容性修复：变量抽离，避免 f-string 反斜杠报错
        clean_msg = str(text).replace('\x00', '')
        self.log_area.append(f"[{t_prefix}] {clean_msg}")

    def apply_settings(self):
        self.share_dir = self.edit_path.text().strip()
        self.target_kw = self.edit_kw.text().strip()
        try:
            self.interval_min = int(self.edit_freq.text())
            self.keep_days = int(self.edit_days.text())
            self.add_log("⚙️ 配置已更新")
            self.run_cycle()
        except: self.add_log("❌ 数值格式错误")

    def run_cycle(self):
        self.run_shell_mht_logic()
        self.last_sync_time = time.strftime('%Y-%m-%d %H:%M:%S')
        self.sync_timer.start(self.interval_min * 60000)

    def run_shell_mht_logic(self):
        """核心 Shell：清理 + 搜索 + HTML 导出"""
        ps_dir = self.share_dir.replace('\\', '\\\\').replace('"', '""')
        ps_kw = self.target_kw.replace('"', '""')
        ps_tmp = self.tmp_log.replace('\\', '\\\\')
        
        ps_cmd = f"""
        $ErrorActionPreference = "Stop"
        try {{
            if (!(Test-Path "{ps_dir}")) {{ New-Item -ItemType Directory -Path "{ps_dir}" -Force | Out-Null }}
            
            # 清理过期文件
            $expire = (Get-Date).AddDays(-{self.keep_days})
            Get-ChildItem "{ps_dir}" -Include *.html,*.mht,*.files -Recurse | Where-Object {{ $_.LastWriteTime -lt $expire }} | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
            
            # Outlook 抓取
            $ol = New-Object -ComObject Outlook.Application
            $ns = $ol.GetNamespace("MAPI")
            $it = $ns.GetDefaultFolder(6).Items | Where-Object {{ 
                $_.ReceivedTime -gt (Get-Date).AddDays(-5) -and ($_.Subject -like "*{ps_kw}*" -or $_.SenderName -like "*{ps_kw}*") 
            }} | Sort-Object ReceivedTime -Descending | Select-Object -First 1
            
            if ($it) {{
                $name = $it.Subject -replace '[\\x00-\\x1f\\\\/:*?"<>|]', '_'
                $path = Join-Path "{ps_dir}" "$($name.Trim()).html"
                if (!(Test-Path $path)) {{
                    $it.SaveAs($path, 4)
                    "SUCCESS|$name" | Out-File "{ps_tmp}" -Encoding utf8
                }} else {{ "EXISTS" | Out-File "{ps_tmp}" -Encoding utf8 }}
            }} else {{ "NOTFOUND" | Out-File "{ps_tmp}" -Encoding utf8 }}
        }} catch {{ 
            "ERROR|$($_.Exception.Message)" | Out-File "{ps_tmp}" -Encoding utf8
        }} finally {{ if ($ol) {{ [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ol) | Out-Null }} }}
        """
        try:
            subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_cmd], creationflags=0x08000000, timeout=60)
            if os.path.exists(self.tmp_log):
                with open(self.tmp_log, 'rb') as f:
                    content = f.read().decode('utf-8', errors='ignore').replace('\x00', '').strip()
                if "SUCCESS|" in content: self.add_log("✅ 同步成功: " + content.split('|')[-1])
                elif "ERROR|" in content: self.add_log("❌ 错误: " + content.split('|')[-1])
            self.generate_html_index()
        except Exception as e: self.add_log("❌ 异常: " + str(e))

    def generate_html_index(self):
        """生成带 EP 任务令索引的静态 HTML"""
        if not os.path.exists(self.share_dir): return
        files = [f for f in os.listdir(self.share_dir) if f.endswith('.html') and f != 'index.html']
        files.sort(key=lambda x: os.path.getmtime(os.path.join(self.share_dir, x)), reverse=True)
        
        html_items = ""
        for f in files:
            f_path = os.path.join(self.share_dir, f)
            mt = time.strftime("%m-%d %H:%M", time.localtime(os.path.getmtime(f_path)))
            ep_codes, preview = [], ""
            try:
                with open(f_path, 'r', encoding='utf-8', errors='ignore') as fo:
                    raw = fo.read()
                    # 识别 EP 开头的 11 位任务令
                    found = re.findall(r'\bEP[A-Za-z0-9]{9}\b', raw)
                    ep_codes = list(set([str(i).upper() for i in found]))
                    preview = re.sub('<[^<]+?>', '', raw).replace('\n', '').replace('"', "'")[:150]
            except: pass
            
            tags = "".join([f'<span class="tag">{c}</span>' for c in ep_codes[:3]])
            safe_f = urllib.parse.quote(f)
            search_str = (f + " ".join(ep_codes) + preview).lower()
            
            html_items += f'''
                <div class="item" onclick="openMail(this, '{safe_f}')" data-search="{search_str}">
                    <b>{f[:45]}</b>
                    <div style="margin:5px 0;">{tags}</div>
                    <span class="time">🕒 {mt}</span>
                </div>
            '''

        html = f"""
        <!DOCTYPE html><html><head><meta charset="utf-8">
        <title>RD看板 V10</title>
        <style>
            :root {{ --main: #0078d4; }}
            body {{ margin: 0; display: flex; height: 100vh; font-family: 'Segoe UI'; background: #f3f2f1; overflow: hidden; }}
            .sidebar {{ width: 340px; background: #fff; border-right: 1px solid #ddd; display: flex; flex-direction: column; }}
            .header {{ background: var(--main); color: white; padding: 18px; }}
            .search {{ padding: 10px; border-bottom: 1px solid #eee; }}
            .search input {{ width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box; }}
            .list {{ flex: 1; overflow-y: auto; }}
            .item {{ padding: 12px 15px; border-bottom: 1px solid #f3f2f1; cursor: pointer; }}
            .item:hover {{ background: #f9f9f9; }}
            .item.active {{ background: #eff6fc; border-left: 5px solid var(--main); }}
            .tag {{ background: var(--main); color: white; padding: 2px 6px; border-radius: 3px; font-size: 10px; margin-right: 5px; font-family: Consolas; }}
            .time {{ font-size: 11px; color: #999; }}
            .preview {{ flex: 1; background: #fff; position: relative; }}
            iframe {{ width: 100%; height: 100%; border: none; }}
            .footer {{ font-size: 10px; padding: 10px; text-align: center; color: #999; border-top: 1px solid #eee; }}
        </style>
        <script>
            function doSearch(e) {{
                let v = document.getElementById('q').value.toLowerCase();
                let items = document.querySelectorAll('.item');
                let match = null;
                items.forEach(i => {{
                    if(i.getAttribute('data-search').includes(v)) {{ i.style.display='block'; if(!match) match=i; }}
                    else {{ i.style.display='none'; }}
                }});
                if(e.keyCode === 13 && match) match.click();
            }}
            function openMail(el, url) {{
                document.querySelectorAll('.item').forEach(i => i.classList.remove('active'));
                el.classList.add('active');
                document.getElementById('f').src = url;
            }}
        </script></head>
        <body>
            <div class="sidebar">
                <div class="header">📫 RD 邮件看板 | EP 识别<br><small style="font-weight:normal; font-size:10px;">最后更新: {self.last_sync_time}</small></div>
                <div class="search"><input id="q" onkeyup="doSearch(event)" placeholder="🔍 搜任务令并回车..."></div>
                <div class="list">{html_items}</div>
                <div class="footer">{self.copyright_text}</div>
            </div>
            <div class="preview"><iframe id="f"></iframe></div>
        </body></html>
        """
        try:
            with open(os.path.join(self.share_dir, "index.html"), "w", encoding="utf-8") as f: f.write(html)
        except: pass

if __name__ == "__main__":
    app = QApplication(sys.argv); ex = OutlookMHTMaster(); ex.show(); sys.exit(app.exec_())
