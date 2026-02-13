import sys, os, re, time, subprocess, tempfile, urllib.parse, base64
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QTextEdit, QSystemTrayIcon, QMenu)
from PyQt5.QtCore import QTimer, Qt

class OutlookMHTMaster(QWidget):
    def __init__(self):
        super().__init__()
        # --- 配置区 ---
        self.share_dir = r'\\10.1.93.32\DT_HU_RDteam_F\视频\Z\ZOUQIU\paican'
        self.target_kw = 'EDFA' 
        self.interval_min = 10     
        self.copyright_text = "© 2024-2026 RD Team | 视频组技术支持"
        
        self.tmp_log = os.path.join(tempfile.gettempdir(), "outlook_sync_v13.txt")
        self.last_sync_time = "尚未开始"
        
        self.init_ui()
        self.init_tray()
        self.sync_timer = QTimer(self)
        self.sync_timer.timeout.connect(self.run_cycle)
        QTimer.singleShot(2000, self.run_cycle)

    def init_ui(self):
        self.setWindowTitle("RD 邮件同步 V13.0 (全量索引预览版)")
        self.resize(620, 600)
        layout = QVBoxLayout()
        
        form = QVBoxLayout()
        self.edit_path = QLineEdit(self.share_dir); form.addWidget(QLabel("📂 共享网络路径 (UNC):")); form.addWidget(self.edit_path)
        
        h_box = QHBoxLayout()
        self.edit_kw = QLineEdit(self.target_kw); h_box.addWidget(QLabel("📧 关键词:")); h_box.addWidget(self.edit_kw)
        self.edit_freq = QLineEdit(str(self.interval_min)); h_box.addWidget(QLabel("⏱ 频率(分):")); h_box.addWidget(self.edit_freq)
        form.addLayout(h_box)
        layout.addLayout(form)

        self.btn_apply = QPushButton("🚀 启动全量扫描同步")
        self.btn_apply.setStyleSheet("background: #107c10; color: white; font-weight: bold; padding: 12px; border-radius: 4px;")
        self.btn_apply.clicked.connect(self.apply_settings); layout.addWidget(self.btn_apply)

        self.log_area = QTextEdit(); self.log_area.setReadOnly(True)
        self.log_area.setStyleSheet("background: #1e1e1e; color: #00ff00; font-family: 'Consolas';")
        layout.addWidget(self.log_area); self.setLayout(layout)

    def init_tray(self):
        self.tray = QSystemTrayIcon(self); self.tray.setIcon(self.style().standardIcon(21))
        menu = QMenu(); menu.addAction("显示界面", self.showNormal); menu.addAction("退出", QApplication.instance().quit)
        self.tray.setContextMenu(menu); self.tray.show()

    def add_log(self, text):
        t_str = time.strftime('%H:%M:%S')
        # 彻底规避 f-string 反斜杠 SyntaxError
        msg = str(text).replace('\x00', '')
        self.log_area.append(f"[{t_str}] {msg}")

    def apply_settings(self):
        self.share_dir, self.target_kw = self.edit_path.text().strip(), self.edit_kw.text().strip()
        self.run_cycle()

    def run_cycle(self):
        self.run_shell_logic()
        self.last_sync_time = time.strftime('%Y-%m-%d %H:%M:%S')
        try: freq = int(self.edit_freq.text())
        except: freq = 10
        self.sync_timer.start(freq * 60000)

    def run_shell_logic(self):
        """Shell逻辑：强制网络激活+Base64指令传输"""
        ps_dir = self.share_dir.replace('"', '""')
        ps_kw = self.target_kw.replace('"', '""')
        ps_tmp = self.tmp_log.replace('"', '""')
        
        # PowerShell 脚本：保存为 HTML 格式
        ps_script = f"""
        $ErrorActionPreference = "Stop"
        try {{
            if (!(Test-Path "{ps_dir}")) {{ New-Item -ItemType Directory -Path "{ps_dir}" -Force | Out-Null }}
            $ol = New-Object -ComObject Outlook.Application
            $ns = $ol.GetNamespace("MAPI")
            $it = $ns.GetDefaultFolder(6).Items | Where-Object {{ 
                $_.ReceivedTime -gt (Get-Date).AddDays(-3) -and ($_.Subject -like "*{ps_kw}*" -or $_.SenderName -like "*{ps_kw}*") 
            }} | Sort-Object ReceivedTime -Descending | Select-Object -First 1
            if ($it) {{
                $name = $it.Subject -replace '[\\x00-\\x1f\\\\/:*?"<>|]', '_'
                $path = Join-Path "{ps_dir}" "$($name.Trim()).html"
                if (!(Test-Path $path)) {{
                    $it.SaveAs($path, 4)
                    "SUCCESS|$name" | Out-File "{ps_tmp}" -Encoding utf8
                }} else {{ "EXISTS" | Out-File "{ps_tmp}" -Encoding utf8 }}
            }} else {{ "NOTFOUND" | Out-File "{ps_tmp}" -Encoding utf8 }}
        }} catch {{ "ERROR|$($_.Exception.Message)" | Out-File "{ps_tmp}" -Encoding utf8
        }} finally {{ if ($ol) {{ [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ol) | Out-Null }} }}
        """
        try:
            # 解决编码和乱码的终极武器：Base64
            ps_b64 = base64.b64encode(ps_script.encode('utf-16-le')).decode('ascii')
            subprocess.run(["powershell", "-NoProfile", "-EncodedCommand", ps_b64], creationflags=0x08000000, timeout=60)
            
            if os.path.exists(self.tmp_log):
                with open(self.tmp_log, 'rb') as f:
                    res = f.read().decode('utf-8', errors='ignore').strip()
                if "SUCCESS" in res: self.add_log("✅ 抓取新邮件: " + res.split('|')[-1])
            self.generate_html_index()
        except Exception as e: self.add_log("❌ 异常: " + str(e))

    def generate_html_index(self):
        """扫描所有文件，从表格中抓取 EP 任务令"""
        if not os.path.exists(self.share_dir): return
        files = [f for f in os.listdir(self.share_dir) if f.endswith('.html') and f != 'index.html']
        files.sort(key=lambda x: os.path.getmtime(os.path.join(self.share_dir, x)), reverse=True)
        
        items_html = ""
        self.add_log(f"🔎 正在扫描 {len(files)} 个历史文件建立索引...")
        
        for f in files:
            path = os.path.join(self.share_dir, f)
            mt = time.strftime("%m-%d %H:%M", time.localtime(os.path.getmtime(path)))
            ep_list = []
            try:
                # 二进制读，防止表格特殊字节导致乱码崩溃
                with open(path, 'rb') as fo:
                    # 只读前 50KB，足够覆盖大表格提取 EP 任务令
                    raw = fo.read(51200).decode('utf-8', errors='ignore')
                    # 剥离标签，在纯文本中匹配 EP 开头 11 位
                    text_only = re.sub('<[^<]+?>', '', raw)
                    ep_list = list(set(re.findall(r'\bEP[A-Za-z0-9]{9}\b', text_only)))
            except: pass
            
            tags = "".join([f'<span class="ep-tag">{c}</span>' for c in ep_list[:3]])
            safe_f = urllib.parse.quote(f)
            # data-search 存入文件名+所有任务令，方便 JS 搜索
            search_key = (f + " ".join(ep_list)).lower()
            
            items_html += f'''
                <div class="item" onclick="view(this, '{safe_f}')" data-search="{search_key}">
                    <b>{f[:45]}</b>
                    <div style="margin:5px 0;">{tags}</div>
                    <span class="time">🕒 {mt}</span>
                </div>
            '''

        full_html = f"""
        <!DOCTYPE html><html><head><meta charset="utf-8">
        <style>
            body {{ margin: 0; display: flex; height: 100vh; font-family: 'Segoe UI', sans-serif; overflow: hidden; background: #f3f2f1; }}
            .sidebar {{ width: 350px; background: #fff; border-right: 1px solid #ddd; display: flex; flex-direction: column; box-shadow: 2px 0 5px rgba(0,0,0,0.05); }}
            .header {{ background: #107c10; color: white; padding: 15px; }}
            .search-bar {{ padding: 10px; border-bottom: 1px solid #eee; }}
            .search-bar input {{ width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box; }}
            .list {{ flex: 1; overflow-y: auto; }}
            .item {{ padding: 12px 15px; border-bottom: 1px solid #f3f2f1; cursor: pointer; }}
            .item:hover {{ background: #f9f9f9; }}
            .item.active {{ background: #eff6fc; border-left: 5px solid #107c10; }}
            .ep-tag {{ background: #107c10; color: white; padding: 1px 5px; font-size: 10px; margin-right: 5px; border-radius: 3px; font-family: Consolas; }}
            .time {{ font-size: 11px; color: #999; }}
            .preview {{ flex: 1; background: #fff; position: relative; }}
            iframe {{ width: 100%; height: 100%; border: none; }}
            .footer {{ font-size: 10px; padding: 10px; text-align: center; color: #999; border-top: 1px solid #eee; }}
        </style>
        <script>
            function doSearch(e) {{
                let val = document.getElementById('q').value.toLowerCase();
                let items = document.querySelectorAll('.item');
                let match = null;
                items.forEach(i => {{
                    if(i.getAttribute('data-search').includes(val)) {{ i.style.display='block'; if(!match) match=i; }}
                    else {{ i.style.display='none'; }}
                }});
                if(e.keyCode === 13 && match) match.click();
            }}
            function view(el, url) {{
                document.querySelectorAll('.item').forEach(i => i.classList.remove('active'));
                el.classList.add('active');
                document.getElementById('f').src = url;
            }}
        </script></head>
        <body>
            <div class="sidebar">
                <div class="header">📫 RD 邮件分发看板 | EP 任务令版<br><small>最后更新: {self.last_sync_time}</small></div>
                <div class="search-bar"><input id="q" onkeyup="doSearch(event)" placeholder="🔍 搜索 EP 任务令或标题..."></div>
                <div class="list">{items_html}</div>
                <div class="footer">{self.copyright_text}</div>
            </div>
            <div class="preview"><iframe id="f"></iframe></div>
        </body></html>
        """
        try:
            with open(os.path.join(self.share_dir, "index.html"), "w", encoding="utf-8") as f_idx:
                f_idx.write(full_html)
            self.add_log("📊 看板索引重建完成")
        except: pass

if __name__ == "__main__":
    app = QApplication(sys.argv); ex = OutlookMHTMaster(); ex.show(); sys.exit(app.exec_())
