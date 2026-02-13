import sys, os, re, time, subprocess, tempfile, urllib.parse, base64, email
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QTextEdit, QSystemTrayIcon, QMenu)
from PyQt5.QtCore import QTimer, Qt

class OutlookMHTMaster(QWidget):
    def __init__(self):
        super().__init__()
        self.share_dir = r'\\10.1.93.32\DT_HU_RDteam_F\视频\Z\ZOUQIU\paican'
        self.target_kw = 'EDFA' 
        self.interval_min = 10     
        self.tmp_log = os.path.join(tempfile.gettempdir(), "sync_v16_res.txt")
        self.init_ui()
        self.init_tray()
        self.sync_timer = QTimer(self)
        self.sync_timer.timeout.connect(self.run_cycle)
        QTimer.singleShot(2000, self.run_cycle)

    def init_ui(self):
        self.setWindowTitle("RD 邮件同步 V16.0 (全正文搜索版)")
        self.resize(620, 600)
        layout = QVBoxLayout()
        form = QVBoxLayout()
        self.edit_path = QLineEdit(self.share_dir); form.addWidget(QLabel("📂 共享路径:")); form.addWidget(self.edit_path)
        self.edit_kw = QLineEdit(self.target_kw); form.addWidget(QLabel("📧 关键词:")); form.addWidget(self.edit_kw)
        layout.addLayout(form)
        self.btn_apply = QPushButton("🚀 同步并建立全文索引"); self.btn_apply.setStyleSheet("background:#107c10;color:white;padding:12px;font-weight:bold;"); self.btn_apply.clicked.connect(self.apply_settings); layout.addWidget(self.btn_apply)
        self.log_area = QTextEdit(); self.log_area.setReadOnly(True); self.log_area.setStyleSheet("background:#1e1e1e;color:#00ff00;"); layout.addWidget(self.log_area); self.setLayout(layout)

    def init_tray(self):
        self.tray = QSystemTrayIcon(self); self.tray.setIcon(self.style().standardIcon(21)); menu = QMenu(); menu.addAction("显示", self.showNormal); menu.addAction("退出", QApplication.instance().quit); self.tray.setContextMenu(menu); self.tray.show()

    def add_log(self, text):
        t = time.strftime('%H:%M:%S')
        msg = str(text).replace('\x00', '')
        self.log_area.append(f"[{t}] {msg}")

    def apply_settings(self):
        self.share_dir, self.target_kw = self.edit_path.text().strip(), self.edit_kw.text().strip()
        self.run_cycle()

    def run_cycle(self):
        self.run_shell_logic()
        self.sync_timer.start(10 * 60000)

    def run_shell_logic(self):
        ps_dir = self.share_dir.replace('"', '""')
        ps_kw = self.target_kw.replace('"', '""')
        ps_tmp = self.tmp_log.replace('"', '""')
        ps_script = f"""
        try {{
            if (!(Test-Path "{ps_dir}")) {{ New-Item -ItemType Directory -Path "{ps_dir}" -Force | Out-Null }}
            $ol = New-Object -ComObject Outlook.Application
            $ns = $ol.GetNamespace("MAPI")
            $it = $ns.GetDefaultFolder(6).Items | Where-Object {{ 
                $_.ReceivedTime -gt (Get-Date).AddDays(-3) -and ($_.Subject -like "*{ps_kw}*" -or $_.SenderName -like "*{ps_kw}*") 
            }} | Sort-Object ReceivedTime -Descending | Select-Object -First 1
            if ($it) {{
                $name = $it.Subject -replace '[\\x00-\\x1f\\\\/:*?"<>|]', '_'
                $path = Join-Path "{ps_dir}" "$($name.Trim()).mht"
                if (!(Test-Path $path)) {{ $it.SaveAs($path, 10); "SUCCESS|$name" | Out-File "{ps_tmp}" -Encoding utf8 }}
                else {{ "EXISTS" | Out-File "{ps_tmp}" -Encoding utf8 }}
            }} else {{ "NOTFOUND" | Out-File "{ps_tmp}" -Encoding utf8 }}
        }} catch {{ "ERROR|$($_.Exception.Message)" | Out-File "{ps_tmp}" -Encoding utf8
        }} finally {{ if ($ol) {{ [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ol) | Out-Null }} }}
        """
        try:
            ps_b64 = base64.b64encode(ps_script.encode('utf-16-le')).decode('ascii')
            subprocess.run(["powershell", "-NoProfile", "-EncodedCommand", ps_b64], creationflags=0x08000000, timeout=60)
            self.convert_and_index()
        except Exception as e: self.add_log(f"同步故障: {e}")

    def convert_and_index(self):
        """扫描MHT，生成带全文索引的静态HTML"""
        if not os.path.exists(self.share_dir): return
        mht_files = [f for f in os.listdir(self.share_dir) if f.endswith('.mht')]
        items_html = ""
        
        for f in mht_files:
            mht_path = os.path.join(self.share_dir, f)
            html_name = f.replace('.mht', '.static.html')
            html_path = os.path.join(self.share_dir, html_name)
            mt = time.strftime("%m-%d %H:%M", time.localtime(os.path.getmtime(mht_path)))
            
            ep_list, full_text_index, mail_body = [], "", ""
            try:
                with open(mht_path, 'rb') as fp:
                    msg = email.message_from_binary_file(fp)
                    for part in msg.walk():
                        if part.get_content_type() == "text/html":
                            payload = part.get_payload(decode=True)
                            for enc in ['utf-8', 'gbk', 'gb18030']:
                                try:
                                    mail_body = payload.decode(enc)
                                    # 提取全文索引（去标签，去换行）
                                    full_text_index = re.sub('<[^<]+?>', '', mail_body).replace('\n', ' ').replace('\r', '').replace('"', "'")
                                    # 提取EP任务令
                                    ep_list = list(set(re.findall(r'\bEP[A-Z0-9]{9}\b', full_text_index, re.I)))
                                    break
                                except: continue
                if mail_body:
                    with open(html_path, 'w', encoding='utf-8') as hw: hw.write(mail_body)
            except: pass
            
            tags = "".join([f'<span style="background:#107c10;color:white;padding:2px 5px;margin-right:5px;border-radius:3px;font-size:10px;">{c}</span>' for c in ep_list[:2]])
            safe_html_f = urllib.parse.quote(html_name)
            # 搜索权重：文件名 + EP任务令 + 全正文内容
            search_key = (f + " " + " ".join(ep_list) + " " + full_text_index[:1500]).lower()
            
            items_html += f'''<div class="item" onclick="viewMail(this, '{safe_html_f}')" data-search="{search_key}" style="padding:12px;border-bottom:1px solid #eee;cursor:pointer;">
                <b style="font-size:13px;color:#333;">{f[:40]}</b><br>{tags}<br><span style="color:#999;font-size:11px;">🕒 {mt}</span></div>'''

        full_html = f"""
        <!DOCTYPE html><html><head><meta charset='utf-8'>
        <style>
            body {{ margin:0; display:flex; height:100vh; font-family:'Segoe UI',sans-serif; background:#f4f4f4; overflow:hidden; }}
            .sidebar {{ width:320px; background:#fff; border-right:1px solid #ddd; display:flex; flex-direction:column; box-shadow:2px 0 5px rgba(0,0,0,0.05); }}
            .search-box {{ padding:12px; border-bottom:1px solid #eee; background:#fafafa; }}
            .search-box input {{ width:100%; padding:10px; border:1px solid #ddd; border-radius:4px; box-sizing:border-box; outline:none; }}
            .search-box input:focus {{ border-color:#107c10; }}
            .list {{ flex:1; overflow-y:auto; }}
            .item:hover {{ background:#f9f9f9; }}
            .item.active {{ background:#dff6dd; border-left:5px solid #107c10; }}
            .preview {{ flex:1; background:#fff; position:relative; }}
            iframe {{ width:100%; height:100%; border:none; }}
            .empty {{ position:absolute; top:50%; left:50%; transform:translate(-50%,-50%); color:#aaa; text-align:center; }}
        </style>
        <script>
            function doSearch(v) {{
                v = v.toLowerCase();
                let match = null;
                document.querySelectorAll('.item').forEach(i => {{
                    let isMatch = i.getAttribute('data-search').includes(v);
                    i.style.display = isMatch ? 'block' : 'none';
                    if(isMatch && !match) match = i;
                }});
                // 回车逻辑由 input 的 onkeyup 处理
            }}
            function viewMail(el, url) {{
                document.querySelectorAll('.item').forEach(i => i.classList.remove('active'));
                el.classList.add('active');
                document.getElementById('f').src = url;
                document.getElementById('msg').style.display = 'none';
            }}
        </script></head>
        <body>
            <div class="sidebar">
                <div style="background:#107c10;color:white;padding:15px;"><b>📫 RD 全文看板 Pro</b></div>
                <div class="search-box">
                    <input id="q" onkeyup="if(event.keyCode==13) {{ doSearch(this.value); }} else {{ doSearch(this.value); }}" placeholder="🔍 搜正文、表格、任务令...">
                </div>
                <div class="list">{items_html}</div>
            </div>
            <div class="preview">
                <div id="msg" class="empty"><h3>📂 请选择邮件</h3><p>支持全文检索及回车定位</p></div>
                <iframe id="f"></iframe>
            </div>
        </body></html>
        """
        with open(os.path.join(self.share_dir, "index.html"), "w", encoding="utf-8") as f_idx: f_idx.write(full_html)
        self.add_log(f"✅ 全文索引重建完成 ({len(mht_files)} 封)")

if __name__ == "__main__":
    app = QApplication(sys.argv); ex = OutlookMHTMaster(); ex.show(); sys.exit(app.exec_())
