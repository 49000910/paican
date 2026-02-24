import sys, os, re, time, subprocess, base64, email, json
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QTextEdit, 
                             QSystemTrayIcon, QMenu, QAction, QStyle)
from PyQt5.QtCore import QTimer, Qt

class OutlookMHTMaster(QWidget):
    def __init__(self):
        super().__init__()
        # --- 默认参数设置（保留您的所有原始定义） ---
        self.share_dir = r'\\10.1.93.32\DT_HU_RDteam_F\视频\Z\ZOUQIU\paican'
        self.target_kw = 'EDFA' 
        self.tag_regex = r'\bEP[A-Z0-9]{9}\b' 
        self.interval_min = 10     
        self.sync_count = 3       
        self.start_hour = 9       
        self.end_hour = 12        
        self.theme_color = "#107c10" 
        self.web_title = "EDFA 排产看板"
        self.web_sub_title = "自动抓取 zouqiu@hauwei.com"
        self.copyright_text = "© 2024-2026 R1231685 | 技术支持"
        
        self.init_ui()
        self.init_tray()
        self.sync_timer = QTimer(self)
        self.sync_timer.timeout.connect(self.run_cycle)
        QTimer.singleShot(2000, self.run_cycle)

    def init_ui(self):
        self.setWindowTitle("EDFA 看板后台 V32.0 - 日历集成版")
        self.resize(500, 850)
        layout = QVBoxLayout()
        layout.setContentsMargins(15, 15, 15, 15)

        def quick_edit(label, val, attr):
            l = QHBoxLayout()
            lb = QLabel(label); lb.setFixedWidth(100); l.addWidget(lb)
            edit = QLineEdit(str(val)); setattr(self, attr, edit)
            l.addWidget(edit); layout.addLayout(l)

        quick_edit("📂 共享路径", self.share_dir, "ui_path")
        quick_edit("📧 邮件关键词", self.target_kw, "ui_kw")
        quick_edit("🔍 提取正则", self.tag_regex, "ui_regex")
        quick_edit("🚩 网页大标题", self.web_title, "ui_title")
        quick_edit("📝 网页小字备注", self.web_sub_title, "ui_subtitle")
        
        h1 = QHBoxLayout()
        h1.addWidget(QLabel("⏱ 频率(分)")); self.ui_freq = QLineEdit(str(self.interval_min)); h1.addWidget(self.ui_freq)
        h1.addWidget(QLabel("🔢 抓取数")); self.ui_count = QLineEdit(str(self.sync_count)); h1.addWidget(self.ui_count)
        layout.addLayout(h1)

        h2 = QHBoxLayout()
        h2.addWidget(QLabel("⏰ 开始时")); self.ui_start = QLineEdit(str(self.start_hour)); h2.addWidget(self.ui_start)
        h2.addWidget(QLabel("⏰ 结束时")); self.ui_end = QLineEdit(str(self.end_hour)); h2.addWidget(self.ui_end)
        layout.addLayout(h2)

        quick_edit("🎨 主题颜色", self.theme_color, "ui_color")
        quick_edit("🔒 版权内容", self.copyright_text, "ui_copy")

        self.btn_apply = QPushButton("🚀 立即部署看板 (含2026日历)")
        self.btn_apply.setFixedHeight(50); self.btn_apply.clicked.connect(self.apply_settings)
        layout.addWidget(self.btn_apply)

        self.log_area = QTextEdit(); self.log_area.setReadOnly(True)
        layout.addWidget(self.log_area); self.setLayout(layout); self.restyle()

    def init_tray(self):
        self.tray = QSystemTrayIcon(self)
        self.tray.setIcon(self.style().standardIcon(QStyle.SP_ComputerIcon))
        tm = QMenu(); tm.addAction("显示", self.showNormal); tm.addAction("退出", QApplication.instance().quit)
        self.tray.setContextMenu(tm); self.tray.show()

    def restyle(self):
        c = self.ui_color.text().strip() or "#107c10"
        self.setStyleSheet(f"QPushButton{{background:{c};color:white;font-weight:bold;border-radius:4px;}} QTextEdit{{background:#1e1e1e;color:#0f0;font-family:Consolas;}}")

    def add_log(self, txt): self.log_area.append(f"[{time.strftime('%H:%M:%S')}] {str(txt)}")

    def apply_settings(self): self.restyle(); self.add_log("⚙️ 配置已同步..."); self.run_cycle()

    def run_cycle(self):
        now_h = int(time.strftime("%H"))
        try: s, e = int(self.ui_start.text()), int(self.ui_end.text())
        except: s, e = 9, 12
        if not (s <= now_h < e):
            self.add_log(f"💤 非活跃时段 ({now_h}点)"); self.sync_timer.start(30 * 60000); return
        self.run_shell()
        try: f = int(self.ui_freq.text()); self.sync_timer.start(f * 60000)
        except: self.sync_timer.start(600000)

    def run_shell(self):
        d, k = self.ui_path.text().replace('"', '""'), self.ui_kw.text().replace('"', '""')
        try: c_num = int(self.ui_count.text())
        except: c_num = 3
        ps_cmd = f"""
        try {{
            $ol = New-Object -ComObject Outlook.Application
            $it = $ol.GetNamespace("MAPI").GetDefaultFolder(6).Items | Where-Object {{ $_.ReceivedTime -gt (Get-Date).AddDays(-3) -and ($_.Subject -like "*{k}*" -or $_.SenderName -like "*{k}*") }} | Sort-Object ReceivedTime -Descending | Select-Object -First {c_num}
            if ($it) {{ foreach($m in $it) {{ $n = ($m.Subject -replace '[\\x00-\\x1f\\\\/:*?"<>|]', '_').Trim(); $p = Join-Path "{d}" "$n.mht"; if (!(Test-Path $p)) {{ $m.SaveAs($p, 10) }} }} }}
        }} catch {{ }} finally {{ if ($ol) {{ [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ol) | Out-Null }} }}
        """
        try:
            ps_b64 = base64.b64encode(ps_cmd.encode('utf-16-le')).decode('ascii')
            subprocess.run(["powershell", "-NoProfile", "-EncodedCommand", ps_b64], creationflags=0x08000000, timeout=60)
            self.process_web()
        except Exception as e: self.add_log(f"ERR: {e}")

    def process_web(self):
        d = self.ui_path.text().strip()
        if not os.path.exists(d): return
        for f in [x for x in os.listdir(d) if x.endswith('.mht')]:
            p_m, p_h = os.path.join(d, f), os.path.join(d, f.replace('.mht', '.html'))
            try:
                with open(p_m, 'rb') as fp:
                    msg = email.message_from_binary_file(fp)
                    for part in msg.walk():
                        if part.get_content_type() == "text/html":
                            raw = part.get_payload(decode=True).decode('utf-8','ignore')
                            with open(p_h, 'w', encoding='utf-8') as hw: hw.write(raw)
                            break
                os.remove(p_m)
            except: pass
        self.build_index()

    def build_index(self):
        d = self.ui_path.text().strip()
        all_files = [f for f in os.listdir(d) if f.endswith('.html') and f != 'index.html']
        all_files.sort(key=lambda x: os.path.getmtime(os.path.join(d, x)), reverse=True)
        
        c, t1, t2, cp = self.ui_color.text().strip(), self.ui_title.text().strip(), self.ui_subtitle.text().strip(), self.ui_copy.text().strip()
        reg_pattern = self.ui_regex.text().strip() or r'\bEP[A-Z0-9]{9}\b'

        items_html, mails_data_html = "", ""
        search_db = {} 

        for i, f in enumerate(all_files[:150]):
            p = os.path.join(d, f)
            with open(p, 'r', encoding='utf-8', errors='ignore') as fc: raw_h = fc.read()
            text_only = re.sub(r'<(style|script)[^>]*>.*?<\/\1>', '', raw_h, flags=re.DOTALL|re.IGNORECASE)
            text_only = re.sub(r'<[^>]+>', ' ', text_only)
            pure_text = " ".join(text_only.split()).lower()
            search_db[f"m_{i}"] = pure_text

            tags_raw = re.findall(reg_pattern, pure_text, re.I)
            tags_upper = list(set([x.upper() for x in tags_raw]))
            tag_ui = "".join([f'<span class="et" onclick="fastGo(\'{x}\')">{x}</span>' for x in tags_upper[:5]])

            items_html += f'''<div class="item" id="li_m_{i}" onclick="jump('m_{i}', this)">
                <div class="ti">{f[:-5]}</div><div class="tags">{tag_ui}</div>
                <div class="tm">{time.strftime("%m-%d %H:%M", time.localtime(os.path.getmtime(p)))}</div></div>'''
            
            mails_data_html += f'''<div id="m_{i}" class="m-box">
                <div class="m-bar" style="border-left:5px solid {c}">{f[:-5]}</div>
                <div class="m-body">{raw_h}</div></div>'''

        db_b64 = base64.b64encode(json.dumps(search_db).encode('utf-8')).decode('ascii')

        index_tpl = f'''
        <!DOCTYPE html><html><head><meta charset="UTF-8"><title>{t1}</title>
        <style>
            body {{ display:flex; height:100vh; margin:0; font-family:sans-serif; background:#f0f2f5; overflow:hidden; }}
            #side {{ width:380px; background:#fff; border-right:1px solid #ddd; display:flex; flex-direction:column; }}
            #main {{ flex:1; overflow-y:auto; padding:20px; scroll-behavior:smooth; }}
            .head {{ padding:15px; background:{c}; color:#fff; }}
            #q {{ width:100%; padding:10px; border:none; border-radius:4px; margin-top:8px; outline:none; }}
            .item {{ padding:12px; border-bottom:1px solid #eee; cursor:pointer; }}
            .ti {{ font-size:13px; font-weight:bold; color:#222; }}
            .et {{ background:#e8f5e9; color:{c}; padding:1px 4px; border-radius:3px; font-size:10px; border:1px solid {c}; margin:2px 4px 0 0; display:inline-block; }}
            .tm {{ font-size:11px; color:#999; margin-top:4px; }}
            .m-box {{ background:#fff; margin-bottom:30px; border-radius:4px; box-shadow:0 1px 3px rgba(0,0,0,0.1); }}
            .m-bar {{ padding:12px; background:#fafafa; font-weight:bold; }}
            .m-body {{ padding:15px; font-size:14px; overflow-x:auto; }}
            mark {{ background: yellow; color: black; font-weight:bold; }}
            .active {{ background:#e8f5e9 !important; border-right:5px solid {c}; }}
            /* 日历样式 */
            .cal-btn {{ margin: 10px; padding: 10px; background: #333; color: #fff; text-align: center; border-radius: 4px; cursor: pointer; font-size: 13px; font-weight: bold; border: none; }}
            .cal-btn:hover {{ background: #000; }}
            #calModal {{ display: none; position: fixed; z-index: 9999; left: 0; top: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.85); align-items: center; justify-content: center; flex-direction: column; }}
            .cal-content {{ position: relative; max-width: 85%; background: #fff; padding: 20px; border-radius: 8px; text-align: center; }}
            #calModal img {{ max-width: 100%; max-height: 65vh; border: 1px solid #ddd; }}
            .cal-info {{ margin-top: 15px; text-align: left; font-size: 13px; color: #444; column-count: 2; border-top: 1px solid #eee; padding-top: 10px; }}
            .close-cal {{ position: absolute; top: -35px; right: -5px; color: #fff; font-size: 35px; cursor: pointer; }}
        </style></head>
        <body>
            <div id="side">
                <div class="head"><strong>{t1}</strong><br><small>{t2}</small>
                <input type="text" id="q" placeholder="输入 EP号 自动定位并高亮..." oninput="doSearch(this.value)"></div>
                <div style="flex:1; overflow-y:auto;">{items_html}</div>
                <button class="cal-btn" onclick="toggleCal(true)">📅 2026 华为工作日历</button>
                <div style="padding:10px; font-size:10px; color:#999; text-align:center;">{cp}</div>
            </div>
            <div id="main">{mails_data_html}</div>
            <div id="calModal" onclick="if(event.target==this) toggleCal(false)">
                <div class="cal-content">
                    <span class="close-cal" onclick="toggleCal(false)">&times;</span>
                    <h3>2026 华为工作日历预告</h3>
                    <img src="2026cal.jpg" onerror="this.src='https://raw.githubusercontent.com'; this.alt='本地图未找到，加载备用图';" />
                    <div class="cal-info">
                        <strong>春节：</strong>2/15-2/23放假(9天)，2/14及2/28上班<br>
                        <strong>劳动节：</strong>5/1-5/5放假，5/9上班<br>
                        <strong>国庆：</strong>10/1-10/7放假，9/20及10/10上班<br>
                        <strong>元旦/清明/端午/中秋：</strong>按标准法定安排
                    </div>
                </div>
            </div>
            <script>
                const db = JSON.parse(atob("{db_b64}"));
                function toggleCal(s) {{ document.getElementById('calModal').style.display = s ? 'flex' : 'none'; }}
                function fastGo(v) {{ document.getElementById('q').value = v; doSearch(v); }}
                function doSearch(kw) {{
                    const v = kw.toLowerCase().trim();
                    let first = null;
                    Object.keys(db).forEach(id => {{
                        const match = db[id].includes(v);
                        document.getElementById('li_'+id).style.display = match ? 'block' : 'none';
                        if(match && !first) first = id;
                    }});
                    if(first && v.length > 5) jump(first, document.getElementById('li_'+first));
                }}
                function jump(id, el) {{
                    const target = document.getElementById(id);
                    const kw = document.getElementById('q').value.trim();
                    target.scrollIntoView({{ behavior: 'smooth', block: 'start' }});
                    document.querySelectorAll('.m-body').forEach(b => b.innerHTML = b.innerHTML.replace(/<mark>|<\/mark>/g, ""));
                    if(kw.length > 2) {{
                        const body = target.querySelector('.m-body');
                        const reg = new RegExp("("+kw+")", "gi");
                        body.innerHTML = body.innerHTML.replace(reg, "<mark>$1</mark>");
                    }}
                    document.querySelectorAll('.item').forEach(i => i.classList.remove('active'));
                    el.classList.add('active');
                }}
            </script>
        </body></html>'''
        with open(os.path.join(d, 'index.html'), 'w', encoding='utf-8') as f: f.write(index_tpl)
        self.add_log(f"🌍 看板已更新 (集成2026日历)")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = OutlookMHTMaster(); win.show()
    sys.exit(app.exec_())
