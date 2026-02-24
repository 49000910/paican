import sys, os, re, time, subprocess, base64, email, json
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QTextEdit, 
                             QSystemTrayIcon, QMenu, QAction, QStyle)
from PyQt5.QtCore import QTimer, Qt

class OutlookMHTMaster(QWidget):
    def __init__(self):
        super().__init__()
        # --- 原始参数保持不变 ---
        self.share_dir = r'\\10.1.93.32\DT_HU_RDteam_F\视频\Z\ZOUQIU\paican'
        self.target_kw = 'EDFA' 
        self.tag_regex = r'\bEP[A-Z0-9]{9}\b' 
        self.interval_min = 10     
        self.sync_count = 3       
        self.start_hour = 9       
        self.end_hour = 12        
        self.theme_color = "#107c10" 
        self.web_title = "EDFA 排产看板"
        self.web_sub_title = "Excel日历原子更新版"
        self.copyright_text = "© 2024-2026 R1231685 | 技术支持"
        
        self.init_ui()
        self.init_tray()
        self.sync_timer = QTimer(self)
        self.sync_timer.timeout.connect(self.run_cycle)
        QTimer.singleShot(2000, self.run_cycle)

    def init_ui(self):
        self.setWindowTitle("EDFA 看板后台 V41.0")
        self.resize(500, 850)
        layout = QVBoxLayout()
        layout.setContentsMargins(15, 15, 15, 15)

        def quick_edit(label, val, attr):
            l = QHBoxLayout(); lb = QLabel(label); lb.setFixedWidth(100); l.addWidget(lb)
            edit = QLineEdit(str(val)); setattr(self, attr, edit); l.addWidget(edit); layout.addLayout(l)

        quick_edit("📂 共享路径", self.share_dir, "ui_path")
        quick_edit("📧 邮件关键词", self.target_kw, "ui_kw")
        quick_edit("🔍 提取正则", self.tag_regex, "ui_regex")
        quick_edit("🚩 网页大标题", self.web_title, "ui_title")
        quick_edit("📝 网页小字备注", self.web_sub_title, "ui_subtitle")
        
        h1 = QHBoxLayout()
        h1.addWidget(QLabel("⏱ 同步频率(分)")); self.ui_freq = QLineEdit(str(self.interval_min)); h1.addWidget(self.ui_freq)
        h1.addWidget(QLabel("🔢 抓取数")); self.ui_count = QLineEdit(str(self.sync_count)); h1.addWidget(self.ui_count)
        layout.addLayout(h1)

        h2 = QHBoxLayout()
        h2.addWidget(QLabel("⏰ 开始时")); self.ui_start = QLineEdit(str(self.start_hour)); h2.addWidget(self.ui_start)
        h2.addWidget(QLabel("⏰ 结束时")); self.ui_end = QLineEdit(str(self.end_hour)); h2.addWidget(self.ui_end)
        layout.addLayout(h2)

        quick_edit("🎨 主题颜色", self.theme_color, "ui_color")
        quick_edit("🔒 版权内容", self.copyright_text, "ui_copy")

        self.btn_apply = QPushButton("🚀 立即全量解析并同步")
        self.btn_apply.setFixedHeight(50); self.btn_apply.clicked.connect(self.apply_settings)
        layout.addWidget(self.btn_apply)

        self.log_area = QTextEdit(); self.log_area.setReadOnly(True)
        layout.addWidget(self.log_area); self.setLayout(layout); self.restyle()

    def init_tray(self):
        self.tray = QSystemTrayIcon(self)
        self.tray.setIcon(self.style().standardIcon(QStyle.SP_ComputerIcon))
        tm = QMenu(); tm.addAction("显示", self.showNormal); tm.addAction("退出", QApplication.instance().quit)
        self.tray.setContextMenu(tm); self.tray.show()

    def closeEvent(self, event):
        if self.tray.isVisible(): self.hide(); event.ignore()

    def restyle(self):
        c = self.ui_color.text().strip() or "#107c10"
        self.setStyleSheet(f"QPushButton{{background:{c};color:white;font-weight:bold;}} QTextEdit{{background:#1e1e1e;color:#0f0;}}")

    def add_log(self, txt): self.log_area.append(f"[{time.strftime('%H:%M:%S')}] {str(txt)}")
    def apply_settings(self): self.restyle(); self.add_log("⚙️ 配置下发..."); self.run_cycle()

    def run_cycle(self):
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
            $it = $ol.GetNamespace("MAPI").GetDefaultFolder(6).Items | Where-Object {{ $_.ReceivedTime -gt (Get-Date).AddDays(-3) -and ($_.Subject -like "*{k}*") }} | Sort-Object ReceivedTime -Descending | Select-Object -First {c_num}
            foreach($m in $it) {{
                $n = ($m.Subject -replace '[\\x00-\\x1f\\\\/:*?"<>|]', '_').Trim()
                $m.SaveAs((Join-Path "{d}" "$n.mht"), 10)
            }}
        }} catch {{ }}
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
                            with open(p_h, 'w', encoding='utf-8') as hw: hw.write(part.get_payload(decode=True).decode('utf-8','ignore'))
                os.remove(p_m)
            except: pass
        
        # 解析本地 Excel 日历
        cal_html = "<p style='color:red;'>未找到 2026日历.xlsx</p>"
        for f in os.listdir(d):
            if "2026日历" in f and f.endswith(('.xlsx', '.xls')):
                try:
                    df = pd.read_excel(os.path.join(d, f))
                    def color_weekend(v):
                        try:
                            dt = pd.to_datetime(v)
                            if dt.weekday() >= 5: return 'color:red; font-weight:bold; background-color:#fff0f0;'
                        except: pass
                        return ''
                    cal_html = df.style.applymap(color_weekend).to_html(classes='cal-table', index=False, na_rep='')
                    break
                except: pass
        self.build_index(cal_html)

    def build_index(self, cal_html):
        d = self.ui_path.text().strip()
        all_files = [f for f in os.listdir(d) if f.endswith('.html') and f != 'index.html']
        all_files.sort(key=lambda x: os.path.getmtime(os.path.join(d, x)), reverse=True)
        c, t1, t2, cp = self.ui_color.text().strip(), self.ui_title.text().strip(), self.ui_subtitle.text().strip(), self.ui_copy.text().strip()
        
        items_html, mails_data_html, search_db = "", "", {}
        for i, f in enumerate(all_files):
            p = os.path.join(d, f)
            with open(p, 'r', encoding='utf-8', errors='ignore') as fc: raw_h = fc.read()
            pure_text = " ".join(re.sub(r'<[^>]+>', ' ', raw_h).split()).lower()
            search_db[f"m_{i}"] = pure_text
            tags = list(set([x.upper() for x in re.findall(self.ui_regex.text().strip(), pure_text, re.I)]))
            tag_ui = "".join([f'<span class="et" onclick="fastGo(\'{x}\')">{x}</span>' for x in tags[:5]])
            items_html += f'<div class="item" id="li_m_{i}" onclick="jump(\'m_{i}\', this)"><div class="ti">{f[:-5]}</div><div class="tags">{tag_ui}</div></div>'
            mails_data_html += f'<div id="m_{i}" class="m-box"><div class="m-bar" style="border-left:5px solid {c}">{f[:-5]}</div><div class="m-body">{raw_h}</div></div>'

        db_b64 = base64.b64encode(json.dumps(search_db).encode('utf-8')).decode('ascii')
        
        # 网页模板：内置 60秒 自动刷新，移除所有图片请求
        index_tpl = f'''
        <!DOCTYPE html><html><head><meta charset="UTF-8">
        <meta http-equiv="refresh" content="60"> 
        <style>
            body {{ display:flex; height:100vh; margin:0; font-family:sans-serif; background:#f0f2f5; overflow:hidden; }}
            #side {{ width:380px; background:#fff; border-right:1px solid #ddd; display:flex; flex-direction:column; }}
            #main {{ flex:1; overflow-y:auto; padding:20px; scroll-behavior:smooth; }}
            .head {{ padding:15px; background:{c}; color:#fff; }}
            #q {{ width:100%; padding:8px; border:none; border-radius:4px; margin-top:10px; outline:none; }}
            .item {{ padding:10px; border-bottom:1px solid #eee; cursor:pointer; font-size:12px; }}
            .et {{ background:#e8f5e9; color:{c}; padding:1px 3px; border-radius:2px; font-size:10px; border:1px solid {c}; margin-right:3px; }}
            .m-box {{ background:#fff; margin-bottom:30px; border-radius:4px; box-shadow:0 1px 3px rgba(0,0,0,0.1); }}
            .m-bar {{ padding:10px; background:#fafafa; font-weight:bold; }}
            .cal-table {{ border-collapse:collapse; width:100%; font-size:12px; }}
            .cal-table th {{ background:#f5f5f5; border:1px solid #ddd; padding:6px; }}
            .cal-table td {{ border:1px solid #ddd; padding:6px; text-align:center; }}
            .btn-cal {{ margin:10px; padding:10px; background:#333; color:#fff; text-align:center; border-radius:4px; cursor:pointer; font-weight:bold; font-size:13px; }}
            #calModal {{ display:none; position:fixed; z-index:999; left:0; top:0; width:100%; height:100%; background:rgba(0,0,0,0.8); align-items:center; justify-content:center; }}
            .cal-card {{ background:#fff; width:90%; max-height:85%; padding:20px; border-radius:6px; overflow-y:auto; position:relative; }}
            mark {{ background:yellow; font-weight:bold; }}
            .active {{ background:#e8f5e9 !important; border-right:5px solid {c}; }}
        </style></head>
        <body>
            <div id="side">
                <div class="head"><strong>{t1}</strong><br><small>{t2}</small><input type="text" id="q" placeholder="输入 EP号 定位..." oninput="doSearch(this.value)"></div>
                <div style="flex:1; overflow-y:auto;">{items_html}</div>
                <div class="btn-cal" onclick="toggleCal(true)">📅 2026 华为工作日历 (Excel)</div>
                <div style="padding:10px; font-size:10px; color:#999; text-align:center;">{cp}</div>
            </div>
            <div id="main">{mails_data_html}</div>
            <div id="calModal" onclick="if(event.target==this) toggleCal(false)">
                <div class="cal-card"><span onclick="toggleCal(false)" style="position:absolute;right:15px;top:10px;cursor:pointer;font-size:20px;">&times;</span>
                <h3>📅 2026 华为工作日历 (本地数据)</h3>{cal_html}</div>
            </div>
            <script>
                const db = JSON.parse(atob("{db_b64}"));
                function toggleCal(s) {{ document.getElementById('calModal').style.display = s ? 'flex' : 'none'; }}
                function fastGo(v) {{ document.getElementById('q').value = v; doSearch(v); }}
                function doSearch(kw) {{
                    const v = kw.toLowerCase().trim(); let first = null;
                    Object.keys(db).forEach(id => {{
                        const match = db[id].includes(v); document.getElementById('li_'+id).style.display = match ? 'block' : 'none';
                        if(match && !first) first = id;
                    }});
                    if(first && v.length > 5) jump(first, document.getElementById('li_'+first));
                }}
                function jump(id, el) {{
                    const target = document.getElementById(id); const kw = document.getElementById('q').value.trim();
                    target.scrollIntoView({{ behavior: 'smooth', block: 'start' }});
                    document.querySelectorAll('.m-body').forEach(b => b.innerHTML = b.innerHTML.replace(/<mark>|<\/mark>/g, ""));
                    if(kw.length > 2) {{
                        const body = target.querySelector('.m-body');
                        const reg = new RegExp("("+kw+")", "gi");
                        body.innerHTML = body.innerHTML.replace(reg, "<mark>$1</mark>");
                    }}
                    document.querySelectorAll('.item').forEach(i => i.classList.remove('active')); el.classList.add('active');
                }}
            </script>
        </body></html>'''
        
        # 🔥 原子化保存 index.html：先写临时文件，再更名，解决浏览器加载时的转圈问题
        tmp_path = os.path.join(d, 'index_tmp.html')
        final_path = os.path.join(d, 'index.html')
        with open(tmp_path, 'w', encoding='utf-8') as f:
            f.write(index_tpl)
        if os.path.exists(final_path):
            os.remove(final_path)
        os.rename(tmp_path, final_path)
        self.add_log("🌍 看板网页已原子化更新")

if __name__ == "__main__":
    app = QApplication(sys.argv); win = OutlookMHTMaster(); win.show(); sys.exit(app.exec_())
