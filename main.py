import sys, os, re, time, subprocess, base64, email, json
import pandas as pd
import openpyxl
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QTextEdit, 
                             QSystemTrayIcon, QMenu, QAction, QStyle)
from PyQt5.QtCore import QTimer, Qt

class OutlookMHTMaster(QWidget):
    def __init__(self):
        super().__init__()
        # --- 默认参数设置 ---
        self.share_dir = r'\\10.1.93.32\DT_HU_RDteam_F\视频\Z\ZOUQIU\paican'
        self.target_kw = 'EDFA' 
        self.tag_regex = r'\bEP[A-Z0-9]{9}\b' 
        self.interval_min = 10     
        self.web_refresh_sec = 60  
        self.sync_count = 3       
        self.start_hour = 9       
        self.end_hour = 18        
        self.theme_color = "#107c10" 
        self.web_title = "EDFA 排产看板"
        self.web_sub_title = "Excel 原生排版优化版"
        self.copyright_text = "© 2024-2026 R1231685 | 技术支持"
        
        self.init_ui()
        self.init_tray()
        self.sync_timer = QTimer(self)
        self.sync_timer.timeout.connect(self.run_cycle)
        QTimer.singleShot(2000, self.run_cycle)

    def init_ui(self):
        self.setWindowTitle("EDFA 看板后台 V45.2")
        self.resize(500, 850)
        layout = QVBoxLayout()
        layout.setContentsMargins(15, 15, 15, 15)

        def quick_edit(label, val, attr):
            l = QHBoxLayout(); lb = QLabel(label); lb.setFixedWidth(110); l.addWidget(lb)
            edit = QLineEdit(str(val)); setattr(self, attr, edit); l.addWidget(edit); layout.addLayout(l)

        quick_edit("📂 共享路径", self.share_dir, "ui_path")
        quick_edit("📧 邮件关键词", self.target_kw, "ui_kw")
        quick_edit("🔍 提取正则", self.tag_regex, "ui_regex")
        quick_edit("🚩 网页大标题", self.web_title, "ui_title")
        quick_edit("📝 网页小字备注", self.web_sub_title, "ui_subtitle")

        h1 = QHBoxLayout()
        h1.addWidget(QLabel("⏱ 同步频率(分)")); self.ui_freq = QLineEdit(str(self.interval_min)); h1.addWidget(self.ui_freq)
        h1.addWidget(QLabel("🌐 网页刷新(秒)")); self.ui_web_freq = QLineEdit(str(self.web_refresh_sec)); h1.addWidget(self.ui_web_freq)
        layout.addLayout(h1)

        h2 = QHBoxLayout()
        h2.addWidget(QLabel("🔢 抓取数")); self.ui_count = QLineEdit(str(self.sync_count)); h2.addWidget(self.ui_count)
        h2.addWidget(QLabel("⏰ 时段")); self.ui_start = QLineEdit(str(self.start_hour)); h2.addWidget(self.ui_start)
        h2.addWidget(QLabel("-")); self.ui_end = QLineEdit(str(self.end_hour)); h2.addWidget(self.ui_end)
        layout.addLayout(h2)

        quick_edit("🎨 主题颜色", self.theme_color, "ui_color")
        quick_edit("🔒 版权内容", self.copyright_text, "ui_copy")

        self.btn_apply = QPushButton("🚀 立即全量解析并同步"); self.btn_apply.setFixedHeight(50)
        self.btn_apply.clicked.connect(self.apply_settings); layout.addWidget(self.btn_apply)
        
        self.log_area = QTextEdit(); self.log_area.setReadOnly(True); layout.addWidget(self.log_area)
        self.setLayout(layout); self.restyle()

    def init_tray(self):
        self.tray = QSystemTrayIcon(self); self.tray.setIcon(self.style().standardIcon(QStyle.SP_ComputerIcon))
        tm = QMenu(); tm.addAction("显示", self.showNormal); tm.addAction("退出", QApplication.instance().quit)
        self.tray.setContextMenu(tm); self.tray.show()
        self.tray.activated.connect(lambda r: self.showNormal() if r == QSystemTrayIcon.DoubleClick else None)

    def closeEvent(self, event):
        if self.tray.isVisible(): self.hide(); event.ignore()

    def restyle(self):
        c = self.ui_color.text().strip() or "#107c10"
        self.setStyleSheet(f"QPushButton{{background:{c};color:white;font-weight:bold;border-radius:4px;}}")

    def add_log(self, txt): self.log_area.append(f"[{time.strftime('%H:%M:%S')}] {str(txt)}")
    
    def apply_settings(self): 
        self.restyle(); self.add_log("⚙️ 配置同步并重绘..."); self.run_cycle()

    def run_cycle(self):
        now_h = int(time.strftime("%H"))
        try: s, e = int(self.ui_start.text()), int(self.ui_end.text())
        except: s, e = 9, 18
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
            $it = $ol.GetNamespace("MAPI").GetDefaultFolder(6).Items | Where-Object {{ $_.ReceivedTime -gt (Get-Date).AddDays(-7) -and ($_.Subject -like "*{k}*") }} | Sort-Object ReceivedTime -Descending | Select-Object -First {c_num}
            foreach($m in $it) {{
                $n = ($m.Subject -replace '[\\x00-\\x1f\\\\/:*?"<>|]', '_').Trim()
                $m.SaveAs((Join-Path "{d}" "$n.mht"), 10)
            }}
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
        
        # 1. 邮件 HTML 转换
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
        
        # 2. 核心日历解析 (OpenPyXL 逻辑)
        cal_html = "<p style='color:red;'>未在目录找到 [2026日历.xlsx]</p>"
        for f_name in os.listdir(d):
            if "2026日历" in f_name and f_name.lower().endswith('.xlsx'):
                try:
                    wb = openpyxl.load_workbook(os.path.join(d, f_name), data_only=True)
                    ws = wb.active
                    rows_html = ""
                    for i, row in enumerate(ws.iter_rows(values_only=True)):
                        if not any(row): continue 
                        row_content = ""
                        for cell_val in row:
                            val = "" if cell_val is None else str(cell_val)
                            style = "style='color:red;font-weight:bold;background:#fff0f0;'" if ("六" in val or "日" in val) else ""
                            tag = "th" if i == 0 else "td"
                            row_content += f"<{tag} {style}>{val}</{tag}>"
                        rows_html += f"<tr>{row_content}</tr>"
                    cal_html = f"<table class='cal-table'>{rows_html}</table>"
                    self.add_log(f"📅 日历已通过 Py 解析: {f_name}")
                    break
                except Exception as e: self.add_log(f"解析失败: {e}")
        self.build_index(cal_html)

    def build_index(self, cal_html):
        d = self.ui_path.text().strip()
        all_files = [f for f in os.listdir(d) if f.endswith('.html') and f != 'index.html']
        all_files.sort(key=lambda x: os.path.getmtime(os.path.join(d, x)), reverse=True)
        
        c = self.ui_color.text().strip() or "#107c10"
        t1, t2 = self.ui_title.text().strip(), self.ui_subtitle.text().strip()
        cp = self.ui_copy.text().strip()
        try: w_ref = int(self.ui_web_freq.text())
        except: w_ref = 60
        
        items_html, mails_content_html = "", ""
        regex_ptr = self.ui_regex.text().strip()

        for i, f in enumerate(all_files):
            file_path = os.path.join(d, f)
            mtime = time.strftime('%m-%d %H:%M', time.localtime(os.path.getmtime(file_path)))
            with open(file_path, 'r', encoding='utf-8') as tf: content = tf.read()
            tags = " ".join(list(set(re.findall(regex_ptr, content))))
            active_cls = "active" if i == 0 else ""
            items_html += f'<div class="mail-item {active_cls}" onclick="showMail(\'{i}\', this)" data-tags="{tags}"><b>{f[:-5]}</b><br><small>🕒 {mtime}</small></div>'
            disp = "block" if i == 0 else "none"
            mails_content_html += f'<div id="mail-{i}" class="mail-body" style="display:{disp}">{content}</div>'

        full_html = f"""
        <html><head><meta charset="utf-8"><meta http-equiv="refresh" content="{w_ref}"><title>{t1}</title>
        <style>
            body {{ font-family: 'Segoe UI', sans-serif; margin: 0; display: flex; height: 100vh; overflow: hidden; background:#f3f2f1; }}
            .sidebar {{ width: 320px; background: white; border-right: 1px solid #ddd; display: flex; flex-direction: column; flex-shrink: 0; }}
            .header {{ padding: 15px; background: {c}; color: white; }}
            .search-box {{ padding: 10px; border-bottom: 1px solid #eee; position: relative; }}
            .search-box input {{ width: 100%; padding: 8px 30px 8px 8px; border: 1px solid #ccc; border-radius: 4px; }}
            .clear-btn {{ position: absolute; right: 18px; top: 16px; cursor: pointer; color: #ccc; font-weight: bold; display:none; }}
            .mail-list {{ flex: 1; overflow-y: auto; }}
            .mail-item {{ padding: 12px; border-bottom: 1px solid #eee; cursor: pointer; }}
            .mail-item.active {{ border-left: 5px solid {c}; background: #f0f9f0; }}
            .content {{ flex: 1; display: flex; flex-direction: column; min-width: 0; }}
            .cal-section {{ height: 220px; overflow: auto; background: #fff; border-bottom: 2px solid {c}; padding: 10px; }}
            .mail-display {{ flex: 1; overflow: auto; padding: 20px; background: #fff; }}
            .cal-table {{ border-collapse: collapse; font-size: 12px; min-width: max-content; }}
            .cal-table th, .cal-table td {{ border: 1px solid #ddd; padding: 6px 12px; text-align: center; white-space: nowrap; }}
            .mail-body {{ min-width: max-content; }}
        </style></head>
        <body>
            <div class="sidebar">
                <div class="header"><b>{t1}</b><br><small>{t2}</small></div>
                <div class="search-box">
                    <input type="text" id="s" placeholder="搜索..." onkeyup="filter()">
                    <span id="cb" class="clear-btn" onclick="cls()">×</span>
                </div>
                <div class="mail-list" id="ml">{items_html}</div>
            </div>
            <div class="content">
                <div class="cal-section">{cal_html}</div>
                <div class="mail-display">{mails_content_html}</div>
            </div>
            <script>
                function showMail(id, el) {{
                    document.querySelectorAll('.mail-body').forEach(b => b.style.display = 'none');
                    document.querySelectorAll('.mail-item').forEach(i => i.classList.remove('active'));
                    document.getElementById('mail-'+id).style.display = 'block';
                    el.classList.add('active');
                }}
                function filter() {{
                    let v = document.getElementById('s').value.toUpperCase();
                    document.getElementById('cb').style.display = v ? 'block' : 'none';
                    document.querySelectorAll('.mail-item').forEach(item => {{
                        let txt = item.innerText.toUpperCase() + item.getAttribute('data-tags').toUpperCase();
                        item.style.display = txt.includes(v) ? "block" : "none";
                    }});
                }}
                function cls() {{ document.getElementById('s').value=''; filter(); }}
            </script>
        </body></html>"""
        with open(os.path.join(d, "index.html"), 'w', encoding='utf-8') as f: f.write(full_html)
        self.add_log("✅ 网页更新完成")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = OutlookMHTMaster()
    window.show()
    sys.exit(app.exec_())
