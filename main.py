import sys, os, re, time, subprocess, base64, email, json, datetime
import pandas as pd
import openpyxl
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QTextEdit, 
                             QSystemTrayIcon, QMenu, QAction, QStyle)
from PyQt5.QtCore import QTimer, Qt

class OutlookMHTMaster(QWidget):
    def __init__(self):
        super().__init__()
        # --- 默认参数 (完全还原您的定义) ---
        self.share_dir = r'\\10.1.93.32\DT_HU_RDteam_F\视频\Z\ZOUQIU\paican'
        self.target_kw = 'EDFA' 
        self.tag_regex = r'\bEP[A-Z0-9]{9}\b' 
        self.interval_min = 10     
        self.web_refresh_sec = 60  
        self.sync_count = 3       
        self.start_hour = 9       
        self.end_hour = 12        
        self.theme_color = "#107c10" 
        self.web_title = "EDFA 看板"
        self.web_sub_title = "Excel 原生排版优化版"
        self.copyright_text = "© 2024-2026 R1231685 | 技术支持"
        
        self.init_ui()
        self.init_tray()
        self.sync_timer = QTimer(self)
        self.sync_timer.timeout.connect(self.run_cycle)
        QTimer.singleShot(2000, self.run_cycle)

    def init_ui(self):
        self.setWindowTitle("EDFA 看板后台 V49.5")
        self.resize(520, 900)
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
    def apply_settings(self): self.restyle(); self.add_log("⚙️ 重新解析中..."); self.run_cycle()

    def run_cycle(self):
        now_h = int(time.strftime("%H"))
        try: s, e = int(self.ui_start.text()), int(self.ui_end.text())
        except: s, e = 9, 12
        if not (s <= now_h < e):
            self.add_log(f"💤 非活动时段 ({now_h}点)"); self.sync_timer.start(30 * 60000); return
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
            $it = $ol.GetNamespace("MAPI").GetDefaultFolder(6).Items | Where-Object {{ $_.ReceivedTime -gt (Get-Date).AddDays(-5) -and ($_.Subject -like "*{k}*") }} | Sort-Object ReceivedTime -Descending | Select-Object -First {c_num}
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
        for f in [x for x in os.listdir(d) if x.endswith('.mht')]:
            p_m, p_h = os.path.join(d, f), os.path.join(d, f.replace('.mht', '.html'))
            try:
                with open(p_m, 'rb') as fp:
                    msg = email.message_from_binary_file(fp)
                    for part in msg.walk():
                        if part.get_content_type() == "text/html":
                            raw_content = part.get_payload(decode=True).decode('utf-8','ignore')
                            clean_content = re.sub(r'width[:=]["\']?\d+(px|pt|in|cm)?["\']?', '', raw_content, flags=re.I)
                            with open(p_h, 'w', encoding='utf-8') as hw: hw.write(clean_content)
                os.remove(p_m)
            except: pass
        
        # --- 智能日历解析逻辑 ---
        cal_html = "<p style='color:red;padding:20px;'>未找到 2026日历.xlsx</p>"
        auto_zoom = "1.0"
        now_dt = datetime.datetime.now()
        today_str = now_dt.strftime('%Y-%m-%d')
        WORK_DATES_2026 = ["2026-01-04", "2026-02-14", "2026-02-28", "2026-05-09", "2026-09-20", "2026-10-10", "2026-10-11"]

        for f_name in os.listdir(d):
            if "2026日历" in f_name and f_name.lower().endswith('.xlsx'):
                try:
                    wb = openpyxl.load_workbook(os.path.join(d, f_name), data_only=True)
                    ws = wb.active
                    col_count = ws.max_column
                    if col_count > 30: auto_zoom = "0.55"
                    elif col_count > 20: auto_zoom = "0.75"
                    elif col_count > 12: auto_zoom = "0.85"
                    
                    rows_html = ""
                    for i, row in enumerate(ws.iter_rows(values_only=True)):
                        if not any(row): continue
                        row_content = ""
                        for cell_val in row:
                            val = "" if cell_val is None else str(cell_val)
                            style, extra_cls = "", ""
                            if i > 0 and cell_val:
                                try:
                                    dt = cell_val if isinstance(cell_val, datetime.datetime) else pd.to_datetime(val)
                                    dt_str = dt.strftime('%Y-%m-%d')
                                    if dt_str == today_str:
                                        extra_cls = "today-cell"
                                        style = "outline: 3px solid #ffba00; outline-offset: -3px; background-color: #fff9e6; font-weight: bold;"
                                    elif dt_str in WORK_DATES_2026:
                                        style = "background-color: #e6f2ff; color: #005a9e; font-weight: bold;"
                                    elif dt.weekday() >= 5:
                                        style = "background-color: #fff0f0; color: #e81123; font-weight: bold;"
                                    else:
                                        style = "background-color: #f9fff9; color: #333;"
                                    if (dt + datetime.timedelta(days=1)).month != dt.month:
                                        style += "border-bottom: 3px double #107c10;"
                                except:
                                    if "休" in val: style = "background-color: #fff4ce; color: #995d00;"
                                    if "班" in val: style = "background-color: #e6f2ff; color: #005a9e;"

                            tag = "th" if i == 0 else "td"
                            row_content += f"<{tag} class='{extra_cls}' style='{style}'>{val}</{tag}>"
                        rows_html += f"<tr>{row_content}</tr>"
                    cal_html = f"<div id='zoomWrap' style='zoom:{auto_zoom};'><table class='cal-table'>{rows_html}</table></div>"
                    self.add_log(f"📅 工作日历染色完成: {f_name}")
                    break
                except Exception as e: self.add_log(f"日历解析失败: {e}")
        self.build_index(cal_html)

    def build_index(self, cal_html):
        d, c = self.ui_path.text().strip(), self.ui_color.text().strip() or "#107c10"
        t1, t2, cp = self.ui_title.text().strip(), self.ui_subtitle.text().strip(), self.ui_copy.text().strip()
        all_files = [f for f in os.listdir(d) if f.endswith('.html') and f != 'index.html']
        all_files.sort(key=lambda x: os.path.getmtime(os.path.join(d, x)), reverse=True)
        try: w_ref = int(self.ui_web_freq.text())
        except: w_ref = 60
        update_time = time.strftime('%Y-%m-%d %H:%M:%S')
        items_html, mails_content_html, regex_ptr = "", "", self.ui_regex.text().strip()

        for i, f in enumerate(all_files):
            file_path = os.path.join(d, f)
            mtime = time.strftime('%m-%d %H:%M', time.localtime(os.path.getmtime(file_path)))
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as tf: content = tf.read()
            tags = " ".join(list(set(re.findall(regex_ptr, content))))
            items_html += f'<div class="mail-item {"active" if i==0 else ""}" onclick="showMail(\'{i}\', this)" data-tags="{tags}"><b>{f[:-5]}</b><br><small>🕒 {mtime}</small></div>'
            mails_content_html += f'<div id="mail-{i}" class="mail-body" style="display:{"block" if i==0 else "none"}"><div class="mail-inner-zoom">{content}</div></div>'

        full_html = f"""
        <!DOCTYPE html>
        <html><head><meta charset="utf-8"><meta http-equiv="refresh" content="{w_ref}"><title>{t1}</title>
        <style>
            body {{ font-family: 'Segoe UI', 'Microsoft YaHei', sans-serif; margin: 0; display: flex; height: 100vh; overflow: hidden; background:#f3f2f1; }}
            .sidebar {{ width: 340px; background: white; border-right: 1px solid #edebe9; display: flex; flex-direction: column; flex-shrink: 0; height: 100vh; z-index: 10; }}
            .header {{ padding: 18px 16px; background: white; flex-shrink: 0; border-bottom: 1px solid #f3f2f1; }}
            .search-box {{ padding: 12px 16px; background: #fff; border-bottom: 1px solid #f3f2f1; position: sticky; top: 0; flex-shrink: 0; z-index: 100; }}
            .search-box input {{ width: 100%; padding: 8px 35px 8px 10px; border: 1px solid #ddd; border-radius: 4px; outline: none; box-sizing: border-box; }}
            .clear-btn {{ position: absolute; right: 26px; top: 50%; transform: translateY(-50%); cursor: pointer; color: #bbb; display: none; font-size: 20px; }}
            .mail-list {{ flex: 1; overflow-y: auto; }}
            .mail-item {{ padding: 14px 16px; border-bottom: 1px solid #f3f2f1; cursor: pointer; transition: 0.1s; }}
            .mail-item.search-hit {{ background-color: #fff9c4 !important; border-left: 5px solid #fbc02d !important; font-weight: bold; }}
            .mail-item.active {{ border-left: 5px solid {c}; background: #eff6ef; }}
            .content {{ flex: 1; display: flex; flex-direction: column; min-width: 0; background: white; }}
            .mail-display {{ flex: 1; overflow: auto; }}
            .mail-inner-zoom {{ padding: 20px; zoom: 0.9; min-width: fit-content; }}
            .mail-inner-zoom table {{ table-layout: auto !important; width: auto !important; border-collapse: collapse; min-width: 600px; }}
            .mail-inner-zoom td, .mail-inner-zoom th {{ white-space: nowrap !important; padding: 5px 10px !important; font-size: 13px !important; border: 1px solid #ddd; }}
            .footer {{ font-size: 11px; color: #888; padding: 10px 16px; background: #fdfdfd; border-top: 1px solid #f3f2f1; display: flex; justify-content: space-between; align-items: center; }}
            .cal-trigger {{ cursor: pointer; color: {c}; font-weight: bold; text-decoration: underline; }}
            .modal {{ display: none; position: fixed; z-index: 9999; left: 0; top: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.6); backdrop-filter: blur(5px); }}
            .modal-content {{ background: white; margin: 1vh auto; width: 98%; height: 96%; border-radius: 8px; display: flex; flex-direction: column; overflow: hidden; }}
            .modal-body {{ flex: 1; overflow: auto; padding: 10px; display: flex; justify-content: center; }}
            .cal-table {{ border-collapse: collapse; font-size: 13px; width: 100%; }}
            .cal-table th, .cal-table td {{ border: 1px solid #ddd; padding: 8px 12px; text-align: center; white-space: nowrap; }}
            .today-cell {{ box-shadow: 0 0 10px rgba(255, 186, 0, 0.8); transform: scale(1.01); }}
        </style></head>
        <body>
            <div class="sidebar">
                <div class="header">
                    <div style="font-size:19px; font-weight:700; color:{c};">{t1}</div>
                    <div style="font-size:12px; color:#666; margin-top:3px;">{t2}</div>
                </div>
                <div class="search-box">
                    <input type="text" id="s" placeholder="搜索任务令/日期..." onkeyup="flt()">
                    <span id="cb" class="clear-btn" onclick="cls()">×</span>
                </div>
                <div class="mail-list" id="ml">{items_html}</div>
                <div class="footer">
                    <div><span>{cp}</span><br><small style="color:#aaa;">更新: {update_time}</small></div>
                    <span class="cal-trigger" onclick="tgl(true)">📅 工作日历</span>
                </div>
            </div>
            <div class="content"><div class="mail-display" id="mailDisplay">{mails_content_html}</div></div>
            <div id="mdl" class="modal">
                <div class="modal-content">
                    <div style="padding:15px; border-bottom:1px solid #eee; display:flex; justify-content:space-between; align-items:center;">
                        <h3 style="margin:0; color:{c};">📅 工作日历 (全屏预览)</h3>
                        <span style="cursor:pointer; font-size:35px;" onclick="tgl(false)">&times;</span>
                    </div>
                    <div class="modal-body">{cal_html}</div>
                </div>
            </div>
            <script>
                function showMail(id, el) {{
                    document.querySelectorAll('.mail-body').forEach(b => b.style.display = 'none');
                    document.querySelectorAll('.mail-item').forEach(i => i.classList.remove('active'));
                    document.getElementById('mail-'+id).style.display = 'block';
                    el.classList.add('active');
                    document.getElementById('mailDisplay').scrollTop = 0;
                }}
                function flt() {{
                    let v = document.getElementById('s').value.toUpperCase();
                    document.getElementById('cb').style.display = v ? 'block' : 'none';
                    let items = document.querySelectorAll('.mail-item'); let first = null;
                    items.forEach(item => {{
                        item.classList.remove('search-hit');
                        let txt = (item.innerText + item.getAttribute('data-tags')).toUpperCase();
                        if (v && txt.indexOf(v) > -1) {{
                            item.style.display = "block"; item.classList.add('search-hit'); if(!first) first = item;
                        }} else if(!v) {{ item.style.display = "block"; }} else {{ item.style.display = "none"; }}
                    }});
                    if(first && v) first.scrollIntoView({{ behavior: 'smooth', block: 'center' }});
                }}
                function cls() {{ document.getElementById('s').value=''; flt(); document.getElementById('s').focus(); }}
                function tgl(s) {{ 
                    document.getElementById('mdl').style.display = s ? 'block' : 'none'; 
                    if(s) {{ setTimeout(() => {{ let t = document.querySelector('.today-cell'); if(t) t.scrollIntoView({{ behavior: 'smooth', block: 'center' }}); }}, 150); }}
                }}
                window.onclick = function(e) {{ if(e.target == document.getElementById('mdl')) tgl(false); }}
            </script>
        </body></html>"""
        with open(os.path.join(d, "index.html"), 'w', encoding='utf-8') as f: f.write(full_html)
        self.add_log(f"✅ 网页同步完成 ({update_time})")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = OutlookMHTMaster()
    window.show()
    sys.exit(app.exec_())
