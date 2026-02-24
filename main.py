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
        # --- 默认参数 (100% 还原) ---
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
        self.setWindowTitle("EDFA 看板后台 V59.0")
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
        self.btn_apply = QPushButton("🚀 立即同步并解析"); self.btn_apply.setFixedHeight(50)
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
    def apply_settings(self): self.restyle(); self.add_log("⚙️ 重新执行同步..."); self.run_cycle()

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
        
        # 1. 邮件解析
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
            
        all_htmls = [x for x in os.listdir(d) if x.endswith('.html') and x != "index.html"]
        all_htmls.sort(key=lambda x: os.path.getmtime(os.path.join(d, x)), reverse=True)

        # 2. 日历解析 (保留原逻辑)
        cal_html = ""
        for f_name in os.listdir(d):
            if "2026日历" in f_name and f_name.lower().endswith('.xlsx'):
                try:
                    wb = openpyxl.load_workbook(os.path.join(d, f_name), data_only=True)
                    ws = wb.active
                    rows_html = "<table border='1' style='border-collapse:collapse;width:100%;font-size:12px;'>"
                    for row in ws.iter_rows():
                        rows_html += "<tr>"
                        for cell in row:
                            val = "" if cell.value is None else str(cell.value)
                            bg = f"#{cell.fill.start_color.rgb[2:]}" if cell.fill and hasattr(cell.fill.start_color, 'rgb') and len(str(cell.fill.start_color.rgb))>2 else "white"
                            rows_html += f"<td style='background:{bg};padding:5px;'>{val}</td>"
                        rows_html += "</tr>"
                    cal_html = rows_html + "</table>"
                except: pass

        # 3. 构造邮件项
        mail_items_html = ""
        for h in all_htmls:
            tags = re.findall(self.ui_regex.text(), h)
            tag_spans = "".join([f"<span style='font-size:10px;background:{self.ui_color.text()};color:white;padding:2px 5px;border-radius:3px;margin-right:5px;'>{t}</span>" for t in tags])
            mail_items_html += f"""
            <div style='background:white;padding:12px;margin-bottom:10px;border-left:5px solid {self.ui_color.text()};cursor:pointer;box-shadow:0 2px 4px rgba(0,0,0,0.05);' onclick="viewMail('{h}')">
                <div style='font-weight:bold;margin-bottom:5px;'>{h.replace('.html','')}</div>
                <div>{tag_spans}</div>
            </div>"""

        # 4. 生成 index.html
        index_path = os.path.join(d, "index.html")
        with open(index_path, 'w', encoding='utf-8') as f:
            f.write(f"""
            <html><head><meta charset="utf-8">
            <title>{self.ui_title.text()}</title>
            <style>
                body {{ display: flex; margin: 0; height: 100vh; font-family: 'Microsoft YaHei'; overflow: hidden; background:#eee; }}
                #sidebar {{ width: 380px; background: #f8f9fa; border-right: 1px solid #ddd; display: flex; flex-direction: column; box-shadow: 2px 0 5px rgba(0,0,0,0.1); }}
                #list_container {{ flex: 1; overflow-y: auto; padding: 15px; }}
                #content_area {{ flex: 1; border: none; background: white; }}
                h3 {{ font-size: 14px; color: #555; border-bottom: 1px solid #eee; padding-bottom: 5px; margin-top:15px; }}
                .btn-cal {{ background:{self.ui_color.text()}; color:white; border:none; padding:10px; width:100%; border-radius:4px; cursor:pointer; font-weight:bold; margin-bottom:10px; }}
                
                /* 底部倒计时与绿点样式 */
                #footer {{ padding: 15px; background: white; border-top: 1px solid #eee; text-align: center; }}
                .status-container {{ display: flex; align-items: center; justify-content: center; gap: 8px; font-size: 12px; color: #666; margin-bottom: 5px; }}
                .dot {{ width: 8px; height: 8px; background-color: #28a745; border-radius: 50%; display: inline-block; animation: blink 1.5s infinite; }}
                @keyframes blink {{ 0% {{ opacity: 1; }} 50% {{ opacity: 0.3; }} 100% {{ opacity: 1; }} }}
                
                /* 弹窗 */
                #cal_modal {{ display:none; position:fixed; top:5%; left:5%; width:90%; height:85%; background:white; z-index:1000; box-shadow:0 0 20px rgba(0,0,0,0.5); border-radius:8px; overflow:auto; padding:20px; }}
                #modal_mask {{ display:none; position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.6); z-index:999; }}
                .close-btn {{ float:right; cursor:pointer; color:red; font-weight:bold; font-size:18px; }}
            </style>
            </head><body>
            <div id="sidebar">
                <div style="padding:20px; background:white; border-bottom:1px solid #eee;">
                    <h2 style="margin:0; font-size:20px; color:{self.ui_color.text()};">{self.ui_title.text()}</h2>
                    <div style="font-size:11px; color:#999; margin-top:4px;">{self.web_sub_title}</div>
                </div>
                <div id="list_container">
                    <button class="btn-cal" onclick="showCal()">📅 查看生产排产日历 (Excel)</button>
                    <h3>📧 最新邮件列表</h3>
                    {mail_items_html}
                </div>
                
                <!-- 底部计时布局还原 -->
                <div id="footer">
                    <div class="status-container">
                        <span class="dot"></span>
                        <span id="timer_ui">正在等待同步...</span>
                    </div>
                    <div style="font-size:10px; color:#bbb;">{self.copyright_text} | 最后同步: {time.strftime('%H:%M:%S')}</div>
                </div>
            </div>
            
            <iframe id="content_area" name="content_area" src="about:blank"></iframe>

            <div id="modal_mask" onclick="hideCal()"></div>
            <div id="cal_modal">
                <span class="close-btn" onclick="hideCal()">[关闭窗口]</span>
                <h2 style="color:{self.ui_color.text()}">生产排产日历预览</h2>
                <hr>{cal_html}
            </div>

            <script>
                function showCal() {{ document.getElementById('cal_modal').style.display='block'; document.getElementById('modal_mask').style.display='block'; }}
                function hideCal() {{ document.getElementById('cal_modal').style.display='none'; document.getElementById('modal_mask').style.display='none'; }}
                function viewMail(url) {{ document.getElementById('content_area').src = url; localStorage.setItem('current_mail', url); }}

                let countdown = {self.ui_web_freq.text()};
                setInterval(() => {{
                    countdown--;
                    document.getElementById('timer_ui').innerText = '数据活跃中: 预计 ' + countdown + ' 秒后刷新';
                    if(countdown <= 0) {{ location.reload(); }}
                }}, 1000);

                window.onload = () => {{
                    let lastUrl = localStorage.getItem('current_mail');
                    if(lastUrl && lastUrl !== 'about:blank') {{ document.getElementById('content_area').src = lastUrl; }}
                }};
            </script>
            </body></html>
            """)
        self.add_log("✅ 看板 V59.0 布局还原完成（含底部倒计时与绿点闪烁）")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = OutlookMHTMaster()
    window.show()
    sys.exit(app.exec_())
