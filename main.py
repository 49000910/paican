import sys, os, re, time, subprocess, tempfile, urllib.parse, base64, email
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QTextEdit, QSystemTrayIcon, QMenu, QAction)
from PyQt5.QtCore import QTimer, Qt
from PyQt5.QtGui import QStyle

class OutlookMHTMaster(QWidget):
    def __init__(self):
        super().__init__()
        # --- 默认参数设置 (对应您的最新需求) ---
        self.share_dir = r'\\10.1.93.32\DT_HU_RDteam_F\视频\Z\ZOUQIU\paican'
        self.target_kw = 'EDFA' 
        self.interval_min = 10     
        self.sync_count = 3       
        self.start_hour = 9       
        self.end_hour = 12        
        self.theme_color = "#107c10" 
        self.copyright_text = "© 2024-2026 R1231685 | 技术支持"
        
        self.init_ui()
        self.init_tray()
        
        self.sync_timer = QTimer(self)
        self.sync_timer.timeout.connect(self.run_cycle)
        QTimer.singleShot(2000, self.run_cycle) # 启动2秒后执行首次同步

    def init_ui(self):
        self.setWindowTitle("EDFA 排产看板同步工具 V26.9")
        self.resize(480, 680)
        layout = QVBoxLayout()
        layout.setContentsMargins(15, 15, 15, 15)

        def quick_edit(label, val, attr):
            l = QHBoxLayout()
            lb = QLabel(label); lb.setFixedWidth(100); l.addWidget(lb)
            edit = QLineEdit(str(val)); setattr(self, attr, edit)
            l.addWidget(edit); layout.addLayout(l)

        quick_edit("📂 共享路径", self.share_dir, "ui_path")
        quick_edit("📧 关键词", self.target_kw, "ui_kw")
        
        h1 = QHBoxLayout()
        h1.addWidget(QLabel("⏱ 频率(分)"))
        self.ui_freq = QLineEdit(str(self.interval_min)); h1.addWidget(self.ui_freq)
        h1.addWidget(QLabel("🔢 数量"))
        self.ui_count = QLineEdit(str(self.sync_count)); h1.addWidget(self.ui_count)
        layout.addLayout(h1)

        h2 = QHBoxLayout()
        h2.addWidget(QLabel("⏰ 开始时")); self.ui_start = QLineEdit(str(self.start_hour)); h2.addWidget(self.ui_start)
        h2.addWidget(QLabel("⏰ 结束时")); self.ui_end = QLineEdit(str(self.end_hour)); h2.addWidget(self.ui_end)
        layout.addLayout(h2)

        quick_edit("🎨 主题颜色", self.theme_color, "ui_color")
        quick_edit("📝 版权内容", self.copyright_text, "ui_copy")

        self.btn_apply = QPushButton("🚀 立即部署并同步 (看板全屏化)")
        self.btn_apply.setFixedHeight(45); self.btn_apply.clicked.connect(self.apply_settings)
        layout.addWidget(self.btn_apply)

        self.log_area = QTextEdit(); self.log_area.setReadOnly(True)
        layout.addWidget(self.log_area); self.setLayout(layout); self.restyle()

    def init_tray(self):
        self.tray = QSystemTrayIcon(self)
        self.tray.setIcon(self.style().standardIcon(QStyle.SP_ComputerIcon))
        tray_menu = QMenu()
        show_action = QAction("显示主界面", self); show_action.triggered.connect(self.showNormal)
        quit_action = QAction("退出程序", self); quit_action.triggered.connect(QApplication.instance().quit)
        tray_menu.addAction(show_action); tray_menu.addSeparator(); tray_menu.addAction(quit_action)
        self.tray.setContextMenu(tray_menu); self.tray.activated.connect(lambda r: self.showNormal() if r==QSystemTrayIcon.Trigger else None); self.tray.show()

    def closeEvent(self, event):
        if self.tray.isVisible():
            self.hide()
            self.tray.showMessage("同步工具", "已缩回托盘后台运行", QSystemTrayIcon.Information, 2000)
            event.ignore()

    def restyle(self):
        c = self.ui_color.text().strip()
        if not c.startswith("#"): c = "#107c10"
        self.setStyleSheet(f"QPushButton{{background:{c};color:white;font-weight:bold;border-radius:4px;}} QTextEdit{{background:#1e1e1e;color:#0f0;border:1px solid {c};font-family:Consolas;}}")

    def add_log(self, txt):
        t_str = time.strftime('%H:%M:%S')
        self.log_area.append(f"[{t_str}] {str(txt)}")

    def apply_settings(self):
        self.share_dir, self.target_kw = self.ui_path.text().strip(), self.ui_kw.text().strip()
        self.restyle(); self.add_log("⚙️ 参数已更新并重写索引..."); self.run_cycle()

    def run_cycle(self):
        now_h = int(time.strftime("%H"))
        try: s_h, e_h = int(self.ui_start.text()), int(self.ui_end.text())
        except: s_h, e_h = 9, 12
        if not (s_h <= now_h < e_h):
            self.add_log(f"💤 休眠 (时段外:{now_h}点)"); self.sync_timer.start(30 * 60000); return
        self.run_shell() 
        try: freq = int(self.ui_freq.text()); self.sync_timer.start(freq * 60000)
        except: self.sync_timer.start(600000)

    def run_shell(self):
        ps_dir, ps_kw = self.share_dir.replace('"', '""'), self.target_kw.replace('"', '""')
        try: count = int(self.ui_count.text())
        except: count = 3
        ps_cmd = f"""
        try {{
            if (!(Test-Path "{ps_dir}")) {{ New-Item -ItemType Directory -Path "{ps_dir}" -Force | Out-Null }}
            $ol = New-Object -ComObject Outlook.Application
            $it = $ol.GetNamespace("MAPI").GetDefaultFolder(6).Items | Where-Object {{ $_.ReceivedTime -gt (Get-Date).AddDays(-3) -and ($_.Subject -like "*{ps_kw}*" -or $_.SenderName -like "*{ps_kw}*") }} | Sort-Object ReceivedTime -Descending | Select-Object -First {count}
            if ($it) {{ foreach($m in $it) {{ $n = ($m.Subject -replace '[\\x00-\\x1f\\\\/:*?"<>|]', '_').Trim(); $p = Join-Path "{ps_dir}" "$n.mht"; if (!(Test-Path $p)) {{ $m.SaveAs($p, 10) }} }} }}
        }} catch {{ }} finally {{ if ($ol) {{ [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ol) | Out-Null }} }}
        """
        try:
            ps_b64 = base64.b64encode(ps_cmd.encode('utf-16-le')).decode('ascii')
            subprocess.run(["powershell", "-NoProfile", "-EncodedCommand", ps_b64], creationflags=0x08000000, timeout=60)
            self.process_web()
        except Exception as e: self.add_log(f"ERR: {e}")

    def process_web(self):
        if not os.path.exists(self.share_dir): return
        mhts = [f for f in os.listdir(self.share_dir) if f.endswith('.mht')]
        for f in mhts:
            p_m, p_h = os.path.join(self.share_dir, f), os.path.join(self.share_dir, f.replace('.mht', '.html'))
            try:
                with open(p_m, 'rb') as fp:
                    msg = email.message_from_binary_file(fp)
                    for part in msg.walk():
                        if part.get_content_type() == "text/html":
                            payload = part.get_payload(decode=True)
                            with open(p_h, 'w', encoding='utf-8') as hw: hw.write(payload.decode('utf-8','ignore'))
                            break
                os.remove(p_m); self.add_log(f"✅ 抓取成功: {f}")
                self.tray.showMessage("看板已更新", f"新邮件: {f[:15]}...", QSystemTrayIcon.Information, 3000)
            except: pass
        self.build_index()

    def build_index(self):
        all_files = [f for f in os.listdir(self.share_dir) if f.endswith('.html') and f not in ['index.html', 'list_inner.html']]
        all_files.sort(key=lambda x: os.path.getmtime(os.path.join(self.share_dir, x)), reverse=True)
        c = self.ui_color.text().strip()
        if not c.startswith("#"): c = "#107c10"
        try: r_sec = int(self.ui_freq.text()) * 60
        except: r_sec = 600
        now_ts, last_sync = time.time(), time.strftime("%H:%M:%S")

        items_html = ""
        for f in all_files[:150]: # 索引最近150封正文
            p = os.path.join(self.share_dir, f); mtime = os.path.getmtime(p)
            try:
                raw_h = open(p, 'r', encoding='utf-8', errors='ignore').read()
                ep = list(set(re.findall(r'\bEP[A-Z0-9]{9}\b', raw_h, re.I)))
                tags = "".join([f'<span class="et" style="background:{c}">{x}</span>' for x in ep[:2]])
                txt = re.sub(r'<[^>]+>', '', raw_h).replace('\\n',' ').lower() # 提取纯文本索引
                sk = (f + txt).replace("'", "").replace('"', '')
            except: tags, sk = "", f.lower()
            
            nt = '<span class="nt">● NEW</span>' if (now_ts - mtime) < 1800 else ""
            items_html += f'''<div class="item" onclick="selectItem(this, '{urllib.parse.quote(f)}')" data-s="{sk[:5000]}">
                <div class="ti">{f[:28]}...{nt}</div><div class="tr">{tags}</div>
                <div class="tm">{time.strftime("%m-%d %H:%M", time.localtime(mtime))}</div></div>'''

        # 列表页生成 (list_inner.html)
        list_html = f"""<html><head><meta charset='utf-8'><meta http-equiv="refresh" content="{r_sec}">
        <style>
            body {{ margin:0; padding:0; font-family:'Segoe UI',sans-serif; background:#fff; overflow-x:hidden; }}
            .sb {{ position:sticky; top:0; background:#fff; padding:10px; border-bottom:1px solid #eee; z-index:9; }}
            .sb input {{ width:100%; padding:8px; border:1px solid #ddd; border-radius:4px; font-size:12px; outline:none; box-sizing:border-box; }}
            .sb input:focus {{ border-color:{c}; }}
            .item {{ padding:12px 15px; border-bottom:1px solid #f2f2f2; cursor:pointer; border-left:4px solid transparent; transition:0.2s; }}
            .item:hover {{ background:#f9f9f9; }}
            .active {{ background:#f0f7f0 !important; border-left-color:{c} !important; }}
            .active .ti {{ color:{c}; font-weight:bold; }}
            .ti {{ font-size:13px; color:#333; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }}
            .et {{ color:white; padding:1px 4px; border-radius:3px; font-size:10px; margin-right:4px; }}
            .nt {{ color:#ff4d4f; font-size:10px; margin-left:5px; animation:bk 1s infinite; }}
            .tm {{ font-size:11px; color:#999; margin-top:4px; }}
            @keyframes bk {{ 0%{{opacity:1;}} 50%{{opacity:0.3;}} 100%{{opacity:1;}} }}
            ::-webkit-scrollbar {{ width:4px; }} ::-webkit-scrollbar-thumb {{ background:#ddd; border-radius:2px; }}
        </style>
        <script>
            function doSearch(k) {{
                var its = document.getElementsByClassName('item');
                k = k.toLowerCase();
                for(var i=0; i<its.length; i++) its[i].style.display = its[i].getAttribute('data-s').includes(k) ? 'block' : 'none';
            }}
            function selectItem(el, url) {{
                var its = document.getElementsByClassName('item');
                for(var i=0; i<its.length; i++) its[i].classList.remove('active');
                el.classList.add('active'); parent.loadMail(url);
            }}
        </script></head><body>
            <div class="sb"><input type="text" placeholder="全文索引搜索 (标题、正文、EP)..." oninput="doSearch(this.value)"></div>
            {items_html}
        </body></html>"""

        # 主控页生成 (index.html) - 全屏无白边布局
        main_ui = f"""<!DOCTYPE html><html><head><meta charset='utf-8'>
        <title>EDFA 排产看板</title>
        <style>
            body, html {{ margin:0; padding:0; height:100%; width:100%; overflow:hidden; font-family:'Segoe UI',sans-serif; background:#fff; }}
            .layout {{ display:flex; width:100%; height:100%; }}
            .side {{ width:280px; height:100%; display:flex; flex-direction:column; border-right:1px solid #eee; flex-shrink:0; }}
            .hd {{ background:{c}; color:white; padding:15px; border-bottom:1px solid rgba(0,0,0,0.1); }}
            .brand {{ font-size:16px; font-weight:bold; letter-spacing:1px; }}
            .sub-brand {{ font-size:11px; opacity:0.8; margin-top:4px; }}
            .st {{ font-size:10px; opacity:0.6; margin-top:8px; border-top:1px solid rgba(255,255,255,0.2); padding-top:5px; }}
            iframe {{ border:none; width:100%; height:100%; display:block; }}
        </style>
        <script>function loadMail(u) {{ document.getElementById('vf').src = u; }}</script>
        </head><body><div class="layout">
            <div class="side">
                <div class="hd">
                    <div class="brand">EDFA 排产看板</div>
                    <div class="sub-brand">自动抓取 zouqiu@hauwei.com</div>
                    <div class="st">● 活跃时段 (9-12点) | 更新: {last_sync}</div>
                </div>
                <iframe src="list_inner.html"></iframe>
            </div>
            <div style="flex:1;"><iframe id="vf" src="about:blank"></iframe></div>
        </div></body></html>"""
        
        with open(os.path.join(self.share_dir, 'list_inner.html'), 'w', encoding='utf-8') as f1: f1.write(list_html)
        with open(os.path.join(self.share_dir, 'index.html'), 'w', encoding='utf-8') as f2: f2.write(main_ui)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = OutlookMHTMaster()
    window.show()
    sys.exit(app.exec_())
