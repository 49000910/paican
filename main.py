import sys, os, re, time, subprocess, tempfile, urllib.parse, base64, email
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QTextEdit, QSystemTrayIcon, QMenu, QAction)
from PyQt5.QtCore import QTimer, Qt
from PyQt5.QtGui import QStyle

class OutlookMHTMaster(QWidget):
    def __init__(self):
        super().__init__()
        # --- 默认参数 ---
        self.share_dir = r'\\10.1.93.32\DT_HU_RDteam_F\视频\Z\ZOUQIU\paican'
        self.target_kw = 'EDFA' 
        self.interval_min = 10     
        self.sync_count = 3       
        self.start_hour = 9       
        self.end_hour = 12        
        self.theme_color = "#107c10" 
        self.web_title = "EDFA 排产看板"
        self.web_sub_title = "自动抓取 zouqiu@hauwei.com"
        
        self.init_ui()
        self.init_tray()
        self.sync_timer = QTimer(self)
        self.sync_timer.timeout.connect(self.run_cycle)
        QTimer.singleShot(2000, self.run_cycle)

    def init_ui(self):
        self.setWindowTitle("EDFA 排产看板管理后台 V26.9")
        self.resize(500, 750)
        layout = QVBoxLayout()
        layout.setContentsMargins(15, 15, 15, 15)

        def quick_edit(label, val, attr):
            l = QHBoxLayout()
            lb = QLabel(label); lb.setFixedWidth(100); l.addWidget(lb)
            edit = QLineEdit(str(val)); setattr(self, attr, edit)
            l.addWidget(edit); layout.addLayout(l)

        quick_edit("📂 共享路径", self.share_dir, "ui_path")
        quick_edit("📧 邮件关键词", self.target_kw, "ui_kw")
        quick_edit("🚩 网页大标题", self.web_title, "ui_title")
        quick_edit("📝 网页小字备注", self.web_sub_title, "ui_subtitle")
        
        h1 = QHBoxLayout()
        h1.addWidget(QLabel("⏱ 频率(分)"))
        self.ui_freq = QLineEdit(str(self.interval_min)); h1.addWidget(self.ui_freq)
        h1.addWidget(QLabel("🔢 每次抓取数"))
        self.ui_count = QLineEdit(str(self.sync_count)); h1.addWidget(self.ui_count)
        layout.addLayout(h1)

        h2 = QHBoxLayout()
        h2.addWidget(QLabel("⏰ 开始时")); self.ui_start = QLineEdit(str(self.start_hour)); h2.addWidget(self.ui_start)
        h2.addWidget(QLabel("⏰ 结束时")); self.ui_end = QLineEdit(str(self.end_hour)); h2.addWidget(self.ui_end)
        layout.addLayout(h2)

        quick_edit("🎨 主题颜色", self.theme_color, "ui_color")
        self.btn_apply = QPushButton("🚀 立即部署并同步 (全屏看板模式)")
        self.btn_apply.setFixedHeight(50); self.btn_apply.clicked.connect(self.apply_settings)
        layout.addWidget(self.btn_apply)

        self.log_area = QTextEdit(); self.log_area.setReadOnly(True)
        layout.addWidget(self.log_area); self.setLayout(layout); self.restyle()

    def init_tray(self):
        self.tray = QSystemTrayIcon(self)
        self.tray.setIcon(self.style().standardIcon(QStyle.SP_ComputerIcon))
        tm = QMenu(); tm.addAction("显示主界面", self.showNormal); tm.addAction("退出程序", QApplication.instance().quit)
        self.tray.setContextMenu(tm); self.tray.show()

    def closeEvent(self, event):
        if self.tray.isVisible(): self.hide(); event.ignore()

    def restyle(self):
        c = self.ui_color.text().strip() or "#107c10"
        self.setStyleSheet(f"QPushButton{{background:{c};color:white;font-weight:bold;border-radius:4px;}} QTextEdit{{background:#1e1e1e;color:#0f0;border:1px solid {c};font-family:Consolas;}}")

    def add_log(self, txt):
        self.log_area.append(f"[{time.strftime('%H:%M:%S')}] {str(txt)}")

    def apply_settings(self):
        self.restyle(); self.add_log("⚙️ 配置下发完成..."); self.run_cycle()

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
            if (!(Test-Path "{d}")) {{ New-Item -ItemType Directory -Path "{d}" -Force | Out-Null }}
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
                            with open(p_h, 'w', encoding='utf-8') as hw: hw.write(part.get_payload(decode=True).decode('utf-8','ignore'))
                            break
                os.remove(p_m); self.add_log(f"✅ 抓取: {f}")
            except: pass
        self.build_index()

    def build_index(self):
        d = self.ui_path.text().strip()
        all_files = [f for f in os.listdir(d) if f.endswith('.html') and f not in ['index.html', 'list_inner.html']]
        all_files.sort(key=lambda x: os.path.getmtime(os.path.join(d, x)), reverse=True)
        
        c, t1, t2 = self.ui_color.text().strip(), self.ui_title.text().strip(), self.ui_subtitle.text().strip()
        try: r_sec = int(self.ui_freq.text()) * 60
        except: r_sec = 600

        items_html = ""
        for f in all_files[:100]:
            p = os.path.join(d, f)
            try:
                raw_h = open(p, 'r', encoding='utf-8', errors='ignore').read()
                # --- 核心：提取 EP 开头 11 位编码 ---
                eps = list(set(re.findall(r'\bEP[A-Z0-9]{9}\b', raw_h, re.I)))
                tags = "".join([f'<span class="et" style="background:{c}">{x}</span>' for x in eps[:2]])
                txt = re.sub(r'<[^>]+>', '', raw_h).replace('\\n',' ').lower()
                # 索引包含：EP号 + 文件名 + 正文文本
                sk = (" ".join(eps) + " " + f + " " + txt).replace("'", "").replace('"', '')
            except: tags, sk = "", f.lower()
            
            items_html += f'''<div class="item" onclick="selectItem(this, '{urllib.parse.quote(f)}')" data-s="{sk[:5000]}">
                <div class="ti">{f[:28]}...</div><div class="tr">{tags}</div>
                <div class="tm">{time.strftime("%H:%M", time.localtime(os.path.getmtime(p)))}</div></div>'''

        list_html = f"""<html><head><meta charset='utf-8'><meta http-equiv="refresh" content="{r_sec}"><style>
            body {{ margin:0; padding:0; font-family:sans-serif; overflow-x:hidden; background:#fff; }}
            .sb {{ position:sticky; top:0; background:#fff; padding:10px; border-bottom:1px solid #eee; z-index:9; }}
            .sb input {{ width:100%; padding:8px; border:1px solid #ddd; border-radius:4px; font-size:12px; outline:none; box-sizing:border-box; }}
            .item {{ padding:12px 15px; border-bottom:1px solid #f2f2f2; cursor:pointer; border-left:4px solid transparent; transition:0.1s; }}
            .active {{ background:#f0f7f0 !important; border-left-color:{c} !important; font-weight:bold; }}
            .ti {{ font-size:13px; color:#333; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }}
            .et {{ color:white; padding:1px 4px; border-radius:3px; font-size:10px; margin-right:4px; font-family:Consolas; }}
            .tm {{ font-size:11px; color:#999; margin-top:4px; }}
            ::-webkit-scrollbar {{ width:4px; }} ::-webkit-scrollbar-thumb {{ background:#ddd; }}
        </style><script>
            function doSearch(k) {{ var its=document.querySelectorAll('.item'); k=k.toLowerCase(); its.forEach(i=>{{ i.style.display=i.getAttribute('data-s').includes(k)?'block':'none'; }}); }}
            function selectItem(el, u) {{ document.querySelectorAll('.item').forEach(i=>i.classList.remove('active')); el.classList.add('active'); parent.loadMail(u); }}
        </script></head><body><div class="sb"><input placeholder="搜 EP 号或正文..." oninput="doSearch(this.value)"></div>{items_html}</body></html>"""

        main_ui = f"""<!DOCTYPE html><html><head><meta charset='utf-8'><title>{t1}</title><style>
            * {{ margin:0!important; padding:0!important; box-sizing:border-box!important; }}
            html, body {{ width:100%; height:100%; overflow:hidden; background:#fff; }}
            .layout {{ display:flex; width:100vw; height:100vh; }}
            .side {{ width:280px; height:100%; display:flex; flex-direction:column; border-right:1px solid #eee; flex-shrink:0; }}
            .hd {{ background:{c}; color:white; padding:18px 15px; }}
            .brand {{ font-size:18px; font-weight:bold; }}
            .sub-brand {{ font-size:11px; opacity:0.8; margin-top:5px; }}
            iframe {{ border:none!important; width:100%; height:100%; display:block; }}
        </style><script>function loadMail(u) {{ document.getElementById('vf').src = u; }}</script></head>
        <body><div class="layout">
            <div class="side"><div class="hd"><span class="brand">{t1}</span><span class="sub-brand">{t2}</span><div style="font-size:10px; opacity:0.6; margin-top:10px; padding-top:5px; border-top:1px solid rgba(255,255,255,0.2);">● 活跃 (9-12点) | {time.strftime('%H:%M:%S')}</div></div>
            <iframe src="list_inner.html"></iframe></div>
            <div style="flex:1;"><iframe id="vf" src="about:blank"></iframe></div>
        </div></body></html>"""
        
        with open(os.path.join(d, 'list_inner.html'), 'w', encoding='utf-8') as f1: f1.write(list_html)
        with open(os.path.join(d, 'index.html'), 'w', encoding='utf-8') as f2: f2.write(main_ui)

if __name__ == "__main__":
    app = QApplication(sys.argv); w = OutlookMHTMaster(); w.show(); sys.exit(app.exec_())
