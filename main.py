import sys, os, re, time, subprocess, tempfile, urllib.parse, base64, email, json
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
        self.setWindowTitle("EDFA 看板管理后台 V27.5")
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
        quick_edit("🔒 版权内容", self.copyright_text, "ui_copy")

        self.btn_apply = QPushButton("🚀 立即部署并全屏同步")
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
        self.share_dir, self.target_kw = self.ui_path.text().strip(), self.ui_kw.text().strip()
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
                            # 注入强力背景颜色和基础字体
                            clean_h = f"<style>body{{font-family:sans-serif;padding:15px;font-size:14px;background:#fff !important;}}</style>" + part.get_payload(decode=True).decode('utf-8','ignore')
                            with open(p_h, 'w', encoding='utf-8') as hw: hw.write(clean_h)
                            break
                os.remove(p_m); self.add_log(f"✅ 抓取: {f}")
            except: pass
        self.build_index()

    def build_index(self):
        d = self.ui_path.text().strip()
        all_files = [f for f in os.listdir(d) if f.endswith('.html') and f != 'index.html']
        all_files.sort(key=lambda x: os.path.getmtime(os.path.join(d, x)), reverse=True)
        
        c, t1, t2, cp = self.ui_color.text().strip(), self.ui_title.text().strip(), self.ui_subtitle.text().strip(), self.ui_copy.text().strip()
        try: r_sec = int(self.ui_freq.text()) * 60
        except: r_sec = 600
        last_sync = time.strftime("%H:%M:%S")

        items_html = ""
        mails_data_html = ""
        for i, f in enumerate(all_files[:200]): # 扩充至200封索引
            p = os.path.join(d, f)
            try:
                with open(p, 'r', encoding='utf-8', errors='ignore') as fc: raw_h = fc.read()
                tags_list = list(set(re.findall(r'\bEP[A-Z0-9]{9}\b', raw_h, re.I)))
                tags_ui = "".join([f'<span class="et" style="background:{c}">{x}</span>' for x in tags_list[:2]])
                
                # --- 暴力索引逻辑：正文脱水 + JSON 安全处理 ---
                txt_pure = re.sub(r'<[^>]+>', ' ', raw_h) # 去标签
                txt_dehydrated = re.sub(r'\s+', '', txt_pure).lower() # 全脱水
                search_database = f"{f} {txt_pure} {txt_dehydrated}".lower()
                safe_sk = json.dumps(search_database).strip('"') # 解决引号崩溃问题

                items_html += f'''<div class="item" onclick="showMail({i}, this)" data-s="{safe_sk}">
                    <div class="ti">{f[:28]}...</div><div class="tr">{tags_ui}</div>
                    <div class="tm">{time.strftime("%m-%d %H:%M", time.localtime(os.path.getmtime(p)))}</div></div>'''
                
                safe_content = base64.b64encode(raw_h.encode('utf-8')).decode('utf-8')
                mails_data_html += f'<div id="md_{i}" style="display:none;">{safe_content}</div>'
            except: continue

        main_ui = f"""<!DOCTYPE html><html><head><meta charset='utf-8'>
        <title>{t1}</title><style>
            * {{ margin:0!important; padding:0!important; box-sizing:border-box!important; }}
            html, body {{ width:100%; height:100%; overflow:hidden; font-family:sans-serif; background:#fff; }}
            .layout {{ display:flex; width:100vw; height:100vh; }}
            .side {{ width:300px; height:100%; display:flex; flex-direction:column; border-right:1px solid #eee; flex-shrink:0; }}
            .hd {{ background:{c}; color:white; padding:15px; border:none; }}
            .brand {{ font-size:18px; font-weight:bold; display:block; }}
            .sub-brand {{ font-size:11px; opacity:0.8; margin-top:5px; display:block; }}
            .st {{ font-size:10px; opacity:0.6; margin-top:10px; padding-top:5px; border-top:1px solid rgba(255,255,255,0.2); }}
            .sb {{ padding:10px; border-bottom:1px solid #eee; background:#fcfcfc; }}
            #kw {{ width:100%; padding:10px; border:1px solid #ddd; border-radius:4px; font-size:13px; outline:none; }}
            #kw:focus {{ border-color:{c}; }}
            .list {{ flex:1; overflow-y:auto; background:#fff; }}
            .item {{ padding:12px 15px; border-bottom:1px solid #f2f2f2; cursor:pointer; border-left:4px solid transparent; transition:0.1s; }}
            .active {{ background:#f0f7f0 !important; border-left-color:{c} !important; font-weight:bold; }}
            .ti {{ font-size:13px; color:#333; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }}
            .et {{ color:white; padding:2px 5px; border-radius:3px; font-size:10px; margin-right:4px; font-family:Consolas; }}
            .tm {{ font-size:11px; color:#999; margin-top:4px; }}
            #vf {{ border:none; width:100%; height:100%; flex:1; display:block; background:#fff; }}
            ::-webkit-scrollbar {{ width:4px; }} ::-webkit-scrollbar-thumb {{ background:#ddd; }}
        </style>
        <script>
            function doSearch() {{
                var k = document.getElementById('kw').value.replace(/\s+/g, '').toLowerCase().trim();
                localStorage.setItem('last_kw', document.getElementById('kw').value); 
                var its = document.querySelectorAll('.item');
                its.forEach(i => {{
                    i.style.display = i.getAttribute('data-s').indexOf(k) !== -1 ? 'block' : 'none';
                }});
            }}
            function showMail(idx, el) {{
                document.querySelectorAll('.item').forEach(i => i.classList.remove('active'));
                el.classList.add('active');
                localStorage.setItem('last_idx', idx);
                var b64 = document.getElementById('md_' + idx).innerText;
                var html = decodeURIComponent(escape(window.atob(b64)));
                document.getElementById('vf').srcdoc = html;
            }}
            window.onload = function() {{
                var savedKw = localStorage.getItem('last_kw');
                if(savedKw) {{ document.getElementById('kw').value = savedKw; doSearch(); }}
                document.getElementById('kw').focus();
                // 使用 JS 控制定时刷新，保持搜索焦点
                setTimeout(() => {{ window.location.reload(); }}, {r_sec * 1000});
            }}
        </script></head>
        <body><div class="layout">
            <div class="side">
                <div class="hd"><span class="brand">{t1}</span><span class="sub-brand">{t2}</span><div class="st">同步时间: {last_sync}<br>{cp}</div></div>
                <div class="sb"><input id="kw" placeholder="输入任务令(自动去空格)..." oninput="doSearch()"></div>
                <div class="list">{items_html}</div>
            </div>
            <iframe id="vf" srcdoc='<body style="display:flex;justify-content:center;align-items:center;height:100vh;color:#999;font-family:sans-serif;">点击左侧查看排产详情</body>'></iframe>
        </div><div id="store" style="display:none;">{mails_data_html}</div></body></html>"""
        
        with open(os.path.join(d, 'index.html'), 'w', encoding='utf-8') as f: f.write(main_ui)

if __name__ == "__main__":
    app = QApplication(sys.argv); w = OutlookMHTMaster(); w.show(); sys.exit(app.exec_())
