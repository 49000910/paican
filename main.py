import sys, os, re, time, subprocess, tempfile, urllib.parse, base64, email
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QTextEdit, QSystemTrayIcon, QMenu)
from PyQt5.QtCore import QTimer, Qt

class OutlookMHTMaster(QWidget):
    def __init__(self):
        super().__init__()
        # --- 默认配置 ---
        self.share_dir = r'\\10.1.93.32\DT_HU_RDteam_F\视频\Z\ZOUQIU\paican'
        self.target_kw = 'EDFA' 
        self.interval_min = 10     
        self.sync_count = 1       
        self.start_hour = 8       
        self.end_hour = 20        
        self.theme_color = "#107c10" 
        self.copyright_text = "© 2024-2026 RD Team | 视频组技术支持"
        
        self.tmp_log = os.path.join(tempfile.gettempdir(), "sync_v26_res.txt")
        self.init_ui()
        self.init_tray()
        
        self.sync_timer = QTimer(self)
        self.sync_timer.timeout.connect(self.run_cycle)
        QTimer.singleShot(2000, self.run_cycle)

    def init_ui(self):
        self.setWindowTitle("RD Sync Tool V26.0")
        self.resize(480, 680)
        layout = QVBoxLayout()
        layout.setContentsMargins(15, 15, 15, 15)

        # 极简后台配置项
        def quick_edit(label, val, attr):
            l = QHBoxLayout()
            l.addWidget(QLabel(label))
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

        self.btn_apply = QPushButton("🚀 部署并同步 (网页已加框)")
        self.btn_apply.setFixedHeight(45); self.btn_apply.clicked.connect(self.apply_settings)
        layout.addWidget(self.btn_apply)

        self.log = QTextEdit(); self.log.setReadOnly(True)
        layout.addWidget(self.log); self.setLayout(layout); self.restyle()

    def restyle(self):
        c = self.ui_color.text().strip()
        if not c.startswith("#"): c = "#107c10"
        self.setStyleSheet(f"QPushButton{{background:{c};color:white;font-weight:bold;border-radius:4px;}} QTextEdit{{background:#1e1e1e;color:#0f0;border:1px solid {c};font-family:Consolas;}}")

    def add_log(self, txt):
        self.log.append(f"[{time.strftime('%H:%M:%S')}] {str(txt).replace('\x00','')}")

    def apply_settings(self):
        self.share_dir, self.target_kw = self.ui_path.text().strip(), self.ui_kw.text().strip()
        self.restyle(); self.add_log("⚙️ 配置生效..."); self.run_cycle()

    def run_cycle(self):
        now_h = int(time.strftime("%H"))
        if not (int(self.ui_start.text()) <= now_h < int(self.ui_end.text())):
            self.add_log(f"💤 休眠中 ({now_h}点)"); self.sync_timer.start(30 * 60000); return
        self.run_shell(); self.sync_timer.start(int(self.ui_freq.text()) * 60000)

    def run_shell(self):
        ps_dir = self.share_dir.replace('"', '""'); ps_kw = self.target_kw.replace('"', '""'); ps_tmp = self.tmp_log.replace('"', '""')
        count = self.ui_count.text().strip()
        ps_cmd = f"""
        try {{
            if (!(Test-Path "{ps_dir}")) {{ New-Item -ItemType Directory -Path "{ps_dir}" -Force | Out-Null }}
            $ol = New-Object -ComObject Outlook.Application
            $it = $ol.GetNamespace("MAPI").GetDefaultFolder(6).Items | Where-Object {{ $_.ReceivedTime -gt (Get-Date).AddDays(-3) -and ($_.Subject -like "*{ps_kw}*" -or $_.SenderName -like "*{ps_kw}*") }} | Sort-Object ReceivedTime -Descending | Select-Object -First {count}
            if ($it) {{ foreach($m in $it) {{ $n = ($m.Subject -replace '[\\x00-\\x1f\\\\/:*?"<>|]', '_').Trim(); $p = Join-Path "{ps_dir}" "$n.mht"; if (!(Test-Path $p)) {{ $m.SaveAs($p, 10) }} }} }}
        }} catch {{ "ERR|$($_.Exception.Message)" | Out-File "{ps_tmp}" -Encoding utf8 }} finally {{ if ($ol) {{ [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ol) | Out-Null }} }}
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
            p_m = os.path.join(self.share_dir, f); p_h = os.path.join(self.share_dir, f.replace('.mht', '.html'))
            try:
                with open(p_m, 'rb') as fp:
                    msg = email.message_from_binary_file(fp)
                    for part in msg.walk():
                        if part.get_content_type() == "text/html":
                            with open(p_h, 'w', encoding='utf-8') as hw: hw.write(part.get_payload(decode=True).decode('utf-8','ignore'))
                            break
                os.remove(p_m); self.add_log(f"✅ Sync: {f}")
            except: pass
        self.build_index()

    def build_index(self):
        files = [f for f in os.listdir(self.share_dir) if f.endswith('.html') and f != 'index.html']
        files.sort(key=lambda x: os.path.getmtime(os.path.join(self.share_dir, x)), reverse=True)
        c = self.ui_color.text().strip()
        items = ""
        for f in files:
            p = os.path.join(self.share_dir, f)
            mt = time.strftime("%m-%d %H:%M", time.localtime(os.path.getmtime(p)))
            ep = list(set(re.findall(r'\bEP[A-Z0-9]{9}\b', open(p, 'r', encoding='utf-8', errors='ignore').read(5000), re.I)))
            tags = "".join([f'<span style="background:{c};color:white;padding:2px 5px;border-radius:3px;font-size:10px;margin-right:5px;">{x}</span>' for x in ep[:2]])
            items += f'<div class="item" onclick="v(this,\'{urllib.parse.quote(f)}\')" data-s="{(f+str(ep)).lower()}"><b>{f[:35]}...</b><br>{tags}<br><span class="time">{mt}</span></div>'

        # 网页 UI：核心增加了容器边框布局
        web_ui = f"""
        <!DOCTYPE html><html><head><meta charset='utf-8'>
        <style>
            :root {{ --main: {c}; }}
            body {{ 
                margin: 0; padding: 12px; background: #e9e9e9; height: 100vh; 
                display: flex; box-sizing: border-box; font-family: 'Segoe UI', sans-serif;
            }}
            .app-container {{ 
                display: flex; width: 100%; height: 100%; background: white; 
                border: 1px solid #ccc; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 15px rgba(0,0,0,0.1);
            }}
            .sidebar {{ width: 320px; border-right: 1px solid #eee; display: flex; flex-direction: column; background: #fff; }}
            .header {{ background: var(--main); color: white; padding: 18px; font-weight: bold; }}
            .list {{ flex: 1; overflow-y: auto; }}
            .item {{ padding: 15px; border-bottom: 1px solid #f9f9f9; cursor: pointer; transition: 0.2s; }}
            .item:hover {{ background: #fcfcfc; }}
            .item.active {{ background: #f0f7ff; border-left: 5px solid var(--main); }}
            .item b {{ font-size: 13px; color: #333; }}
            .time {{ font-size: 11px; color: #999; display: block; margin-top: 5px; }}
            .preview {{ flex: 1; background: #fff; display: flex; flex-direction: column; }}
            iframe {{ flex: 1; border: none; }}
            .footer {{ font-size: 10px; padding: 10px; text-align: center; color: #bbb; border-top: 1px solid #eee; }}
        </style>
        <script>
            function ds(v) {{ v=v.toLowerCase(); document.querySelectorAll('.item').forEach(i=>{{ i.style.display=i.getAttribute('data-s').includes(v)?'block':'none'; }}); }}
            function v(el, url) {{ document.querySelectorAll('.item').forEach(i=>i.classList.remove('active')); el.classList.add('active'); document.getElementById('f').src=url; }}
        </script></head>
        <body>
            <div class="app-container">
                <div class="sidebar">
                    <div class="header">📫 RD 邮件分发看板</div>
                    <div style="padding:10px;"><input onkeyup="ds(this.value)" placeholder="🔍 搜索单号或标题..." style="width:100%;padding:8px;box-sizing:border-box;border:1px solid #ddd;border-radius:4px;"></div>
                    <div class="list">{items}</div>
                    <div class="footer">{self.ui_copy.text()}</div>
                </div>
                <div class="preview"><iframe id="f"></iframe></div>
            </div>
        </body></html>
        """
        with open(os.path.join(self.share_dir, "index.html"), "w", encoding="utf-8") as fx: fx.write(web_ui)
        self.add_log("📊 网页边框 UI 已更新部署")

    def init_tray(self):
        self.tray = QSystemTrayIcon(self); self.tray.setIcon(self.style().standardIcon(21))
        m = QMenu(); m.addAction("Show", self.showNormal); m.addAction("Exit", QApplication.instance().quit); self.tray.setContextMenu(m); self.tray.show()

if __name__ == "__main__":
    app = QApplication(sys.argv); ex = OutlookMHTMaster(); ex.show(); sys.exit(app.exec_())
