import sys, os, re, time, subprocess, tempfile, urllib.parse, base64, email
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QTextEdit, QSystemTrayIcon, QMenu, QComboBox)
from PyQt5.QtCore import QTimer, Qt

class OutlookMHTMaster(QWidget):
    def __init__(self):
        super().__init__()
        # --- 默认配置 ---
        self.share_dir = r'\\10.1.93.32\DT_HU_RDteam_F\视频\Z\ZOUQIU\paican'
        self.target_kw = 'EDFA' 
        self.interval_min = 10     
        self.sys_title = "RD Team 邮件看板"
        self.copyright_text = "© 2024-2026 RD Team | 视频组技术支持"
        
        # --- 颜色方案 ---
        self.themes = {
            "默认科技蓝": "#0078d4",
            "生产安全绿": "#107c10",
            "商务极客黑": "#201f1e",
            "警告活力橙": "#d83b01",
            "自定义颜色": "CUSTOM"
        }
        self.current_theme_color = "#0078d4"
        
        self.tmp_log = os.path.join(tempfile.gettempdir(), "sync_v22_res.txt")
        self.last_sync_time = "尚未同步"
        
        self.init_ui()
        self.init_tray()
        
        self.sync_timer = QTimer(self)
        self.sync_timer.timeout.connect(self.run_cycle)
        QTimer.singleShot(2000, self.run_cycle)

    def init_ui(self):
        self.setWindowTitle("RD 邮件助手 V22.0 (全界面联动换色版)")
        self.resize(600, 750)
        self.main_layout = QVBoxLayout()
        
        # 配置表单
        form = QVBoxLayout()
        self.edit_path = QLineEdit(self.share_dir); form.addWidget(QLabel("📂 共享网络路径 (UNC):")); form.addWidget(self.edit_path)
        
        h1 = QHBoxLayout()
        self.edit_kw = QLineEdit(self.target_kw); h1.addWidget(QLabel("📧 关键词:")); h1.addWidget(self.edit_kw)
        self.edit_freq = QLineEdit(str(self.interval_min)); h1.addWidget(QLabel("⏱ 间隔(分):")); h1.addWidget(self.edit_freq)
        form.addLayout(h1)
        
        h2 = QHBoxLayout()
        self.edit_title = QLineEdit(self.sys_title); h2.addWidget(QLabel("🏷️ 网页标题:")); h2.addWidget(self.edit_title)
        self.combo_theme = QComboBox(); self.combo_theme.addItems(self.themes.keys())
        h2.addWidget(QLabel("🎨 预设主题:")); h2.addWidget(self.combo_theme)
        form.addLayout(h2)

        h3 = QHBoxLayout()
        self.edit_custom_color = QLineEdit(self.current_theme_color)
        h3.addWidget(QLabel("🌈 HEX 颜色代码:")); h3.addWidget(self.edit_custom_color)
        form.addLayout(h3)

        self.edit_copy = QLineEdit(self.copyright_text); form.addWidget(QLabel("📝 版权信息:")); form.addWidget(self.edit_copy)
        self.main_layout.addLayout(form)

        # 同步按钮
        self.btn_apply = QPushButton("🚀 保存配置并应用视觉主题")
        self.btn_apply.setFixedHeight(50)
        self.btn_apply.clicked.connect(self.apply_settings)
        self.main_layout.addWidget(self.btn_apply)

        # 日志区
        self.log_area = QTextEdit(); self.log_area.setReadOnly(True)
        self.main_layout.addWidget(self.log_area)
        
        self.setLayout(self.main_layout)
        self.update_app_style() # 初始化样式

    def update_app_style(self):
        """核心：动态更新后台界面的 QSS 样式"""
        c = self.current_theme_color
        # 确保颜色以 # 开头
        if not c.startswith("#"): c = "#0078d4"
        
        qss = f"""
            QWidget {{ background-color: #f5f5f5; font-family: 'Microsoft YaHei'; }}
            QLabel {{ color: #333; font-weight: bold; }}
            QLineEdit {{ padding: 6px; border: 2px solid #ddd; border-radius: 4px; background: white; }}
            QLineEdit:focus {{ border-color: {c}; }}
            QPushButton {{ background-color: {c}; color: white; border-radius: 6px; font-size: 14px; font-weight: bold; border: none; }}
            QPushButton:hover {{ background-color: {c}cc; }}
            QTextEdit {{ background-color: #1e1e1e; color: #00ff00; border: 3px solid {c}; border-radius: 5px; font-family: 'Consolas'; }}
            QComboBox {{ padding: 5px; border: 2px solid #ddd; border-radius: 4px; }}
        """
        self.setStyleSheet(qss)

    def init_tray(self):
        self.tray = QSystemTrayIcon(self); self.tray.setIcon(self.style().standardIcon(21))
        m = QMenu(); m.addAction("显示", self.showNormal); m.addAction("退出", QApplication.instance().quit)
        self.tray.setContextMenu(m); self.tray.show()

    def add_log(self, text):
        clean_msg = str(text).replace('\x00', '')
        self.log_area.append(f"[{time.strftime('%H:%M:%S')}] {clean_msg}")

    def apply_settings(self):
        self.share_dir, self.target_kw = self.edit_path.text().strip(), self.edit_kw.text().strip()
        self.sys_title, self.copyright_text = self.edit_title.text().strip(), self.edit_copy.text().strip()
        
        # 识别颜色
        tk = self.combo_theme.currentText()
        if self.themes[tk] == "CUSTOM":
            self.current_theme_color = self.edit_custom_color.text().strip()
        else:
            self.current_theme_color = self.themes[tk]
            self.edit_custom_color.setText(self.current_theme_color)
        
        try:
            self.interval_min = int(self.edit_freq.text())
            self.update_app_style() # 即时更新后台颜色
            self.add_log(f"🎨 主题已切换为: {self.current_theme_color}")
            self.run_cycle()
        except: self.add_log("❌ 数值错误")

    def run_cycle(self):
        self.run_shell_logic()
        self.last_sync_time = time.strftime('%Y-%m-%d %H:%M:%S')
        self.sync_timer.start(self.interval_min * 60000)

    def run_shell_logic(self):
        ps_dir = self.share_dir.replace('"', '""'); ps_kw = self.target_kw.replace('"', '""'); ps_tmp = self.tmp_log.replace('"', '""')
        ps_script = f"""
        try {{
            if (!(Test-Path "{ps_dir}")) {{ New-Item -ItemType Directory -Path "{ps_dir}" -Force | Out-Null }}
            $ol = New-Object -ComObject Outlook.Application
            $ns = $ol.GetNamespace("MAPI")
            $it = $ns.GetDefaultFolder(6).Items | Where-Object {{ $_.ReceivedTime -gt (Get-Date).AddDays(-3) -and ($_.Subject -like "*{ps_kw}*" -or $_.SenderName -like "*{ps_kw}*") }} | Sort-Object ReceivedTime -Descending | Select-Object -First 1
            if ($it) {{
                $n = $it.Subject -replace '[\\x00-\\x1f\\\\/:*?"<>|]', '_'
                $p = Join-Path "{ps_dir}" "$($n.Trim()).mht"
                if (!(Test-Path $p)) {{ $it.SaveAs($p, 10); "SUCCESS|$n" | Out-File "{ps_tmp}" -Encoding utf8 }}
                else {{ "EXISTS" | Out-File "{ps_tmp}" -Encoding utf8 }}
            }} else {{ "NOTFOUND" | Out-File "{ps_tmp}" -Encoding utf8 }}
        }} catch {{ "ERROR|$($_.Exception.Message)" | Out-File "{ps_tmp}" -Encoding utf8 }} finally {{ if ($ol) {{ [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ol) | Out-Null }} }}
        """
        try:
            ps_b64 = base64.b64encode(ps_script.encode('utf-16-le')).decode('ascii')
            subprocess.run(["powershell", "-NoProfile", "-EncodedCommand", ps_b64], creationflags=0x08000000, timeout=60)
            self.convert_and_index()
        except Exception as e: self.add_log(f"同步异常: {e}")

    def convert_and_index(self):
        if not os.path.exists(self.share_dir): return
        mht_files = [f for f in os.listdir(self.share_dir) if f.endswith('.mht')]
        items_html = ""
        for f in mht_files:
            mht_p, h_n = os.path.join(self.share_dir, f), f.replace('.mht', '.static.html')
            h_p = os.path.join(self.share_dir, h_n)
            mt = time.strftime("%m-%d %H:%M", time.localtime(os.path.getmtime(mht_p)))
            ep_list, mail_body = [], ""
            try:
                with open(mht_p, 'rb') as fp:
                    msg = email.message_from_binary_file(fp)
                    for part in msg.walk():
                        if part.get_content_type() == "text/html":
                            payload = part.get_payload(decode=True)
                            for enc in ['utf-8', 'gbk', 'gb18030']:
                                try:
                                    mail_body = payload.decode(enc)
                                    plain = re.sub('<[^<]+?>', '', mail_body)
                                    ep_list = list(set(re.findall(r'\bEP[A-Z0-9]{9}\b', plain, re.I)))
                                    break
                                except: continue
                if mail_body and not os.path.exists(h_p):
                    with open(h_p, 'w', encoding='utf-8') as hw: hw.write(mail_body)
            except: pass
            tags = "".join([f'<span class="tag">{c}</span>' for c in ep_list[:3]])
            safe_f = urllib.parse.quote(h_n)
            sk = (f + " " + " ".join(ep_list)).lower()
            items_html += f'<div class="item" onclick="v(this,\'{safe_f}\')" data-s="{sk}"><b>{f[:45]}</b>{tags}<span class="time">🕒 {mt}</span></div>'

        full_html = f"""
        <!DOCTYPE html><html><head><meta charset='utf-8'><title>{self.sys_title}</title>
        <style>
            :root {{ --main: {self.current_theme_color}; }}
            body {{ margin:0; display:flex; height:100vh; font-family:'Segoe UI',sans-serif; overflow:hidden; background:#f4f4f4; }}
            .sidebar {{ width:340px; border-right:1px solid #ddd; display:flex; flex-direction:column; background:#fff; }}
            .header {{ background: var(--main); color:white; padding:15px; }}
            .search-box {{ padding:10px; border-bottom:1px solid #eee; }}
            .search-box input {{ width:100%; padding:8px; border:1px solid #ccc; border-radius:4px; outline:none; }}
            .list {{ flex:1; overflow-y:auto; }}
            .item {{ padding:12px; border-bottom:1px solid #eee; cursor:pointer; }}
            .item:hover {{ background:#f9f9f9; }}
            .item.active {{ background:#eff6fc; border-left:5px solid var(--main); }}
            .tag {{ background: var(--main); color:white; padding:1px 4px; border-radius:3px; font-size:10px; margin-right:5px; }}
            .time {{ font-size:11px; color:#999; display:block; margin-top:5px; }}
            iframe {{ width:100%; height:100%; border:none; background:#fff; }}
        </style>
        <script>
            function ds(v) {{ v=v.toLowerCase(); document.querySelectorAll('.item').forEach(i=>{{ i.style.display=i.getAttribute('data-s').includes(v)?'block':'none'; }}); }}
            function v(el, url) {{ document.querySelectorAll('.item').forEach(i=>i.classList.remove('active')); el.classList.add('active'); document.getElementById('f').src=url; }}
        </script></head>
        <body>
            <div class="sidebar">
                <div class="header"><b>{self.sys_title}</b><br><small style="font-size:10px;">最后更新: {self.last_sync_time}</small></div>
                <div class="search-box"><input onkeyup="ds(this.value)" placeholder="🔍 搜任务令或标题..."></div>
                <div class="list">{items_html}</div>
            </div>
            <div style="flex:1;"><iframe id="f"></iframe></div>
        </body></html>
        """
        with open(os.path.join(self.share_dir, "index.html"), "w", encoding="utf-8") as f_idx: f_idx.write(full_html)
        self.add_log(f"🚀 全界面视觉主题 [{self.current_theme_color}] 部署成功")

if __name__ == "__main__":
    app = QApplication(sys.argv); ex = OutlookMHTMaster(); ex.show(); sys.exit(app.exec_())
