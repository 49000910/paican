import sys, os, re, time, subprocess
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QTextEdit, QSystemTrayIcon, QMenu)
from PyQt5.QtCore import QTimer, Qt

class OutlookMHTMaster(QWidget):
    def __init__(self):
        super().__init__()
        # --- 默认配置 ---
        self.share_dir = r'\\10.1.93.32\DT_HU_RDteam_F\视频\Z\ZOUQIU\paican'
        self.target_kw = 'EDFA' 
        self.copyright_text = "© 2024-2026 RD Team | 视频组技术支持"
        self.interval_min = 10     
        self.start_hour = 8        
        self.end_hour = 18         
        
        self.init_ui()
        self.init_tray()
        
        self.sync_timer = QTimer(self)
        self.sync_timer.timeout.connect(self.run_cycle)
        QTimer.singleShot(2000, self.run_cycle)

    def init_ui(self):
        self.setWindowTitle("RD 邮件全自动 MHT 同步助手 (预览增强版)")
        self.resize(550, 500)
        self.setWindowFlags(Qt.WindowStaysOnTopHint)
        layout = QVBoxLayout()
        
        layout.addWidget(QLabel("📂 共享保存目录 (MHT 模式):"))
        self.edit_path = QLineEdit(self.share_dir); layout.addWidget(self.edit_path)

        layout.addWidget(QLabel("📧 监控关键词 (发件人/标题):"))
        self.edit_kw = QLineEdit(self.target_kw); layout.addWidget(self.edit_kw)

        h_time_layout = QHBoxLayout()
        h_time_layout.addWidget(QLabel("⏰ 活跃时段:")); self.edit_start = QLineEdit(str(self.start_hour)); h_time_layout.addWidget(self.edit_start)
        h_time_layout.addWidget(QLabel("至")); self.edit_end = QLineEdit(str(self.end_hour)); h_time_layout.addWidget(self.edit_end)
        layout.addLayout(h_time_layout)

        h_freq_layout = QHBoxLayout()
        h_freq_layout.addWidget(QLabel("⏱ 频率(分):")); self.edit_freq = QLineEdit(str(self.interval_min)); h_freq_layout.addWidget(self.edit_freq)
        layout.addLayout(h_freq_layout)

        layout.addWidget(QLabel("📝 网页底部版权修改:"))
        self.edit_copy = QLineEdit(self.copyright_text); layout.addWidget(self.edit_copy)

        self.btn_apply = QPushButton("🚀 保存配置并立即同步")
        self.btn_apply.setStyleSheet("background: #0078d4; color: white; font-weight: bold; padding: 12px; border-radius: 5px;")
        self.btn_apply.clicked.connect(self.apply_settings)
        layout.addWidget(self.btn_apply)

        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setStyleSheet("background: #1e1e1e; color: #00ff00; font-family: 'Consolas'; font-size: 12px;")
        layout.addWidget(self.log_area)
        self.setLayout(layout)

    def init_tray(self):
        self.tray = QSystemTrayIcon(self)
        self.tray.setIcon(self.style().standardIcon(21)) 
        menu = QMenu()
        menu.addAction("显示界面", self.showNormal)
        menu.addAction("退出程序", QApplication.instance().quit)
        self.tray.setContextMenu(menu)
        self.tray.show()

    def closeEvent(self, event):
        self.hide(); event.ignore()

    def add_log(self, text):
        self.log_area.append(f"[{time.strftime('%H:%M:%S')}] {text}")

    def apply_settings(self):
        self.share_dir = self.edit_path.text().strip()
        self.target_kw = self.edit_kw.text().strip()
        self.copyright_text = self.edit_copy.text().strip()
        try:
            self.interval_min = int(self.edit_freq.text())
            self.start_hour = int(self.edit_start.text())
            self.end_hour = int(self.edit_end.text())
            self.add_log("✅ 配置已更新")
            self.run_cycle()
        except: self.add_log("❌ 数字格式错误")

    def run_cycle(self):
        now_hour = int(time.strftime("%H"))
        if not (self.start_hour <= now_hour < self.end_hour):
            self.add_log(f"💤 非时段 ({now_hour}点)，静默中...")
            self.sync_timer.start(30 * 60000)
            return
        self.run_shell_mht_logic()
        self.sync_timer.start(self.interval_min * 60000)

    def run_shell_mht_logic(self):
        """核心：通过 PowerShell 驱动 Outlook COM 接口"""
        safe_dir = self.share_dir.replace('"', '""')
        kw = self.target_kw.replace('"', '""')
        
        ps_cmd = f"""
        $ErrorActionPreference = "Stop"
        try {{
            if (!(Test-Path "{safe_dir}")) {{ New-Item -ItemType Directory -Path "{safe_dir}" -Force }}
            $ol = New-Object -ComObject Outlook.Application
            $ns = $ol.GetNamespace("MAPI")
            # 搜索最近3天的匹配邮件
            $limit = (Get-Date).AddDays(-3)
            $it = $ns.GetDefaultFolder(6).Items | Where-Object {{ $_.ReceivedTime -gt $limit -and ($_.Subject -like "*{kw}*" -or $_.SenderName -like "*{kw}*") }} | Sort-Object ReceivedTime -Descending | Select-Object -First 1
            
            if ($it) {{
                $name = ($it.Subject -replace '[\x00-\x1f\\\\/:*?"<>|]', '_').Trim()
                if ($name.Length -gt 50) {{ $name = $name.Substring(0,50) }}
                $path = Join-Path "{safe_dir}" "$name.mht"
                
                if (!(Test-Path $path)) {{
                    $it.SaveAs($path, 10)
                    Write-Host "RESULT_SUCCESS: $name"
                }} else {{ Write-Host "RESULT_EXISTS" }}
            }} else {{ Write-Host "RESULT_NOTFOUND" }}
        }} catch {{
            Write-Host "RESULT_ERROR: $($_.Exception.Message)"
        }} finally {{
            if ($ol) {{ [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ol) | Out-Null }}
        }}
        """
        try:
            res = subprocess.run(["powershell", "-NoProfile", "-Command", ps_cmd], capture_output=True, text=True, encoding='gbk', creationflags=0x08000000, timeout=45)
            out = res.stdout.strip()
            if "RESULT_SUCCESS" in out: self.add_log(f"✅ 抓取成功: {out.split(': ')[1]}")
            elif "RESULT_EXISTS" in out: self.add_log("ℹ️ 无新邮件（已存在）")
            elif "RESULT_NOTFOUND" in out: self.add_log("❓ 未发现匹配邮件")
            elif "RESULT_ERROR" in out: self.add_log(f"❌ Shell 报错: {out}")
            self.generate_html_index()
        except Exception as e: self.add_log(f"❌ 调用异常: {e}")

    def generate_html_index(self):
        """生成左侧列表、右侧预览的 index.html"""
        if not os.path.exists(self.share_dir): return
        files = [f for f in os.listdir(self.share_dir) if f.endswith('.mht')]
        files.sort(key=lambda x: os.path.getmtime(os.path.join(self.share_dir, x)), reverse=True)
        
        html = f"""
        <!DOCTYPE html><html><head><meta charset="utf-8">
        <style>
            body {{ margin: 0; display: flex; height: 100vh; font-family: 'Segoe UI', sans-serif; background: #eee; overflow: hidden; }}
            .sidebar {{ width: 320px; background: #fff; border-right: 1px solid #ddd; display: flex; flex-direction: column; }}
            .header {{ background: #0078d4; color: white; padding: 15px; font-size: 14px; font-weight: bold; }}
            .list {{ flex: 1; overflow-y: auto; }}
            .item {{ display: block; padding: 12px; color: #333; text-decoration: none; border-bottom: 1px solid #eee; font-size: 13px; cursor: pointer; }}
            .item:hover {{ background: #f0f7ff; }}
            .item.active {{ background: #e1f0fe; border-left: 4px solid #0078d4; }}
            .time {{ font-size: 11px; color: #999; display: block; margin-top: 4px; }}
            .preview {{ flex: 1; background: #fff; display: flex; flex-direction: column; }}
            iframe {{ width: 100%; height: 100%; border: none; }}
            .empty {{ margin: auto; color: #ccc; text-align: center; }}
            .footer {{ font-size: 11px; padding: 10px; color: #999; border-top: 1px solid #eee; text-align: center; }}
        </style>
        <script>
            function view(el, url) {{
                document.querySelectorAll('.item').forEach(i => i.classList.remove('active'));
                el.classList.add('active');
                document.getElementById('frame').src = url;
                document.getElementById('msg').style.display = 'none';
            }}
        </script></head>
        <body>
            <div class="sidebar">
                <div class="header">📫 RD 邮件分发看板<br><small style="font-weight:normal;opacity:0.8">关键词: {self.target_kw}</small></div>
                <div class="list">
        """
        for i, f in enumerate(files):
            t = time.strftime("%m-%d %H:%M", time.localtime(os.path.getmtime(os.path.join(self.share_dir, f))))
            html += f'<div class="item" onclick="view(this, \'{f}\')"><b>{f[:40]}...</b><span class="time">🕒 {t}</span></div>'
            
        html += f"""
                </div><div class="footer">{self.copyright_text}</div>
            </div>
            <div class="preview">
                <div id="msg" class="empty"><h3>请从左侧选择邮件查看</h3></div>
                <iframe id="frame"></iframe>
            </div>
        </body></html>
        """
        with open(os.path.join(self.share_dir, "index.html"), "w", encoding="utf-8") as f: f.write(html)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = OutlookMHTMaster()
    window.show()
    sys.exit(app.exec_())
