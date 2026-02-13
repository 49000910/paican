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
        # 启动即执行第一次
        QTimer.singleShot(2000, self.run_cycle)

    def init_ui(self):
        self.setWindowTitle("RD 邮件全自动 MHT 同步助手 (V2.0 预览版)")
        self.resize(580, 520)
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

        self.btn_apply = QPushButton("🚀 保存并开始全自动同步")
        self.btn_apply.setStyleSheet("background: #0078d4; color: white; font-weight: bold; padding: 12px; border-radius: 4px;")
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
        menu.addAction("打开主界面", self.showNormal)
        menu.addAction("退出程序", QApplication.instance().quit)
        self.tray.setContextMenu(menu)
        self.tray.show()

    def closeEvent(self, event):
        self.hide(); event.ignore()

    def add_log(self, text):
        # 彻底拦截 Python 端的 Null Character 报错
        safe_text = str(text).replace('\x00', '')
        self.log_area.append(f"[{time.strftime('%H:%M:%S')}] {safe_text}")

    def apply_settings(self):
        self.share_dir = self.edit_path.text().strip()
        self.target_kw = self.edit_kw.text().strip()
        self.copyright_text = self.edit_copy.text().strip()
        try:
            self.interval_min = int(self.edit_freq.text())
            self.start_hour = int(self.edit_start.text())
            self.end_hour = int(self.edit_end.text())
            self.add_log("✅ 配置已更新。")
            self.run_cycle()
        except: self.add_log("❌ 数字格式错误，请检查。")

    def run_cycle(self):
        now_hour = int(time.strftime("%H"))
        if not (self.start_hour <= now_hour < self.end_hour):
            self.add_log(f"💤 非活跃时段 ({now_hour}点)，进入静默模式...")
            self.sync_timer.start(30 * 60000)
            return
        self.run_shell_mht_logic()
        self.sync_timer.start(self.interval_min * 60000)

    def run_shell_mht_logic(self):
        """核心：全量 Shell 抓取搜索逻辑"""
        # 对路径和关键词进行 Shell 安全处理
        ps_dir = self.share_dir.replace('\\', '\\\\').replace('"', '""')
        ps_kw = self.target_kw.replace('"', '""')
        
        # 封装全套 PowerShell 动作：搜索 -> 过滤 -> 去重 -> 保存 -> 释放内存
        ps_cmd = f"""
        $ErrorActionPreference = "Stop"
        try {{
            if (!(Test-Path "{ps_dir}")) {{ New-Item -ItemType Directory -Path "{ps_dir}" -Force | Out-Null }}
            
            $ol = New-Object -ComObject Outlook.Application
            $ns = $ol.GetNamespace("MAPI")
            $limit = (Get-Date).AddDays(-3)
            
            # 搜索：最近3天 + 关键词匹配 (标题或发件人)
            $it = $ns.GetDefaultFolder(6).Items | Where-Object {{ 
                $_.ReceivedTime -gt $limit -and ($_.Subject -like "*{ps_kw}*" -or $_.SenderName -like "*{ps_kw}*") 
            }} | Sort-Object ReceivedTime -Descending | Select-Object -First 1
            
            if ($it) {{
                # 剔除文件名中的特殊字符和空字符 (\x00)
                $name = $it.Subject -replace '[\\x00-\\x1f\\\\/:*?"<>|]', '_'
                $name = $name.Trim()
                if ($name.Length -gt 50) {{ $name = $name.Substring(0,50) }}
                
                $path = Join-Path "{ps_dir}" "$name.mht"
                if (!(Test-Path $path)) {{
                    $it.SaveAs($path, 10) # 10 = olMHTML
                    [Console]::WriteLine("SHELL_RES:SUCCESS:" + $name)
                }} else {{ [Console]::WriteLine("SHELL_RES:EXISTS") }}
            }} else {{ [Console]::WriteLine("SHELL_RES:NOTFOUND") }}
        }} catch {{
            $msg = $_.Exception.Message -replace '[\\x00-\\x1f]', ''
            [Console]::WriteLine("SHELL_RES:ERROR:" + $msg)
        }} finally {{
            # 必须释放，否则 Outlook 进程会卡死后台
            if ($ol) {{ 
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ol) | Out-Null
                [GC]::Collect(); [GC]::WaitForPendingFinalizers() 
            }}
        }}
        """
        try:
            # 增加 errors='replace' 防止因特殊字符抛出 ValueError: embedded null character
            res = subprocess.run(
                ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_cmd], 
                capture_output=True, text=True, encoding='gbk', errors='replace',
                creationflags=0x08000000, timeout=50
            )
            out = res.stdout.replace('\x00', '').strip()
            
            if "SHELL_RES:SUCCESS:" in out:
                self.add_log(f"✅ 抓取成功: {out.split('SUCCESS:')[-1]}")
            elif "SHELL_RES:EXISTS" in out:
                self.add_log("ℹ️ 邮件已在共享盘，跳过同步。")
            elif "SHELL_RES:NOTFOUND" in out:
                self.add_log(f"❓ 未发现匹配 '{self.target_kw}' 的新邮件。")
            elif "SHELL_RES:ERROR:" in out:
                self.add_log(f"❌ Shell 逻辑报错: {out.split('ERROR:')[-1]}")
            
            self.generate_html_index()
        except Exception as e:
            self.add_log(f"❌ 系统异常: {str(e)}")

    def generate_html_index(self):
        """生成支持左右分栏预览的 index.html"""
        if not os.path.exists(self.share_dir): return
        files = [f for f in os.listdir(self.share_dir) if f.endswith('.mht')]
        files.sort(key=lambda x: os.path.getmtime(os.path.join(self.share_dir, x)), reverse=True)
        
        html = f"""
        <!DOCTYPE html><html><head><meta charset="utf-8">
        <title>RD 邮件预览看板</title>
        <style>
            body {{ margin: 0; padding: 0; font-family: 'Segoe UI', sans-serif; display: flex; height: 100vh; overflow: hidden; }}
            .sidebar {{ width: 350px; background: #fdfdfd; border-right: 1px solid #ddd; display: flex; flex-direction: column; }}
            .header {{ background: #0078d4; color: white; padding: 18px; }}
            .list {{ flex: 1; overflow-y: auto; }}
            .item {{ display: block; padding: 15px; border-bottom: 1px solid #eee; text-decoration: none; color: #333; cursor: pointer; }}
            .item:hover {{ background: #f0f7ff; }}
            .item.active {{ background: #e1f0fe; border-left: 4px solid #0078d4; }}
            .time {{ font-size: 11px; color: #888; display: block; margin-top: 5px; }}
            .tag {{ background: #ffeb3b; color: #333; padding: 2px 5px; font-size: 10px; border-radius: 3px; font-weight: bold; }}
            .preview {{ flex: 1; background: #fff; position: relative; }}
            iframe {{ width: 100%; height: 100%; border: none; }}
            .empty {{ position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); color: #bbb; text-align: center; }}
            .footer {{ background: #f5f5f5; padding: 10px; text-align: center; font-size: 11px; color: #999; }}
        </style>
        <script>
            function openMail(el, url) {{
                document.querySelectorAll('.item').forEach(i => i.classList.remove('active'));
                el.classList.add('active');
                document.getElementById('frame').src = url;
                document.getElementById('placeholder').style.display = 'none';
            }}
        </script></head>
        <body>
            <div class="sidebar">
                <div class="header"><div style="font-size:16px;">📫 RD 邮件全自动看板</div><small>关键词: {self.target_kw}</small></div>
                <div class="list">
        """
        for i, f in enumerate(files):
            mtime = time.strftime("%Y-%m-%d %H:%M", time.localtime(os.path.getmtime(os.path.join(self.share_dir, f))))
            tag = '<span class="tag">NEW</span>' if i == 0 else ""
            html += f'''
                <div class="item" onclick="openMail(this, '{f}')">
                    {tag} <b>{f[:40]}...</b>
                    <span class="time">🕒 {mtime}</span>
                </div>
            '''
        html += f"""
                </div><div class="footer">{self.copyright_text}</div>
            </div>
            <div class="preview">
                <div id="placeholder" class="empty"><h2>🔍 点击左侧查看邮件内容</h2></div>
                <iframe id="frame"></iframe>
            </div>
        </body></html>
        """
        with open(os.path.join(self.share_dir, "index.html"), "w", encoding="utf-8") as f: f.write(html)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = OutlookMHTMaster()
    win.show()
    sys.exit(app.exec_())
