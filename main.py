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
        # --- 默认参数 (完全保留您的定义) ---
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
        self.setWindowTitle("EDFA 看板后台 V54.5")
        self.resize(520, 900)
        layout = QVBoxLayout()
        layout.setContentsMargins(15, 15, 15, 15)
        def quick_edit(label, val, attr):
            l = QHBoxLayout(); lb = QLabel(label); lb.setFixedWidth(110); l.addWidget(lb)
            edit = QLineEdit(str(val)); setattr(self, attr, edit); l.addWidget(edit); layout.addLayout(l)
        
        # --- UI 布局完全保留 ---
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
        h2.addWidget(QLabel("-")); self.ui_end = QLineEdit(str(self.ui_end_hour if hasattr(self, 'ui_end_hour') else self.end_hour)); h2.addWidget(self.ui_end)
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
        # --- 保留原有的 PowerShell 逻辑 ---
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
        
        # 1. 提取 Excel 并生成 JS 变量（静态局部更新逻辑）
        excel_html = "<tr><td>未发现2026日历文件</td></tr>"
        for f_name in os.listdir(d):
            if "2026日历" in f_name and f_name.lower().endswith('.xlsx'):
                try:
                    df = pd.read_excel(os.path.join(d, f_name))
                    # 转为 HTML 字符串，去掉换行符以便 JS 读取
                    excel_html = df.to_html(index=False, border=0, classes='calendar-table', escape=False).replace('\n', '')
                    break
                except Exception as e: self.add_log(f"Excel解析异常: {e}")

        # 写入独立的 JS 数据文件
        with open(os.path.join(d, "data_store.js"), "w", encoding="utf-8") as f:
            f.write(f"var WEB_CONTENT = `{excel_html}`;\n")
            f.write(f"var UPDATE_TIME = '{time.strftime('%Y-%m-%d %H:%M:%S')}';")

        # 2. 生成 Index.html (基于 JS 变量注入，支持局域网 file 协议和索引)
        web_path = os.path.join(d, "index.html")
        color = self.ui_color.text()
        full_html = f"""
        <html><head><meta charset="utf-8">
            <title>{self.ui_title.text()}</title>
            <style>
                body {{ font-family: 'Microsoft YaHei', sans-serif; background: #f0f2f5; margin: 0; padding: 20px; }}
                .card {{ background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); max-width: 1400px; margin: auto; }}
                h1 {{ color: {color}; border-bottom: 2px solid {color}; padding-bottom: 10px; }}
                .calendar-table {{ width: 100%; border-collapse: collapse; margin-top: 15px; }}
                .calendar-table th, .calendar-table td {{ border: 1px solid #ddd; padding: 8px; text-align: center; }}
                .calendar-table tr:nth-child(even) {{ background: #fafafa; }}
                #sync-info {{ color: #888; font-size: 12px; margin-top: 10px; }}
            </style>
        </head>
        <body>
            <div class="card">
                <h1>{self.ui_title.text()}</h1>
                <p>{self.ui_subtitle.text()}</p>
                <div id="data-box">正在同步数据...</div>
                <div id="sync-info">最后同步时间：<span id="time-str">-</span></div>
                <hr><div style="text-align:center; color:#999; font-size:11px;">{self.ui_copy.text()}</div>
            </div>
            <script>
                function loadData() {{
                    var script = document.createElement('script');
                    script.src = 'data_store.js?t=' + new Date().getTime();
                    script.onload = function() {{
                        document.getElementById('data-box').innerHTML = WEB_CONTENT;
                        document.getElementById('time-str').innerText = UPDATE_TIME;
                    }};
                    document.head.appendChild(script);
                }}
                loadData();
                setInterval(loadData, {int(self.ui_web_freq.text()) * 1000});
            </script>
        </body></html>
        """
        with open(web_path, 'w', encoding='utf-8') as f: f.write(full_html)
        self.add_log("🌐 网页已静默更新 (支持局域网搜索/定位)")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = OutlookMHTMaster()
    window.show()
    sys.exit(app.exec_())
