import sys, os, re, time, subprocess
import win32com.client
import pythoncom
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QTextEdit, QSystemTrayIcon, QMenu)
from PyQt5.QtCore import QTimer, Qt

class OutlookMasterAuto(QWidget):
    def __init__(self):
        super().__init__()
        # --- 默认配置参数 ---
        self.share_dir = r'\\10.1.93.32\DT_HU_RDteam_F\视频\Z\ZOUQIU\paican'
        self.target_kw = 'EDFA' 
        self.copyright_text = "© 2024-2026 RD Team | 视频组技术支持"
        self.interval_min = 10     # 活跃期同步频率（分钟）
        self.start_hour = 8        # 开始监控小时 (24小时制)
        self.end_hour = 18         # 结束监控小时
        
        self.init_ui()
        self.init_tray()
        
        # 核心定时器
        self.sync_timer = QTimer(self)
        self.sync_timer.timeout.connect(self.run_sync_logic)
        
        # 启动后立即执行第一次（如果是监控时段）
        QTimer.singleShot(2000, self.run_sync_logic)

    def init_ui(self):
        self.setWindowTitle("RD 邮件自动分发看板 - 后台管理")
        self.resize(550, 500)
        self.setWindowFlags(Qt.WindowStaysOnTopHint)
        layout = QVBoxLayout()
        
        # UI 输入配置
        layout.addWidget(QLabel("📂 共享保存目录:"))
        self.edit_path = QLineEdit(self.share_dir)
        layout.addWidget(self.edit_path)

        layout.addWidget(QLabel("📧 监控关键词 (发件人/标题关键词):"))
        self.edit_kw = QLineEdit(self.target_kw)
        layout.addWidget(self.edit_kw)

        # 时间窗设置
        h_time_layout = QHBoxLayout()
        h_time_layout.addWidget(QLabel("⏰ 活跃时段 (时):"))
        self.edit_start = QLineEdit(str(self.start_hour))
        h_time_layout.addWidget(self.edit_start)
        h_time_layout.addWidget(QLabel("至"))
        self.edit_end = QLineEdit(str(self.end_hour))
        h_time_layout.addWidget(self.edit_end)
        layout.addLayout(h_time_layout)

        h_freq_layout = QHBoxLayout()
        h_freq_layout.addWidget(QLabel("⏱ 同步频率 (分钟):"))
        self.edit_freq = QLineEdit(str(self.interval_min))
        h_freq_layout.addWidget(self.edit_freq)
        layout.addLayout(h_freq_layout)

        layout.addWidget(QLabel("📝 网页版权修改:"))
        self.edit_copy = QLineEdit(self.copyright_text)
        layout.addWidget(self.edit_copy)

        self.btn_apply = QPushButton("🚀 保存并立即同步")
        self.btn_apply.setStyleSheet("background: #0078d4; color: white; font-weight: bold; padding: 10px; border-radius: 5px;")
        self.btn_apply.clicked.connect(self.apply_settings)
        layout.addWidget(self.btn_apply)

        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setStyleSheet("background: #1e1e1e; color: #00ff00; font-family: 'Consolas';")
        layout.addWidget(self.log_area)

        self.setLayout(layout)
        self.add_log("系统已就绪。监控时段外将进入静默模式。")

    def init_tray(self):
        self.tray = QSystemTrayIcon(self)
        self.tray.setIcon(self.style().standardIcon(21)) 
        menu = QMenu()
        menu.addAction("打开主界面", self.showNormal)
        menu.addAction("彻底退出程序", QApplication.instance().quit)
        self.tray.setContextMenu(menu)
        self.tray.activated.connect(lambda r: self.showNormal() if r == QSystemTrayIcon.Trigger else None)
        self.tray.show()

    def closeEvent(self, event):
        self.hide()
        event.ignore()

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
            self.add_log(f"✅ 配置已保存。活跃时段: {self.start_hour}-{self.end_hour}点")
            self.run_sync_logic()
        except:
            self.add_log("❌ 输入格式错误，请检查数字。")

    def run_sync_logic(self):
        """主控逻辑：检查时间窗 + 执行同步"""
        now_hour = int(time.strftime("%H"))
        
        # 检查是否在监控时段
        if not (self.start_hour <= now_hour < self.end_hour):
            self.add_log(f"💤 非监控时段 ({now_hour}点)，进入静默模式。")
            self.sync_timer.start(30 * 60000) # 30分钟后重新检查
            return

        self.add_log(f"🔄 活跃时段自动抓取中...")
        self.execute_secure_sync()
        # 设定下一次活跃期同步
        self.sync_timer.start(self.interval_min * 60000)

    def execute_secure_sync(self):
        """核心：全方位检查去重并调用 Shell 生存 MHT 文件"""
        try:
            pythoncom.CoInitialize()
            outlook = win32com.client.Dispatch("Outlook.Application")
            ns = outlook.GetNamespace("MAPI")
            inbox = ns.GetDefaultFolder(6)
            items = inbox.Items
            items.Sort("[ReceivedTime]", True)
            
            target_it = None
            kw = self.target_kw.lower()
            
            for item in items:
                if item.Class == 43:
                    s_name = str(item.SenderName).lower()
                    subj = str(item.Subject).lower()
                    if kw in s_name or kw in subj:
                        target_it = item
                        break
            
            if target_it:
                # 处理原名（去除非法字符并限制长度）
                clean_subj = re.sub(r'[\\/:*?"<>|]', '_', target_it.Subject).strip()
                clean_subj = clean_subj[:45]
                file_name = f"{clean_subj}.mht"
                full_path = os.path.join(self.share_dir, file_name)

                # --- 1. 去重检查 ---
                if os.path.exists(full_path):
                    self.add_log(f"ℹ️ 邮件已同步，跳过: {target_it.Subject}")
                else:
                    self.add_log(f"🆕 发现新邮件，正在生存 MHT 文件...")
                    self.call_ps_saver(clean_subj)
            else:
                self.add_log(f"❓ 未在收件箱发现包含 '{self.target_kw}' 的邮件。")
            
            self.update_web_index()
        except Exception as e:
            self.add_log(f"❌ Outlook 连接异常: {e}")
        finally:
            pythoncom.CoUninitialize()

    def call_ps_saver(self, safe_name):
        """调用 PowerShell 进行物理生存文件，修复路径转义漏洞"""
        # 关键：对路径进行双重转义，确保网络路径 \\10.1... 在 PS 中正常
        ps_dir = os.path.abspath(self.share_dir).replace('\\', '\\\\')
        kw = self.target_kw
        
        ps_script = f"""
        $ol = New-Object -ComObject Outlook.Application
        $ns = $ol.GetNamespace("MAPI")
        $it = $ns.GetDefaultFolder(6).Items | Where-Object {{ $_.Subject -match "{kw}" -or $_.SenderName -match "{kw}" }} | Sort-Object ReceivedTime -Descending | Select-Object -First 1
        if($it){{
            $p = Join-Path "{ps_dir}" "{safe_name}.mht"
            $it.SaveAs($p, 10)
            # 物理生存检测循环：直到文件出现且大小不为0才退出
            $timeout = 0
            while (!(Test-Path $p) -or (Get-Item $p).Length -eq 0) {{
                Start-Sleep -Milliseconds 500
                $timeout++; if ($timeout -gt 10) {{ break }}
            }}
        }}
        """
        try:
            # 隐藏黑窗口运行
            subprocess.run(["powershell", "-Command", ps_script], capture_output=True, creationflags=0x08000000)
            self.add_log("✅ MHT 文件生存成功。")
        except Exception as e:
            self.add_log(f"❌ Shell 调用失败: {e}")

    def update_web_index(self):
        """更新共享盘 index.html"""
        if not os.path.exists(self.share_dir): return
        files = [f for f in os.listdir(self.share_dir) if f.endswith('.mht')]
        files.sort(key=lambda x: os.path.getmtime(os.path.join(self.share_dir, x)), reverse=True)
        
        html = f"""
        <!DOCTYPE html>
        <html>
        <head><meta charset="utf-8"><meta http-equiv="refresh" content="60"><title>邮件看板</title>
        <style>
            body {{ font-family: 'Segoe UI', sans-serif; background: #f3f2f1; padding: 20px; }}
            .card {{ max-width: 800px; margin: 0 auto; background: white; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); overflow: hidden; }}
            .header {{ background: #0078d4; color: white; padding: 20px; text-align: center; }}
            .item {{ display: flex; align-items: center; padding: 15px; border-bottom: 1px solid #eee; text-decoration: none; color: #333; }}
            .item:hover {{ background: #f9f9f9; padding-left: 25px; transition: 0.3s; }}
            .tag {{ background: #dff6dd; color: #107c10; padding: 2px 8px; border-radius: 10px; font-size: 11px; margin-left: 10px; font-weight: bold; }}
            .footer {{ background: #fafafa; text-align: center; padding: 15px; color: #888; font-size: 12px; }}
        </style>
        </head><body>
            <div class="card">
                <div class="header"><h2 style="margin:0;">📫 RD Team 自动监控看板</h2><small>监控关键词: {self.target_kw}</small></div>
                <div style="padding: 10px;">
        """
        for i, f in enumerate(files):
            mtime = time.strftime("%Y-%m-%d %H:%M", time.localtime(os.path.getmtime(os.path.join(self.share_dir, f))))
            tag = '<span class="tag">NEW</span>' if i == 0 else ""
            html += f'<a href="{f}" target="_blank" class="item">📄 <b style="margin-left:10px; flex:1;">{f.replace(".mht","")}</b> {tag} <small style="color:#999;">{mtime}</small></a>'
        
        html += f"""
                </div>
                <div class="footer">{self.copyright_text}</div>
            </div>
        </body></html>
        """
        with open(os.path.join(self.share_dir, "index.html"), "w", encoding="utf-8") as f:
            f.write(html)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    QApplication.setQuitOnLastWindowClosed(False)
    m = OutlookMasterAuto()
    m.show()
    sys.exit(app.exec_())
