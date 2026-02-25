import sys, os, re, time, subprocess, base64, email, datetime
import openpyxl
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout,
                             QLabel, QLineEdit, QPushButton, QTextEdit,
                             QSystemTrayIcon, QMenu, QStyle)
from PyQt5.QtCore import QTimer

class OutlookMHTMaster(QWidget):
    def __init__(self):
        super().__init__()

        # ===== 默认参数 =====
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

    # ================= UI =================
    def init_ui(self):
        self.setWindowTitle("EDFA 看板后台 V55")
        self.resize(520, 900)

        layout = QVBoxLayout()

        def quick_edit(label, val, attr):
            h = QHBoxLayout()
            lb = QLabel(label)
            lb.setFixedWidth(120)
            h.addWidget(lb)
            edit = QLineEdit(str(val))
            setattr(self, attr, edit)
            h.addWidget(edit)
            layout.addLayout(h)

        quick_edit("📂 共享路径", self.share_dir, "ui_path")
        quick_edit("📧 邮件关键词", self.target_kw, "ui_kw")
        quick_edit("🔍 提取正则", self.tag_regex, "ui_regex")
        quick_edit("🚩 网页大标题", self.web_title, "ui_title")
        quick_edit("📝 网页小字备注", self.web_sub_title, "ui_subtitle")

        h1 = QHBoxLayout()
        h1.addWidget(QLabel("⏱ 同步频率(分)"))
        self.ui_freq = QLineEdit(str(self.interval_min))
        h1.addWidget(self.ui_freq)

        h1.addWidget(QLabel("🌐 网页刷新(秒)"))
        self.ui_web_freq = QLineEdit(str(self.web_refresh_sec))
        h1.addWidget(self.ui_web_freq)
        layout.addLayout(h1)

        h2 = QHBoxLayout()
        h2.addWidget(QLabel("🔢 抓取数"))
        self.ui_count = QLineEdit(str(self.sync_count))
        h2.addWidget(self.ui_count)

        h2.addWidget(QLabel("⏰ 时段"))
        self.ui_start = QLineEdit(str(self.start_hour))
        self.ui_end = QLineEdit(str(self.end_hour))
        h2.addWidget(self.ui_start)
        h2.addWidget(QLabel("-"))
        h2.addWidget(self.ui_end)
        layout.addLayout(h2)

        quick_edit("🎨 主题颜色", self.theme_color, "ui_color")
        quick_edit("🔒 版权内容", self.copyright_text, "ui_copy")

        self.btn_apply = QPushButton("🚀 立即同步并解析")
        self.btn_apply.clicked.connect(self.apply_settings)
        layout.addWidget(self.btn_apply)

        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        layout.addWidget(self.log_area)

        self.setLayout(layout)

    # ================= 托盘 =================
    def init_tray(self):
        self.tray = QSystemTrayIcon(self)
        self.tray.setIcon(self.style().standardIcon(QStyle.SP_ComputerIcon))
        menu = QMenu()
        menu.addAction("显示", self.showNormal)
        menu.addAction("退出", QApplication.instance().quit)
        self.tray.setContextMenu(menu)
        self.tray.show()

    def closeEvent(self, event):
        self.hide()
        event.ignore()

    def add_log(self, text):
        self.log_area.append(f"[{time.strftime('%H:%M:%S')}] {text}")

    # ================= 主循环 =================
    def apply_settings(self):
        self.add_log("⚙️ 重新执行同步...")
        self.run_cycle()

    def run_cycle(self):
        now_h = int(time.strftime("%H"))
        try:
            s = int(self.ui_start.text())
            e = int(self.ui_end.text())
        except:
            s, e = 9, 12

        if not (s <= now_h < e):
            self.add_log("💤 非活跃时段")
            self.sync_timer.start(30 * 60000)
            return

        self.fetch_outlook()
        self.process_files()

        try:
            f = int(self.ui_freq.text())
            self.sync_timer.start(f * 60000)
        except:
            self.sync_timer.start(600000)

    # ================= 抓 Outlook =================
    def fetch_outlook(self):
        d = self.ui_path.text().strip()
        k = self.ui_kw.text().strip()
        try:
            c = int(self.ui_count.text())
        except:
            c = 3

        ps = f"""
        try {{
            $ol = New-Object -ComObject Outlook.Application
            $items = $ol.GetNamespace("MAPI").GetDefaultFolder(6).Items |
                Where-Object {{ $_.Subject -like "*{k}*" }} |
                Sort-Object ReceivedTime -Descending |
                Select-Object -First {c}
            foreach($m in $items) {{
                $n = ($m.Subject -replace '[\\\\/:*?"<>|]', '_')
                $m.SaveAs((Join-Path "{d}" "$n.mht"), 10)
            }}
        }} catch {{ }} finally {{
            if ($ol) {{ [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ol) }}
        }}
        """

        try:
            b = base64.b64encode(ps.encode('utf-16-le')).decode()
            subprocess.run(
                ["powershell.exe", "-NoProfile", "-EncodedCommand", b],
                creationflags=subprocess.CREATE_NO_WINDOW,
                timeout=60
            )
            self.add_log("📥 邮件抓取完成")
        except Exception as e:
            self.add_log(f"ERR: {e}")

    # ================= 处理文件 =================
    def process_files(self):
        d = self.ui_path.text().strip()
        if not os.path.exists(d):
            return

        # mht → html
        for f in os.listdir(d):
            if f.endswith(".mht"):
                p = os.path.join(d, f)
                try:
                    with open(p, 'rb') as fp:
                        msg = email.message_from_binary_file(fp)
                        for part in msg.walk():
                            if part.get_content_type() == "text/html":
                                html = part.get_payload(decode=True).decode("utf-8", "ignore")
                                with open(p.replace(".mht", ".html"), "w", encoding="utf-8") as w:
                                    w.write(html)
                    os.remove(p)
                except:
                    pass

        self.build_index()

    # ================= 构建网页 =================
    def build_index(self):
        d = self.ui_path.text().strip()
        files = [f for f in os.listdir(d) if f.endswith(".html") and f != "index.html"]
        files.sort(key=lambda x: os.path.getmtime(os.path.join(d, x)), reverse=True)

        try:
            refresh_sec = int(self.ui_web_freq.text())
        except:
            refresh_sec = 60

        items_html = ""
        for i, f in enumerate(files):
            items_html += f"<div class='mail-item' onclick=\"showMail({i},this)\">{f[:-5]}</div>"

        full_html = f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>{self.ui_title.text()}</title>
<script>
var waitTime = {refresh_sec};
var counter = waitTime;

function startTimer(){{
    setInterval(function(){{
        counter--;
        if(counter<=0){{
            counter = waitTime;
            refreshList();
        }}
    }},1000);
}}

function refreshList(){{
    fetch("index.html?t="+Date.now())
    .then(r=>r.text())
    .then(html=>{{
        let parser=new DOMParser();
        let doc=parser.parseFromString(html,"text/html");
        let newList=doc.querySelector(".mail-list").innerHTML;
        document.querySelector(".mail-list").innerHTML=newList;
    }});
}}

function showMail(id,el){{
    let bodies=document.querySelectorAll(".mail-body");
    bodies.forEach(b=>b.style.display="none");
    document.getElementById("mail-"+id).style.display="block";
}}

window.onload=startTimer;
</script>
<style>
body{{display:flex;height:100vh;margin:0}}
.sidebar{{width:300px;background:#fff;border-right:1px solid #ccc;overflow:auto}}
.mail-item{{padding:10px;border-bottom:1px solid #eee;cursor:pointer}}
.mail-item:hover{{background:#f3f3f3}}
.content{{flex:1;overflow:auto;padding:20px}}
</style>
</head>
<body>
<div class="sidebar">
<div class="mail-list">
{items_html}
</div>
</div>
<div class="content">
{"".join([f"<div id='mail-{i}' class='mail-body' style='display:{'block' if i==0 else 'none'}'><iframe src='{f}' width='100%' height='800px'></iframe></div>" for i,f in enumerate(files)])}
</div>
</body>
</html>
"""
        with open(os.path.join(d, "index.html"), "w", encoding="utf-8") as f:
            f.write(full_html)

        self.add_log("🌐 index.html 更新完成")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = OutlookMHTMaster()
    w.show()
    sys.exit(app.exec_())
