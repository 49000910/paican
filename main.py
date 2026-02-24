import sys, os, re, time, subprocess, base64, email, json
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QTextEdit, 
                             QSystemTrayIcon, QMenu, QAction, QStyle)
from PyQt5.QtCore import QTimer, Qt

class OutlookMHTMaster(QWidget):
    def __init__(self):
        super().__init__()
        # --- 基础配置 ---
        self.share_dir = r'\\10.1.93.32\DT_HU_RDteam_F\视频\Z\ZOUQIU\paican'
        self.target_kw = 'EDFA' 
        self.interval_min = 10     
        self.sync_count = 20       # 抓取最近20封包含关键词的邮件
        self.theme_color = "#107c10" 
        self.web_title = "EDFA 全量数据看板"
        
        self.init_ui()
        self.init_tray()
        self.sync_timer = QTimer(self)
        self.sync_timer.timeout.connect(self.run_cycle)
        QTimer.singleShot(1000, self.run_cycle)

    def init_ui(self):
        self.setWindowTitle("EDFA 看板后台 V30.0 - 后端脱水索引版")
        self.resize(550, 750)
        layout = QVBoxLayout()
        
        def quick_edit(label, val, attr):
            l = QHBoxLayout()
            lb = QLabel(label); lb.setFixedWidth(100); l.addWidget(lb)
            edit = QLineEdit(str(val)); setattr(self, attr, edit)
            l.addWidget(edit); layout.addLayout(l)

        quick_edit("📂 共享路径", self.share_dir, "ui_path")
        quick_edit("📧 邮件关键词", self.target_kw, "ui_kw")
        quick_edit("🎨 主题颜色", self.theme_color, "ui_color")
        
        self.btn_apply = QPushButton("🚀 立即全量脱水并更新看板")
        self.btn_apply.setFixedHeight(50); self.btn_apply.clicked.connect(self.run_cycle)
        layout.addWidget(self.btn_apply)

        self.log_area = QTextEdit(); self.log_area.setReadOnly(True)
        layout.addWidget(self.log_area); self.setLayout(layout)

    def init_tray(self):
        self.tray = QSystemTrayIcon(self)
        self.tray.setIcon(self.style().standardIcon(QStyle.SP_ComputerIcon))
        tm = QMenu(); tm.addAction("显示界面", self.showNormal); tm.addAction("退出程序", QApplication.instance().quit)
        self.tray.setContextMenu(tm); self.tray.show()

    def add_log(self, txt): self.log_area.append(f"[{time.strftime('%H:%M:%S')}] {str(txt)}")

    def run_cycle(self):
        self.run_shell()
        self.sync_timer.start(int(self.interval_min) * 60000)

    def run_shell(self):
        d, k = self.ui_path.text().replace('"', '""'), self.ui_kw.text().replace('"', '""')
        ps_cmd = f"""
        try {{
            $ol = New-Object -ComObject Outlook.Application
            $it = $ol.GetNamespace("MAPI").GetDefaultFolder(6).Items | Where-Object {{ $_.ReceivedTime -gt (Get-Date).AddDays(-15) -and ($_.Subject -like "*{k}*" -or $_.SenderName -like "*{k}*") }} | Sort-Object ReceivedTime -Descending | Select-Object -First {self.sync_count}
            foreach($m in $it) {{ 
                $n = ($m.Subject -replace '[\\x00-\\x1f\\\\/:*?"<>|]', '_').Trim(); 
                $p = Join-Path "{d}" "$n.mht"; 
                if (!(Test-Path $p)) {{ $m.SaveAs($p, 10) }} 
            }}
        }} catch {{ }}
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
                            body = part.get_payload(decode=True).decode('utf-8','ignore')
                            clean_h = f"<div class='parsed-content'>{body}</div>"
                            with open(p_h, 'w', encoding='utf-8') as hw: hw.write(clean_h)
                os.remove(p_m)
            except: pass
        self.build_index()

    def build_index(self):
        d = self.ui_path.text().strip()
        all_files = [f for f in os.listdir(d) if f.endswith('.html') and f != 'index.html']
        all_files.sort(key=lambda x: os.path.getmtime(os.path.join(d, x)), reverse=True)
        
        c = self.ui_color.text().strip() or "#107c10"
        items_html, mails_data_html = "", ""
        search_db = {} # 后端脱水索引池

        for i, f in enumerate(all_files[:150]):
            p = os.path.join(d, f)
            with open(p, 'r', encoding='utf-8', errors='ignore') as fc: raw_h = fc.read()
            
            # --- 🔥 后端脱水 (Python 处理) ---
            # 移除标签、样式、脚本，合并为空白，转小写
            text_only = re.sub(r'<(style|script)[^>]*>.*?<\/\1>', '', raw_h, flags=re.DOTALL|re.IGNORECASE)
            text_only = re.sub(r'<[^>]+>', ' ', text_only)
            pure_text = " ".join(text_only.split()).lower()
            
            # 存储索引 (key 为 ID)
            search_db[f"m_{i}"] = pure_text

            # 提取 11位 EP 号标签
            ep_tags = list(set(re.findall(r'\bEP[A-Z0-9]{9}\b', pure_text, re.I)))
            tag_ui = "".join([f'<span class="ep-tag" onclick="fastGo(\'{x}\')">{x}</span>' for x in ep_tags[:5]])

            items_html += f'''
            <div class="list-item" id="list_m_{i}" onclick="jumpTo('m_{i}', this)">
                <div class="item-name">{f[:-5]}</div>
                <div class="item-tags">{tag_ui}</div>
            </div>'''
            
            mails_data_html += f'''
            <div id="m_{i}" class="mail-box">
                <div class="mail-bar" style="border-left:5px solid {c}">{f[:-5]}</div>
                <div class="mail-body">{raw_h}</div>
            </div>'''

        db_json = base64.b64encode(json.dumps(search_db).encode('utf-8')).decode('ascii')

        index_tpl = f'''
        <!DOCTYPE html><html><head><meta charset="UTF-8"><style>
            body {{ display:flex; height:100vh; margin:0; font-family: sans-serif; background:#f4f4f4; }}
            #left {{ width:380px; background:#fff; border-right:1px solid #ddd; display:flex; flex-direction:column; }}
            #right {{ flex:1; overflow-y:auto; padding:20px; scroll-behavior:smooth; }}
            .s-box {{ padding:15px; background:{c}; }}
            #q {{ width:100%; padding:10px; border:none; border-radius:4px; outline:none; }}
            .list-item {{ padding:12px; border-bottom:1px solid #eee; cursor:pointer; font-size:13px; }}
            .list-item:hover {{ background:#f9f9f9; }}
            .ep-tag {{ background:#e8f5e9; color:{c}; padding:1px 5px; border-radius:3px; margin-right:5px; font-size:10px; border:1px solid {c}; }}
            .mail-box {{ background:#fff; margin-bottom:30px; border-radius:4px; box-shadow:0 1px 4px rgba(0,0,0,0.1); }}
            .mail-bar {{ padding:12px; background:#fafafa; font-weight:bold; font-size:15px; }}
            .mail-body {{ padding:15px; font-size:14px; overflow-x:auto; }}
            mark {{ background: yellow; color: black; font-weight:bold; }}
            .active {{ background:#e8f5e9 !important; border-right:5px solid {c}; }}
        </style></head>
        <body>
            <div id="left">
                <div class="s-box"><input type="text" id="q" placeholder="输入 EP号 或 关键词定位..." oninput="doSearch(this.value)"></div>
                <div style="overflow-y:auto; flex:1;">{items_html}</div>
            </div>
            <div id="right">{mails_data_html}</div>
            <script>
                const db = JSON.parse(atob("{db_json}"));
                function fastGo(v) {{ document.getElementById('q').value = v; doSearch(v); }}
                
                function doSearch(kw) {{
                    const val = kw.toLowerCase().trim();
                    let first = null;
                    Object.keys(db).forEach(id => {{
                        const match = db[id].includes(val);
                        document.getElementById('list_'+id).style.display = match ? 'block' : 'none';
                        if(match && !first) first = id;
                    }});
                    if(first && val.length > 5) jumpTo(first, document.getElementById('list_'+first));
                }}

                function jumpTo(id, el) {{
                    const target = document.getElementById(id);
                    const kw = document.getElementById('q').value.trim();
                    target.scrollIntoView({{ behavior: 'smooth', block: 'start' }});
                    
                    // 高亮
                    document.querySelectorAll('.mail-body').forEach(b => b.innerHTML = b.innerHTML.replace(/<mark>|<\/mark>/g, ""));
                    if(kw.length > 2) {{
                        const body = target.querySelector('.mail-body');
                        const reg = new RegExp("("+kw+")", "gi");
                        body.innerHTML = body.innerHTML.replace(reg, "<mark>$1</mark>");
                    }}
                    document.querySelectorAll('.list-item').forEach(i => i.classList.remove('active'));
                    el.classList.add('active');
                }}
            </script>
        </body></html>'''
        with open(os.path.join(d, 'index.html'), 'w', encoding='utf-8') as f: f.write(index_tpl)
        self.add_log(f"🌍 看板已就绪，全量索引 {len(all_files)} 封文档")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = OutlookMHTMaster(); win.show()
    sys.exit(app.exec_())
