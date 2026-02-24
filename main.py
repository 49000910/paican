    def process_web(self):
        d = self.ui_path.text().strip()
        if not os.path.exists(d): return
        
        # 1. 邮件解析 (MHT -> HTML)
        for f in [x for x in os.listdir(d) if x.endswith('.mht')]:
            p_m, p_h = os.path.join(d, f), os.path.join(d, f.replace('.mht', '.html'))
            try:
                with open(p_m, 'rb') as fp:
                    msg = email.message_from_binary_file(fp)
                    for part in msg.walk():
                        if part.get_content_type() == "text/html":
                            raw = part.get_payload(decode=True).decode('utf-8','ignore')
                            clean = re.sub(r'width[:=]["\']?\d+(px|pt|in|cm)?["\']?', '', raw, flags=re.I)
                            with open(p_h, 'w', encoding='utf-8') as hw: hw.write(clean)
                os.remove(p_m)
            except: pass

        # 2. Excel 解析 (原生 Table)
        cal_html = ""
        for f_name in os.listdir(d):
            if "2026日历" in f_name and f_name.lower().endswith('.xlsx'):
                try:
                    wb = openpyxl.load_workbook(os.path.join(d, f_name), data_only=True)
                    ws = wb.active
                    rows_data = ""
                    for row in ws.iter_rows():
                        tr = ""
                        for cell in row:
                            val = "" if cell.value is None else str(cell.value)
                            bg = f"#{cell.fill.start_color.rgb[2:]}" if cell.fill and hasattr(cell.fill.start_color, 'rgb') and len(str(cell.fill.start_color.rgb))>2 else "white"
                            ft = f"#{cell.font.color.rgb[-6:]}" if cell.font and cell.font.color and hasattr(cell.font.color, 'rgb') else "black"
                            tr += f'<td style="background:{bg};color:{ft};">{val}</td>'
                        rows_data += f"<tr>{tr}</tr>"
                    cal_html = f'<table class="excel-table">{rows_data}</table>'
                except: pass
        self.build_index(cal_html)

    def build_index(self, cal_html):
        d, c = self.ui_path.text().strip(), self.ui_color.text().strip() or "#107c10"
        t1, t2, cp = self.ui_title.text().strip(), self.ui_subtitle.text().strip(), self.ui_copy.text().strip()
        all_files = [f for f in os.listdir(d) if f.endswith('.html') and f not in ['index.html', 'data.js']]
        all_files.sort(key=lambda x: os.path.getmtime(os.path.join(d, x)), reverse=True)
        try: w_ref = int(self.ui_web_freq.text())
        except: w_ref = 60

        items_html, mails_content_html, regex_ptr = "", "", self.ui_regex.text().strip()

        # 遍历所有 HTML 邮件生成列表数据
        for i, f in enumerate(all_files):
            file_path = os.path.join(d, f)
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as tf: content = tf.read()
            tags = " ".join(list(set(re.findall(regex_ptr, content))))
            # items_html 存储侧边栏列表内容
            items_html += f'<div class="mail-item" data-id="{i}" onclick="showMail(\'{i}\', this)" data-tags="{tags}"><b>{f[:-5]}</b></div>'
            # mails_content_html 存储右侧邮件正文内容
            mails_content_html += f'<div id="mail-{i}" class="mail-body" style="display:none"><div class="mail-inner-zoom">{content}</div></div>'

        # --- 步骤 1: 写入数据拆分文件 data.js ---
        # 使用 json.dumps 确保 HTML 字符串在 JS 中不报错
        js_data = f"""
        var latest_list = {json.dumps(items_html)};
        var latest_mails = {json.dumps(mails_content_html)};
        var latest_cal = {json.dumps(cal_html)};
        """
        with open(os.path.join(d, "data.js"), 'w', encoding='utf-8') as jf:
            jf.write(js_data)

        # --- 步骤 2: 写入主静态框架 index.html ---
        full_html = f"""
        <!DOCTYPE html>
        <html><head><meta charset="utf-8"><title>{t1}</title>
        <style>
            body {{ font-family: 'Segoe UI', 'Microsoft YaHei', sans-serif; margin: 0; display: flex; height: 100vh; overflow: hidden; background:#f3f2f1; }}
            .sidebar {{ width: 340px; background: white; border-right: 1px solid #edebe9; display: flex; flex-direction: column; flex-shrink: 0; height: 100vh; }}
            .header {{ padding: 20px 16px; background: {c}; color: white; flex-shrink: 0; }}
            .search-box {{ padding: 12px 16px; background: #fff; border-bottom: 1px solid #f3f2f1; position: relative; }}
            .search-box input {{ width: 100%; padding: 8px 10px; border: 1px solid #ddd; border-radius: 4px; outline: none; box-sizing: border-box; }}
            .mail-list {{ flex: 1; overflow-y: auto; }}
            .mail-item {{ padding: 14px 16px; border-bottom: 1px solid #f3f2f1; cursor: pointer; transition: 0.2s; }}
            .mail-item.active {{ border-left: 5px solid {c}; background: #eff6ef; }}
            .content {{ flex: 1; display: flex; flex-direction: column; min-width: 0; background: white; }}
            .mail-display {{ flex: 1; overflow: auto; background: #f8f9fa; }}
            .mail-inner-zoom {{ padding: 25px; zoom: 0.9; background: white; margin: 15px auto; width: 95%; box-shadow: 0 2px 15px rgba(0,0,0,0.05); }}
            .footer {{ font-size: 11px; color: #888; padding: 10px 16px; background: #fdfdfd; border-top: 1px solid #edebe9; display: flex; justify-content: space-between; align-items: center; }}
            .cal-trigger {{ cursor: pointer; color: {c}; font-weight: bold; text-decoration: underline; }}
            .modal {{ display: none; position: fixed; z-index: 9999; left: 0; top: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.6); backdrop-filter: blur(5px); }}
            .modal-content {{ background: white; margin: 2vh auto; width: 95%; height: 92%; border-radius: 8px; display: flex; flex-direction: column; overflow: hidden; }}
            .modal-header {{ padding: 15px 20px; background: {c}; color: white; display: flex; justify-content: space-between; align-items:center; }}
            .modal-body {{ flex: 1; overflow: auto; padding: 20px; background: #999; display:flex; justify-content:center; }}
            .excel-table {{ border-collapse: collapse; background: white; zoom: 0.85; }}
            .excel-table td {{ border: 1px solid #d4d4d4; padding: 4px 8px; text-align: center; white-space: nowrap; }}
        </style>
        <script id="data_script" src="data.js"></script>
        </head>
        <body>
            <div class="sidebar">
                <div class="header"><b>{t1}</b><br><small>{t2}</small></div>
                <div class="search-box"><input type="text" id="s" placeholder="搜索任务令..." onkeyup="flt()"></div>
                <div class="mail-list" id="ml"></div>
                <div class="footer">
                    <div>{cp}<br><small><span id="timer_info">正在加载数据...</span></small></div>
                    <span class="cal-trigger" onclick="tgl(true)">📅 工作日历</span>
                </div>
            </div>
            <div class="content"><div class="mail-display" id="mailDisplay"></div></div>
            <div id="mdl" class="modal"><div class="modal-content"><div class="modal-header"><h3>📅 工作日历预览</h3><span style="cursor:pointer; font-size:24px;" onclick="tgl(false)">&times;</span></div><div id="cb" class="modal-body"></div></div></div>

            <script>
                var waitTime = {w_ref};
                var counter = waitTime;
                var currentActiveTitle = "";

                function updateUI() {{
                    // 记录当前选中的邮件标题
                    let active = document.querySelector('.mail-item.active');
                    if(active) currentActiveTitle = active.innerText;

                    document.getElementById('ml').innerHTML = latest_list;
                    document.getElementById('cb').innerHTML = latest_cal;
                    
                    // 动态注入邮件正文，已存在的 ID 不覆盖（防止正文滚动条跳动）
                    let display = document.getElementById('mailDisplay');
                    let temp = document.createElement('div');
                    temp.innerHTML = latest_mails;
                    Array.from(temp.children).forEach(m => {{
                        if(!document.getElementById(m.id)) display.appendChild(m);
                    }});

                    // 恢复之前的选中状态
                    if(currentActiveTitle) {{
                        Array.from(document.querySelectorAll('.mail-item')).forEach(item => {{
                            if(item.innerText === currentActiveTitle) item.classList.add('active');
                        }});
                    }}
                    flt();
                }}

                function sync() {{
                    setInterval(() => {{
                        counter--;
                        document.getElementById('timer_info').innerText = "下次同步: " + counter + "s";
                        if(counter <= 0) {{
                            counter = waitTime;
                            let old = document.getElementById('data_script');
                            let ns = document.createElement('script');
                            ns.id = 'data_script';
                            ns.src = 'data.js?t=' + Date.now();
                            ns.onload = updateUI;
                            old.parentNode.replaceChild(ns, old);
                        }}
                    }}, 1000);
                }}

                function showMail(id, el) {{
                    document.querySelectorAll('.mail-body').forEach(b => b.style.display = 'none');
                    document.querySelectorAll('.mail-item').forEach(i => i.classList.remove('active'));
                    document.getElementById('mail-' + id).style.display = 'block';
                    el.classList.add('active');
                }}

                function flt() {{
                    let v = document.getElementById('s').value.toUpperCase();
                    document.querySelectorAll('.mail-item').forEach(i => {{
                        let t = i.innerText.toUpperCase() + (i.getAttribute('data-tags')||"").toUpperCase();
                        i.style.display = t.includes(v) ? "" : "none";
                    }});
                }}

                function tgl(v) {{ document.getElementById('mdl').style.display = v ? 'block' : 'none'; }}
                
                window.onload = () => {{ updateUI(); sync(); }};
            </script>
        </body></html>
        """
        with open(os.path.join(d, "index.html"), 'w', encoding='utf-8') as f:
            f.write(full_html)
        self.add_log("🚀 看板静态主页 index.html 已就绪")
