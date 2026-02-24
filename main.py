    def build_index(self, cal_html):
        d, c = self.ui_path.text().strip(), self.ui_color.text().strip() or "#107c10"
        t1, t2, cp = self.ui_title.text().strip(), self.ui_subtitle.text().strip(), self.ui_copy.text().strip()
        all_files = [f for f in os.listdir(d) if f.endswith('.html') and f != 'index.html']
        all_files.sort(key=lambda x: os.path.getmtime(os.path.join(d, x)), reverse=True)
        try: w_ref = int(self.ui_web_freq.text())
        except: w_ref = 60
        update_time = time.strftime('%Y-%m-%d %H:%M:%S')
        items_html, mails_content_html, regex_ptr = "", "", self.ui_regex.text().strip()

        for i, f in enumerate(all_files):
            file_path = os.path.join(d, f)
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as tf: content = tf.read()
            tags = " ".join(list(set(re.findall(regex_ptr, content))))
            items_html += f'<div class="mail-item {"active" if i==0 else ""}" onclick="showMail(\'{i}\', this)" data-tags="{tags}"><b>{f[:-5]}</b></div>'
            mails_content_html += f'<div id="mail-{i}" class="mail-body" style="display:{"block" if i==0 else "none"}"><div class="mail-inner-zoom">{content}</div></div>'

        full_html = f"""
        <!DOCTYPE html>
        <html><head><meta charset="utf-8"><title>{t1}</title>
        <style>
            body {{ font-family: 'Segoe UI', 'Microsoft YaHei', sans-serif; margin: 0; display: flex; height: 100vh; overflow: hidden; background:#f3f2f1; }}
            .sidebar {{ width: 340px; background: white; border-right: 1px solid #edebe9; display: flex; flex-direction: column; flex-shrink: 0; height: 100vh; }}
            .header {{ padding: 20px 16px; background: {c}; color: white; flex-shrink: 0; }}
            .search-box {{ padding: 12px 16px; background: #fff; border-bottom: 1px solid #f3f2f1; position: relative; }}
            .search-box input {{ width: 100%; padding: 8px 10px; border: 1px solid #ddd; border-radius: 4px; outline: none; box-sizing: border-box; }}
            .clear-btn {{ position: absolute; right: 26px; top: 50%; transform: translateY(-50%); cursor: pointer; color: #bbb; display: none; font-size: 20px; }}
            .mail-list {{ flex: 1; overflow-y: auto; }}
            .mail-item {{ padding: 14px 16px; border-bottom: 1px solid #f3f2f1; cursor: pointer; }}
            .mail-item.search-hit {{ background-color: #fff9c4 !important; border-left: 5px solid #fbc02d !important; }}
            .mail-item.active {{ border-left: 5px solid {c}; background: #eff6ef; }}
            .content {{ flex: 1; display: flex; flex-direction: column; min-width: 0; background: white; }}
            .mail-display {{ flex: 1; overflow: auto; background: #f8f9fa; }}
            .mail-inner-zoom {{ padding: 25px; zoom: 0.9; background: white; margin: 15px auto; width: 95%; box-shadow: 0 2px 15px rgba(0,0,0,0.05); }}
            .footer {{ font-size: 11px; color: #888; padding: 10px 16px; background: #fdfdfd; border-top: 1px solid #f3f2f1; display: flex; justify-content: space-between; align-items: center; }}
            .cal-trigger {{ cursor: pointer; color: {c}; font-weight: bold; text-decoration: underline; }}
            .modal {{ display: none; position: fixed; z-index: 9999; left: 0; top: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.6); backdrop-filter: blur(5px); }}
            .modal-content {{ background: #999; margin: 1vh auto; width: 98%; height: 96%; border-radius: 8px; display: flex; flex-direction: column; overflow: hidden; }}
            .modal-header {{ padding: 12px 20px; background: {c}; color: white; display: flex; justify-content: space-between; align-items:center; }}
            .modal-body {{ flex: 1; overflow: auto; padding: 20px; display: flex; justify-content: center; }}
            .excel-table {{ border-collapse: collapse; background: white; zoom: 0.8; }}
            .excel-table td {{ padding: 4px 8px; min-width: 40px; height: 25px; text-align: center; white-space: nowrap; font-size: 13px; border: 1px solid #d4d4d4; }}
            mark {{ background: #ffeb3b; color: #000; font-weight: bold; padding: 0 2px; }}
            #sync_dot {{ width: 8px; height: 8px; background: #4caf50; border-radius: 50%; display: inline-block; margin-right: 4px; animation: blink 2s infinite; }}
            @keyframes blink {{ 0% {{ opacity: 1; }} 50% {{ opacity: 0.3; }} 100% {{ opacity: 1; }} }}
        </style></head>
        <body>
            <div class="sidebar">
                <div class="header">
                    <div style="font-size:19px; font-weight:700;">{t1}</div>
                    <small>{t2}</small>
                </div>
                <div class="search-box">
                    <input type="text" id="s" placeholder="搜索任务令/日期..." onkeyup="flt()">
                    <span id="cb" class="clear-btn" onclick="cls()">×</span>
                </div>
                <div class="mail-list" id="ml">{items_html}</div>
                <div class="footer">
                    <div>{cp}<br><small><span id="sync_dot"></span><span id="timer_info">同步中...</span></small></div>
                    <span class="cal-trigger" onclick="tgl(true)">📅 工作日历</span>
                </div>
            </div>
            <div class="content"><div class="mail-display" id="mailDisplay">{mails_content_html}</div></div>
            <div id="mdl" class="modal">
                <div class="modal-content">
                    <div class="modal-header"><h3>📅 工作日历 (原生预览)</h3><span style="cursor:pointer; font-size:35px;" onclick="tgl(false)">&times;</span></div>
                    <div class="modal-body">{cal_html}</div>
                </div>
            </div>
            <!-- 🚀 隐藏的 Iframe 刷新器 -->
            <iframe id="refresh_frame" src="about:blank" style="display:none;"></iframe>
            <script>
                var ori = {{}};
                var waitTime = {w_ref};
                var counter = waitTime;

                window.onload = function() {{ 
                    initBackup();
                    startCountdown();
                }};

                function initBackup() {{
                    document.querySelectorAll('.mail-body').forEach(b => {{
                        let id = b.id.replace('mail-','');
                        if(!ori[id]) ori[id] = b.innerHTML;
                    }});
                }}

                function startCountdown() {{
                    setInterval(function() {{
                        counter--;
                        document.getElementById('timer_info').innerText = "下次更新: " + counter + "s";
                        if(counter <= 0) {{
                            counter = waitTime;
                            iframeRefresh();
                        }}
                    }}, 1000);
                }}

                function iframeRefresh() {{
                    let frame = document.getElementById('refresh_frame');
                    // 强制刷新 iframe
                    frame.src = "index.html?t=" + Date.now();
                    
                    frame.onload = function() {{
                        try {{
                            let newDoc = frame.contentDocument || frame.contentWindow.document;
                            let newList = newDoc.getElementById('ml').innerHTML;
                            let oldList = document.getElementById('ml');
                            if(newList && oldList.innerHTML !== newList) {{
                                oldList.innerHTML = newList;
                                flt(true); // 保持搜索状态
                                console.log("Iframe 同步成功");
                            }}
                        }} catch(e) {{
                            console.log("Iframe 跨域限制，尝试备用重载...");
                            // 如果 iframe 也被拦截，最后手段：全页刷新
                            // window.location.reload(); 
                        }}
                    }};
                }}

                function showMail(id, el) {{
                    document.querySelectorAll('.mail-body').forEach(b => b.style.display = 'none');
                    document.querySelectorAll('.mail-item').forEach(i => i.classList.remove('active'));
                    document.getElementById('mail-'+id).style.display = 'block';
                    el.classList.add('active'); 
                    flt(true);
                }}

                function flt(r) {{
                    let v = document.getElementById('s').value.toUpperCase();
                    document.getElementById('cb').style.display = v ? 'block' : 'none';
                    document.querySelectorAll('.mail-item').forEach(item => {{
                        let txt = (item.innerText + (item.getAttribute('data-tags')||"")).toUpperCase();
                        item.style.display = (v && txt.indexOf(v) == -1) ? "none" : "block";
                        item.classList.toggle('search-hit', v && txt.indexOf(v) > -1);
                    }});
                    let act = document.querySelector('.mail-body[style*="block"]');
                    if(act) {{
                        let id = act.id.replace('mail-','');
                        if(v && v.length >= 2) {{
                            if(!ori[id]) initBackup();
                            act.innerHTML = ori[id].replace(new RegExp('('+v+')','gi'), '<mark class="m">$1</mark>');
                            let m = act.querySelector('.m'); if(m && !r) m.scrollIntoView({{behavior:'smooth',block:'center'}});
                        }} else {{ if(ori[id]) act.innerHTML = ori[id]; }}
                    }}
                }}
                function cls() {{ document.getElementById('s').value=''; flt(); }}
                function tgl(s) {{ document.getElementById('mdl').style.display = s ? 'block' : 'none'; }}
            </script>
        </body></html>"""
        with open(os.path.join(d, "index.html"), 'w', encoding='utf-8') as f: f.write(full_html)
