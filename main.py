    def run_sync(self):
        try:
            import pythoncom
            pythoncom.CoInitialize() # 确保线程安全
            
            try:
                outlook = win32com.client.GetActiveObject("Outlook.Application")
            except:
                outlook = win32com.client.Dispatch("Outlook.Application")
            
            ns = outlook.GetNamespace("MAPI")
            # 获取主账号的所有顶级文件夹（暴力全盘扫描）
            folders = ns.Folders
            target = self.target_sender.lower().strip()
            
            self.add_log(f"🔎 开始全盘暴力扫描关键词: {target}")
            
            target_item = None
            
            def deep_search(folder_obj):
                nonlocal target_item
                if target_item: return # 已找到则停止
                
                try:
                    # 获取该文件夹下最近的50封邮件（防止过老邮件拖慢速度）
                    items = folder_obj.Items
                    items.Sort("[ReceivedTime]", True)
                    
                    check_count = 0
                    for item in items:
                        if check_count > 50: break # 每个文件夹只看最新的50封
                        check_count += 1
                        
                        if item.Class == 43: # olMail
                            # 暴力获取所有可能的发件人标识
                            s_addr = str(getattr(item, 'SenderEmailAddress', '')).lower()
                            s_name = str(getattr(item, 'SenderName', '')).lower()
                            s_behalf = str(getattr(item, 'SentOnBehalfOfName', '')).lower()
                            
                            if target in s_addr or target in s_name or target in s_behalf:
                                target_item = item
                                return
                                
                    # 递归进入子文件夹
                    for sub in folder_obj.Folders:
                        deep_search(sub)
                        if target_item: return
                except:
                    pass # 跳过加密或无权限文件夹

            # 遍历所有挂载的数据文件/账号
            for f in folders:
                deep_search(f)
                if target_item: break

            if target_item:
                subj = target_item.Subject
                r_time = target_item.ReceivedTime.strftime("%Y%m%d_%H%M%S")
                # 修复 Python 3.9 兼容性：正则移出 f-string
                clean_subj = re.sub(r'[\\/:*?<>|]', '_', subj)[:50]
                fname = f"{r_time}_{clean_subj}.html"
                fpath = os.path.join(self.share_dir, fname)
                
                if not os.path.exists(fpath):
                    target_item.SaveAs(fpath, 4) # olHTML
                    self.add_log(f"✅ 抓取成功: {subj}")
                    
                    if target_item.Attachments.Count > 0:
                        att_p = os.path.join(self.share_dir, f"{r_time}_附件")
                        if not os.path.exists(att_p): os.makedirs(att_p)
                        for i in range(1, target_item.Attachments.Count + 1):
                            target_item.Attachments.Item(i).SaveAsFile(os.path.join(att_p, target_item.Attachments.Item(i).FileName))
                else:
                    self.add_log(f"ℹ️ 邮件已同步: {subj}")
            else:
                self.add_log(f"❌ 全盘扫描完毕，未发现包含 '{target}' 的邮件")
                
            self.generate_html_index()
        except Exception as e:
            self.add_log(f"❌ 运行异常: {str(e)}")
        finally:
            pythoncom.CoUninitialize()
