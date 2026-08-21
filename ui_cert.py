"""憑證分頁的行為：建立/重建 Profile、掃描並遷移憑證。"""

from tkinter import messagebox

import profile_tools
from ui_common import ask_confirm


class UiCertMixin:
    # ---------- 憑證分頁 ----------

    def _refresh_profile_status(self):
        raw = profile_tools.current_raw()
        path = profile_tools.resolve_path(raw)
        if path is None:
            self.profile_status.configure(text="目前 .env 沒有設定 USER_DATA_DIR —— 憑證存不住。")
        elif not path.is_dir():
            self.profile_status.configure(text=f"目前的 Profile：{path}（還沒建立）")
        elif profile_tools.has_cert(path):
            self.profile_status.configure(text=f"目前的 Profile：{path}（已經有 tbbstock 憑證）")
        else:
            self.profile_status.configure(text=f"目前的 Profile：{path}（還沒有 tbbstock 憑證）")

    def create_profile(self):
        """
        建立（或清空重建）自動登入用的 Chrome 使用者資料夾：開一個一般模式的 Chrome
        把資料夾初始化出來，確認初始化完成就自動關掉視窗，不必像 setup-profile.ps1
        那樣手動把視窗關掉。對應 2.1～2.3。
        """
        if self.profile_busy:
            return

        raw = self.profile_name.get().strip()
        path = profile_tools.resolve_path(raw)
        if path is None:
            messagebox.showerror("名稱不能是空的", "請輸入 Profile 名稱。", parent=self.root)
            return
        if profile_tools.is_default_chrome_dir(path):
            messagebox.showerror(
                "不能用這個資料夾",
                "不能指向 Chrome 的預設使用者資料夾，Chrome 136 之後禁止自動化連上它。\n"
                "請換一個名稱（例如 chrome-profile）。", parent=self.root)
            return

        if path.is_dir():
            if profile_tools.chrome_pids_for_profile(path):
                messagebox.showerror(
                    "資料夾正在使用中",
                    f"這個資料夾正被 Chrome 開著，請先把它關掉再試一次：\n{path}", parent=self.root)
                return
            try:
                profile_tools.delete_profile(path)
            except OSError as exc:
                messagebox.showerror("刪不掉", f"刪除資料夾失敗：\n{exc}", parent=self.root)
                return

        chrome_exe = profile_tools.find_chrome()
        if chrome_exe is None:
            messagebox.showerror("找不到 Chrome", "這台電腦找不到 Google Chrome，請先安裝。",
                                 parent=self.root)
            return

        try:
            profile_tools.remember_user_data_dir(raw)
        except OSError as exc:
            messagebox.showwarning("沒寫進 .env", f"這次可以用，但沒能寫進 .env：\n{exc}",
                                   parent=self.root)

        try:
            profile_tools.launch_manual_chrome(chrome_exe, path)
        except OSError as exc:
            messagebox.showerror("開不起來", f"沒辦法開啟 Chrome：\n{exc}", parent=self.root)
            return

        self.profile_busy = True
        self.profile_button.configure(state="disabled")
        self.profile_status.configure(text=f"正在建立 Profile：{path}（Chrome 開起來了，請稍候…）")
        self._poll_profile_init(path, 0)

    def _poll_profile_init(self, path, tries):
        """每半秒確認一次資料夾初始化完成沒，完成就自動把那個 Chrome 視窗關掉（2.3）。"""
        if profile_tools.profile_initialized(path):
            profile_tools.kill_pids(profile_tools.chrome_pids_for_profile(path))
            self.profile_busy = False
            self.profile_button.configure(state="normal")
            self._refresh_profile_status()
            messagebox.showinfo(
                "Profile 建好了",
                f"資料夾已就緒，視窗已經自動關掉：\n{path}\n\n"
                "接下來可以用下面的「遷移憑證」把憑證複製進來；掃不到的話就是還沒申請過，"
                "自己開這個資料夾登入 tbbstock 申請一次即可。", parent=self.root)
            return
        if tries >= 60:   # 30 秒
            self.profile_busy = False
            self.profile_button.configure(state="normal")
            self.profile_status.configure(text=f"還沒看到 Profile 初始化完成：{path}")
            messagebox.showwarning(
                "沒看到初始化完成",
                f"30 秒過去了，還沒看到 Chrome 把資料夾初始化好：\n{path}\n\n"
                "視窗可能還在開啟中，請自己看一下，關掉後再按一次「建立 Profile」確認結果。",
                parent=self.root)
            return
        self.root.after(500, lambda: self._poll_profile_init(path, tries + 1))

    def scan_cert_sources(self):
        """掃描這台電腦上 Chrome／Edge 的每個 profile，看誰有 tbbstock 憑證痕跡（2.4）。"""
        target = profile_tools.resolve_path(profile_tools.current_raw())
        self.migrate_tree.delete(*self.migrate_tree.get_children())
        self._migrate_candidates = {}
        self.migrate_copy_button.configure(state="disabled")

        if target is None:
            self.migrate_status.configure(text="還沒設定 USER_DATA_DIR。")
            return

        candidates = profile_tools.scan_cert_sources(target)
        # 沒憑證痕跡的 profile 只是雜訊 —— 這張表是給人挑「要從哪一個複製」，
        # 不是給人看「這台電腦裝了幾個 profile」，所以只列有痕跡的那幾個。
        found_candidates = [c for c in candidates if c["found"]]
        for candidate in found_candidates:
            item = self.migrate_tree.insert(
                "", "end",
                values=(candidate["browser"], candidate["name"], "有", str(candidate["path"])),
            )
            self._migrate_candidates[item] = candidate

        if found_candidates:
            self.migrate_status.configure(
                text=f"掃了 {len(candidates)} 個 profile，{len(found_candidates)} 個有憑證痕跡。")
        elif candidates:
            self.migrate_status.configure(text=f"掃了 {len(candidates)} 個 profile，沒有找到憑證痕跡。")
        else:
            self.migrate_status.configure(text="這台電腦上沒找到 Chrome／Edge 的 profile。")

    def _on_migrate_select(self, _event=None):
        selection = self.migrate_tree.selection()
        ok = bool(selection) and self._migrate_candidates.get(selection[0], {}).get("found")
        self.migrate_copy_button.configure(state="normal" if ok else "disabled")

    def copy_selected_cert(self):
        """把選中的來源 profile 的憑證複製到自動登入用的 Profile。"""
        selection = self.migrate_tree.selection()
        if not selection:
            return
        source = self._migrate_candidates.get(selection[0])
        if not source or not source["found"]:
            return

        target = profile_tools.resolve_path(profile_tools.current_raw())
        if target is None or not target.is_dir():
            messagebox.showerror(
                "目標 Profile 還不存在",
                "請先用上面的「建立 Profile」把資料夾建出來，再回來複製憑證。", parent=self.root)
            return

        if profile_tools.browser_running(source["exe"]):
            messagebox.showerror(
                "來源瀏覽器還開著",
                f"{source['browser']} 還在執行，複製到的可能是還沒寫進磁碟的舊資料。\n"
                f"請把 {source['browser']} 所有視窗（含背景程序）都關掉再試一次。", parent=self.root)
            return
        if profile_tools.chrome_pids_for_profile(target):
            messagebox.showerror(
                "目標 Profile 正在使用中",
                f"這個資料夾正被 Chrome 開著，請先關掉：\n{target}", parent=self.root)
            return

        # target 是 USER_DATA_DIR 本身，但 Chrome 實際把資料存在它底下的 Default（或
        # BROWSER_PROFILE_DIR）子資料夾裡 —— copy_cert 要對到那一層，不是 USER_DATA_DIR 自己。
        profile_dir = target / profile_tools.profile_subdir_name()
        if not profile_dir.is_dir():
            messagebox.showerror(
                "Profile 還沒初始化",
                f"這個資料夾裡還沒有 {profile_tools.profile_subdir_name()} 子資料夾：\n{target}\n\n"
                "請先用上面的「建立 Profile」把資料夾初始化，再回來複製憑證。", parent=self.root)
            return

        if not ask_confirm(
            self.root,
            "複製憑證",
            f"把「{source['browser']} / {source['name']}」的憑證複製到自動登入用的 Profile？\n\n"
            "目標現有的資料會先備份。來源那張憑證不受影響（這只是複製檔案，不是重新申請）。",
            danger=False,
        ):
            return

        try:
            backup = profile_tools.copy_cert(source["path"], profile_dir)
        except OSError as exc:
            messagebox.showerror("複製失敗", str(exc), parent=self.root)
            return

        self._refresh_profile_status()
        self.scan_cert_sources()
        note = f"（原本的已備份到 {backup.name}）" if backup else ""
        self.migrate_status.configure(text=f"已複製完成{note}")
        messagebox.showinfo(
            "複製完成",
            f"憑證已複製到自動登入用的 Profile{note}。\n"
            "接下來按「登入」實際驗證，不再跳「瀏覽器查無有效數位憑證」就是成功了。", parent=self.root)
