"""
掛單分頁（見 docs/介面規劃.md 第十節）：把今天送出去的委託整批查回來攤在
一張表上。這一步只讀，取消掛單那三顆按鈕在 9.7 第 3 步——那支 API 還沒偵察過。

**為什麼這頁現在就該有**：自動送出委託單已經上線了（order_auto_confirm），
而人在程式裡沒有任何地方能確認「到底送出去了什麼」。這頁就是那個缺口，也是
自動送出唯一的驗證面。所以它顯示的是今天**所有**的委託，不只還掛著的那幾筆
——成交了、被取消了的也要看得到。

跟下單分頁刻意分開：那邊的骨架是「選股票 → 設定 → 勾帳戶 → 執行預覽 → 依序
送出」，這邊沒有股票要選、沒有預覽要看，是對「已經送出去的委託」做事，形狀
完全不同（9 節開頭那張表）。
"""

import datetime
from tkinter import messagebox

import fetch as fetch_mod
import order_query
from ui_common import col_width, wide
from util import show

# 範圍選單的「全部」。掛單這一頁預設就是全部——看掛單本來就是要一次看完所有
# 帳戶（10.1），跟同步分頁那個「通常只更新一位」的預設剛好相反。
PENDING_ALL = "全部帳戶"


class UiPendingMixin:
    # ---------- 掛單分頁 ----------

    def _pending_init_state(self):
        """SyncApp.__init__ 呼叫一次。"""
        self.pending_rows = []        # 上一次查回來的委託（order_query.normalize 的形狀）
        self.pending_busy = False     # 查詢中，還沒回話
        self.pending_scope = None     # 範圍選單的變數，_build_pending_tab 裡建
        self.pending_at = None        # 最後一次查詢完成的時間，畫在右上角

    def _pending_targets(self):
        """
        這次要查哪幾組帳號。

        名單只能從 self.trader_of 來——那是「登入過才知道名字」的既有限制（見
        ui.py），跟同步分頁的範圍選單、下單分頁的帳戶清單同一個限制，不是這裡
        另外加的。
        """
        who = self.pending_scope.get() if self.pending_scope is not None else PENDING_ALL
        # 對得到帳號設定的才算。trader_of 的鍵本來就是從 self.accounts 長出來的，
        # 照理不會超出範圍，但這裡寧可漏掉一個也不要整個分頁掛在 IndexError 上。
        pairs = [(order, self.accounts[order - 1]) for order in sorted(self.trader_of)
                 if 1 <= order <= len(self.accounts)]
        if who and who != PENDING_ALL:
            pairs = [(order, account) for order, account in pairs if self.trader_of[order] == who]
        return pairs

    def _refresh_pending_scope(self):
        """範圍選單的選項跟著已知的帳戶名單走。切到這個分頁時刷新一次就夠。"""
        if self.pending_scope is None:
            return
        names = [PENDING_ALL] + sorted(set(self.trader_of.values()))
        self.pending_choice.configure(values=names)
        if self.pending_scope.get() not in names:
            self.pending_scope.set(PENDING_ALL)

    def refresh_pending(self):
        """
        「查詢掛單」：丟一個指令給瀏覽器背景執行緒，結果回來見 _on_pending_fetched。

        跟下單分頁的「查詢委買賣」「開始下單」共用同一顆 self.busy／同一條背景
        執行緒，理由一樣：這一步也要登入／換 cookie，不能跟同步分頁或下單依序
        執行同時搶同一顆瀏覽器。
        """
        if self.busy or self.pending_busy:
            return

        targets = self._pending_targets()
        if not targets:
            messagebox.showinfo(
                "還沒有帳戶名字",
                "還沒有任何帳戶登入過，名字都還不知道。\n請先到「同步」分頁按「登入」。",
                parent=self.root)
            return

        self.pending_busy = True
        self._set_busy(True, f"查詢掛單中（{len(targets)} 個帳戶）…")
        self._update_pending_ui()
        self._ensure_browser_thread()
        self.browser_waiting += 1
        self.browser_cmd_queue.put(("pending", (targets,)))

    def _pending_job(self, context, store, targets):
        """
        背景執行緒用（只能在 ui_background._browser_worker 裡呼叫）。

        **一組登入完就立刻查那一組**，不是先把全部登入完再回頭查——整個瀏覽器
        只有一組 cookie，全部登入完之後它是最後一組的（見 fetch.ensure_logged_in
        的警告與 fetch.collect 開頭那段）。所以 ensure_logged_in 是在迴圈裡一次
        只帶一組進去呼叫的，跟 ui_order._order_quotes_job 同一個寫法。

        某一組失敗不中斷整批：查掛單是「看看現在外面有什麼」，20 個裡有 1 個
        登入逾時，另外 19 個的委託還是該看得到。失敗的收進 problems 一起回報。
        """
        rows, problems = [], []
        for order, account in targets:
            name = self.trader_of.get(order, f"第 {order} 組")
            try:
                page, session, probs = fetch_mod.ensure_logged_in(
                    context, [(order, account)], store)[order]
                if probs:
                    problems.append(f"{name}：{'；'.join(probs)}")
                    continue
                sheet = (session.get("account") or "").strip() or name
                rows.extend(order_query.query_orders(page, session, sheet))
            except RuntimeError as exc:
                problems.append(str(exc))
        return {"rows": rows, "problems": problems}

    def _on_pending_fetched(self, payload):
        """背景查詢回話。整份換掉，不是併進去——這一頁講的是「現在外面有什麼」。"""
        # 這則回話是從瀏覽器背景執行緒來的，跟「查詢委買賣」一樣要把等待計數減
        # 回去（見 ui_background._check_browser_thread：那個數字減不回去的話，
        # 執行緒真的掛掉時畫面會永遠停在等回話的狀態）。
        self.browser_waiting = max(0, self.browser_waiting - 1)
        self.pending_busy = False
        self._set_busy(False)

        if "error" in payload:
            self._update_pending_ui()
            messagebox.showerror("查詢掛單失敗", payload["error"], parent=self.root)
            return

        self.pending_rows = payload.get("rows", [])
        self.pending_at = datetime.datetime.now()
        self._fill_pending()

        problems = payload.get("problems") or []
        open_count = sum(1 for row in self.pending_rows if row["open"])
        self._say(f"掛單：{len(self.pending_rows)} 筆委託，其中 {open_count} 筆還掛在外面。"
                  + (f"　{len(problems)} 個帳戶查不到。" if problems else ""))
        if problems:
            messagebox.showwarning("有帳戶查不到",
                                   "這幾個帳戶沒查到委託：\n\n" + "\n".join(problems),
                                   parent=self.root)

    def _fill_pending(self):
        """整份重建，跟同步分頁 fill_sync_tree()／下單分頁執行預覽同一個做法。"""
        for item in self.pending_tree.get_children():
            self.pending_tree.delete(item)

        for row in self.pending_rows:
            # 還掛在外面的用買賣底色（跟下單分頁的執行預覽同一組顏色，也跟網站
            # 本身買紅賣綠一致）；已經成交／取消／失敗的淡化——它們是「已經結束
            # 的事」，不該看起來跟還能取消的那幾筆一樣醒目。
            tag = {"B": "buy", "S": "sell"}.get(row["side"], "") if row["open"] else "done"
            self.pending_tree.insert("", "end", values=(
                row["sheet"], f"{row['code']}", row["side_text"],
                show(row["price"]) if row["price"] is not None else "",
                show(row["qty"]), show(row["matched"]), show(row["left"]),
                row["status"], row["ordno"],
            ), tags=(tag,) if tag else ())

        self._resize_pending_columns()
        self._update_pending_ui()

    def _resize_pending_columns(self):
        """「帳戶」欄寬跟著這次查到的名字量一次，理由同下單分頁的同名做法。"""
        names = sorted({row["sheet"] for row in self.pending_rows})
        self.pending_tree.column("sheet", width=col_width(self.family, names, minimum=wide(90)))

    def _update_pending_ui(self):
        """
        查詢按鈕與右上角的時間戳。三顆取消按鈕不在這裡管——它們在 9.7 第 3 步
        之前一律是 disabled（見 ui_layout._build_pending_tab），沒有任何狀態
        會讓它們亮起來。
        """
        if self.pending_busy:
            self.pending_button.configure(text="查詢中…", state="disabled")
        else:
            self.pending_button.configure(
                text="查詢掛單", state="disabled" if self.busy else "normal")

        if self.pending_at is None:
            self.pending_stamp.configure(text="還沒查過")
        else:
            self.pending_stamp.configure(text=f"最後查詢 {self.pending_at:%H:%M:%S}")
