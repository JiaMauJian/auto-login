"""
掛單分頁（見 docs/介面規劃.md 第十節）：把今天送出去的委託整批查回來攤在
一張表上，以及那三顆取消按鈕（10.3）。

取消掛單走的是**頁面**不是 API（憑證簽章在瀏覽器裡做，見 order_cancel.py），
而且是一個帳戶一則指令、不是一則指令跑完整批——這樣底部那顆「停止全部操作」
在帳戶與帳戶之間就是免費的，不必為了中斷另外發明一個跨執行緒的旗標。

**為什麼這頁現在就該有**：自動送出委託單已經上線了（order_auto_confirm），
而人在程式裡沒有任何地方能確認「到底送出去了什麼」。這頁就是那個缺口，也是
自動送出唯一的驗證面。所以它顯示的是今天**所有**的委託，不只還掛著的那幾筆
——成交了、被取消了的也要看得到。

跟下單分頁刻意分開：那邊的骨架是「選帳戶 → 選股票 → 設定 → 執行預覽 → 依序
送出」，這邊沒有股票要選、沒有預覽要看，是對「已經送出去的委託」做事，形狀
完全不同（9 節開頭那張表）。
"""

import datetime

import fetch as fetch_mod
import order_cancel
import order_cancel_reservation
import order_query
from ui_common import ask_confirm, col_width, show_error, show_info, show_warning, wide
from util import show

# 範圍選單的「全部」。掛單這一頁預設就是全部——看掛單本來就是要一次看完所有
# 帳戶（10.1），跟更新分頁那個「通常只更新一位」的預設剛好相反。
PENDING_ALL = "全部帳戶"

# 三顆取消按鈕：按鈕上的字，以及它挑哪一邊（None = 不分買賣）。買賣別用的是
# queryOrder 原始的 buysell（'B'/'S'），不是畫面上那句中文。
CANCEL_KINDS = (("buy", "取消全部買單", "B"),
                ("sell", "取消全部賣單", "S"),
                ("all", "取消全部掛單", None))
CANCEL_SIDES = {kind: side for kind, _text, side in CANCEL_KINDS}
CANCEL_TEXTS = {kind: text for kind, text, _side in CANCEL_KINDS}


def _kind_text(row):
    """委託單／預約單哪一種——兩者取消走的機制不一樣（order_cancel.py／
    order_cancel_reservation.py），表上跟確認視窗都要標出來，別讓人猜。"""
    return "預約單" if row.get("ordstatus") == "1" else "委託單"


class UiPendingMixin:
    # ---------- 掛單分頁 ----------

    def _pending_init_state(self):
        """SyncApp.__init__ 呼叫一次。"""
        self.pending_rows = []        # 上一次查回來的委託（order_query.normalize 的形狀）
        self.pending_busy = False     # 查詢中，還沒回話
        self.pending_scope = None     # 範圍選單的變數，_build_pending_tab 裡建
        self.pending_at = None        # 最後一次查詢完成的時間，畫在右上角
        # 取消掛單（10.3）。跟下單的依序執行同一個形狀：這一輪要跑的帳戶凍結成
        # 一張 queue，一個帳戶一則指令，跑完一則才派下一則。
        self.pending_cancel_active = False   # 整批還在跑（停止鈕看的就是它）
        # [(第幾組帳號, 帳號設定, 帳戶名, [委託單委託書號], [預約單預約書號])]
        self.pending_cancel_queue = []
        self.pending_cancel_pos = 0
        self.pending_cancel_label = ""       # 「取消全部買單」之類，說話用
        self.pending_cancel_results = []     # 每一筆的結果（order_cancel 的形狀）
        self.pending_cancel_problems = []

    def _pending_targets(self):
        """
        這次要查哪幾組帳號。

        **選「全部帳戶」時就是 .env 裡的每一組，不管登入過沒有**（2026/08/31
        使用者要求）：查掛單要用的是「哪幾組帳號」，而 `_pending_job` 裡的
        `fetch.ensure_logged_in` 本來就會替沒登入的那幾組登入，名字也是那時候
        才從 sessionStorage 讀回來的。拿 `self.trader_of`（＝登入過才知道名字）
        當名單，等於逼人先按一次「登入」才能按「查詢掛單」，而那一步程式自己
        做得到。

        挑某一個人的時候才需要 `trader_of`——選單上有那個名字，就代表已經知道
        他是第幾組了。
        """
        who = self.pending_scope.get() if self.pending_scope is not None else PENDING_ALL
        if not who or who == PENDING_ALL:
            return list(enumerate(self.accounts, start=1))
        # trader_of 的鍵本來就是從 self.accounts 長出來的，照理不會超出範圍，
        # 但這裡寧可漏掉一個也不要整個分頁掛在 IndexError 上。
        return [(order, self.accounts[order - 1]) for order in sorted(self.trader_of)
                if 1 <= order <= len(self.accounts) and self.trader_of[order] == who]

    def _refresh_pending_scope(self):
        """
        範圍選單的選項跟著已知的帳戶名單走。

        切到這個分頁時刷一次（_on_tab_changed），知道新名字的時候也要刷一次
        （_refresh_account_choices）——只靠前者的話，人一直待在這個分頁按登入，
        選單會停在登入前的樣子，而且不會自己好（2026/08/31 使用者遇到）。
        """
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
        執行緒，理由一樣：這一步也要登入／換 cookie，不能跟更新分頁或下單依序
        執行同時搶同一顆瀏覽器。
        """
        if self.busy or self.pending_busy:
            return

        targets = self._pending_targets()
        if not targets:
            show_info(self.root, "沒有帳戶可查", "找不到帳號設定，請看 .env。")
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
        rows, problems, names = [], [], {}
        for order, account in targets:
            name = self.trader_of.get(order, f"第 {order} 組")
            try:
                page, session, probs = fetch_mod.ensure_logged_in(
                    context, [(order, account)], store)[order]
                if probs:
                    problems.append(f"{name}：{'；'.join(probs)}")
                    continue
                sheet = (session.get("account") or "").strip() or name
                # 名字是這一趟登入才拿得到的，拿到就帶回去給主執行緒記著
                # （見 _on_pending_fetched）——這一頁現在是「沒登入也能按」的
                # 入口，那份對照就不能只靠「登入」那顆按鈕長出來。
                if session.get("account"):
                    names[order] = sheet
                rows.extend(order_query.query_orders(page, session, sheet))
            except RuntimeError as exc:
                problems.append(str(exc))
        return {"rows": rows, "problems": problems, "names": names}

    # ---------- 取消掛單（10.3） ----------

    def _pending_cancel_rows(self, side):
        """
        這一顆按鈕會動到哪幾列。動的是**畫面上這張表現在看得到的**，不是按下去
        才重新查一次——人是看著這張表按的，重查會讓「看到的」跟「送出去的」變成
        兩份東西（10.3 第一點）。
        """
        return [row for row in self.pending_rows
                if row["open"] and (side is None or row["side"] == side)]

    def cancel_pending(self, kind):
        """三顆取消按鈕共用這一支，差別只有挑哪幾列。"""
        if self.busy or self.pending_busy or self.pending_cancel_active:
            return

        rows = self._pending_cancel_rows(CANCEL_SIDES[kind])
        if not rows:
            show_info(self.root, CANCEL_TEXTS[kind], "現在這張表上沒有可以取消的委託。")
            return

        # 一個帳戶一則指令，所以先照帳戶分組。dict 保持插入順序，分組後的順序
        # 就是表上的順序，跟人看到的一樣。
        groups = {}
        for row in rows:
            groups.setdefault(row["sheet"], []).append(row)

        lines = []
        for sheet, items in groups.items():
            lines.append(f"　{sheet}（{len(items)} 筆）")
            for row in items:
                # 委託單／預約單混在同一批送的時候，兩者走的取消機制完全不同
                # （見 order_cancel.py／order_cancel_reservation.py），標出來
                # 讓人看得出這一筆是哪一種。
                lines.append(f"　　[{_kind_text(row)}] {row['code']} {row['side_text']} {show(row['left'])}"
                             f"　{row['price_text']}　委託書號 {row['ordno']}")
        # 不是只問「確定嗎」：會動到真實委託，要列出這一次會取消幾筆、哪幾檔（10.1）。
        if not ask_confirm(
                self.root, CANCEL_TEXTS[kind],
                f"確定要取消這 {len(rows)} 筆委託嗎？\n\n" + "\n".join(lines)
                + "\n\n送出之後沒有辦法收回來。",
                confirm_text="取消這幾筆", cancel_text="不要"):
            return

        queue = []
        for sheet, items in groups.items():
            order_number = self._order_number_for_sheet(sheet)
            if order_number is None:
                show_error(self.root, "找不到帳戶", f"{sheet} 對不到任何一組帳號，這一次不做。")
                return
            # 委託單走 order_cancel.py（委託查詢頁），預約單走
            # order_cancel_reservation.py（預約查詢頁）——兩套機制不一樣，
            # 但同一個帳戶登入一次就順便兩邊都做，見 _pending_cancel_job。
            committed = [row["ordno"] for row in items if row.get("ordstatus") != "1"]
            reservation = [row["ordno"] for row in items if row.get("ordstatus") == "1"]
            queue.append((order_number, self.accounts[order_number - 1], sheet,
                          committed, reservation))

        self.pending_cancel_active = True
        self.pending_cancel_queue = queue
        self.pending_cancel_pos = 0
        self.pending_cancel_label = CANCEL_TEXTS[kind]
        self.pending_cancel_results = []
        self.pending_cancel_problems = []
        # 借用同一顆 busy：這一輪會換 cookie，跟登入／讀取／下單同時進行就會
        # 動到別人的帳戶（跟 ui_order_exec 那一輪借 busy 的理由一模一樣）。
        self._set_busy(True, f"{self.pending_cancel_label}（{len(rows)} 筆）…")
        self._update_pending_ui()
        self._dispatch_next_cancel()

    def _dispatch_next_cancel(self):
        """派下一個帳戶。停止就是「不派下一則」——這也是唯一的中斷點。"""
        if (not self.pending_cancel_active
                or self.pending_cancel_pos >= len(self.pending_cancel_queue)):
            self._finish_pending_cancel()
            return

        order_number, account, sheet, committed, reservation = \
            self.pending_cancel_queue[self.pending_cancel_pos]
        total = len(committed) + len(reservation)
        self._say(f"{self.pending_cancel_label}：第 {self.pending_cancel_pos + 1}/"
                  f"{len(self.pending_cancel_queue)} 個帳戶（{sheet}，{total} 筆）…")
        self._ensure_browser_thread()
        self.browser_waiting += 1
        self.browser_cmd_queue.put(
            ("pending_cancel", (order_number, account, sheet, committed, reservation)))

    def _pending_cancel_job(self, context, store, order_number, account, sheet,
                            committed_ordnos, reservation_ordnos):
        """
        背景執行緒用（只能在 ui_background._browser_worker 裡呼叫）。

        一次一個帳戶：ensure_logged_in 換到這一組的 cookie，委託單跟預約單各自
        對應不同的頁面／取消機制（order_cancel.py／order_cancel_reservation.py），
        但同一個帳戶只登入這一次，兩邊都做完才換下一個帳戶——不是因為機制一樣，
        是因為換 cookie 的成本要一起省。任何一段丟出 RuntimeError／
        OrderMaybeSubmitted 都直接讓它往上傳，不在這裡攔：ui_background.py 對
        "pending_cancel" 這個 cmd 已經處理過這兩種例外（見那邊的註解），這裡
        攔了只會攔出第二套邏輯。刪單確認視窗一定在各自的 cancel_orders 裡收
        乾淨才回來——視窗還開著就換下一個帳戶的 cookie，那個視窗後來送出去的
        會是新身分（10.3 第六點）。
        """
        page, session, probs = fetch_mod.ensure_logged_in(
            context, [(order_number, account)], store)[order_number]
        if probs:
            raise RuntimeError(f"{sheet}：{'；'.join(probs)}")

        combined = {"results": [], "missing": [], "locked": []}
        if committed_ordnos:
            part = order_cancel.cancel_orders(page, session, sheet, committed_ordnos)
            for key in combined:
                combined[key].extend(part[key])
        if reservation_ordnos:
            part = order_cancel_reservation.cancel_orders(page, session, sheet, reservation_ordnos)
            for key in combined:
                combined[key].extend(part[key])
        return combined

    def _on_pending_cancelled(self, payload):
        """一個帳戶回話。"""
        self.browser_waiting = max(0, self.browser_waiting - 1)
        sheet = payload.get("sheet", "")

        if payload.get("maybe_submitted"):
            # 「確認」已經按下去了，只是沒等到結果——整批停在這裡，不再往下派。
            # 這種狀況要人自己去看，繼續刪別的帳戶只會讓「現在到底發生了什麼」
            # 更難回答（10.3 第九點）。
            self.pending_cancel_active = False
            self.pending_cancel_problems.append(payload.get("error", ""))
            self._finish_pending_cancel(maybe_submitted=True)
            return

        if "error" in payload:
            self.pending_cancel_problems.append(payload["error"])
        else:
            self.pending_cancel_results.extend(payload.get("results", []))
            for ordno in payload.get("missing", []):
                self.pending_cancel_problems.append(
                    f"{sheet}：{ordno} 已經不在委託查詢頁上（可能剛成交或已被刪除），沒有送出。")
            for ordno in payload.get("locked", []):
                self.pending_cancel_problems.append(
                    f"{sheet}：{ordno} 網站不讓刪（那一列沒有勾選框），沒有送出。")

        self.pending_cancel_pos += 1
        self._dispatch_next_cancel()

    def _finish_pending_cancel(self, maybe_submitted=False):
        """
        整批結束。**驗收是重查一次掛單**，不是看剛才那幾句回話（10.3 第八點）——
        刪單成功與否的最終答案在「委託查詢」，不在確認視窗上那四個字。
        """
        results = self.pending_cancel_results
        problems = [p for p in self.pending_cancel_problems if p]
        stopped = (not maybe_submitted and not self.pending_cancel_active
                   and self.pending_cancel_pos < len(self.pending_cancel_queue))
        label = self.pending_cancel_label or "取消掛單"

        self.pending_cancel_active = False
        self.pending_cancel_queue = []
        self.pending_cancel_pos = 0
        self._set_busy(False)
        self._update_pending_ui()

        done = sum(1 for row in results if row["ok"])
        self._say(f"{label}：送出 {len(results)} 筆，其中 {done} 筆刪單成功。"
                  + ("　已停止。" if stopped else "")
                  + (f"　{len(problems)} 筆有問題。" if problems else "")
                  + "　重查掛單中…")

        detail = "\n".join(order_cancel.describe(row) for row in results)
        if maybe_submitted:
            show_warning(self.root, "已經按下確認，但沒等到結果",
                         "\n\n".join(problems)
                         + "\n\n請等下面這次重查的結果，用查回來的表為準，不要直接再按一次。")
        elif problems:
            show_warning(self.root, f"{label}：有幾筆沒送出去",
                         (detail + "\n\n" if detail else "") + "\n".join(problems))
        elif results and done < len(results):
            show_warning(self.root, f"{label}：有幾筆券商回了失敗", detail)

        # 不管上面走哪一條路，最後一定重查——這一頁講的是「現在外面還有什麼」。
        self.refresh_pending()

    def stop_pending_cancel(self):
        """
        底部「停止全部操作」按下來的（見 ui_background.stop_all_operations）。

        停得掉的是「還沒派出去的那幾個帳戶」；正在跑的那一則沒辦法中斷，它回話
        之後就收工。已經送出去的刪單收不回來。

        跟停止下單一樣不問「確定嗎」（2026/09/01 使用者指定，見
        ui_order_exec.stop_order_execution）——底部那顆「停止全部操作」按下去
        兩條路都可能走到，只有一邊跳確認框會變成「有時候問、有時候不問」。
        原本確認框裡那幾句改成停完之後寫在狀態列上。
        """
        if not self.pending_cancel_active:
            return
        left = max(len(self.pending_cancel_queue) - self.pending_cancel_pos - 1, 0)
        self.pending_cancel_active = False
        self._say(f"{self.pending_cancel_label}：停止中，等目前這個帳戶做完。"
                  f"還沒輪到的 {left} 個帳戶不會做；已經送出去的刪單不會收回來。")

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
            show_error(self.root, "查詢掛單失敗", payload["error"])
            return

        # 這一趟登入才知道的名字。補進去之後 _refresh_account_choices 會把
        # 更新分頁的範圍、下單分頁的帳戶清單、還有這一頁自己的範圍選單一起刷新。
        learned = payload.get("names") or {}
        if learned:
            self.trader_of.update(learned)
            self._refresh_account_choices()

        self.pending_rows = payload.get("rows", [])
        self.pending_at = datetime.datetime.now()
        self._fill_pending()

        problems = payload.get("problems") or []
        open_count = sum(1 for row in self.pending_rows if row["open"])
        self._say(f"掛單：{len(self.pending_rows)} 筆委託，其中 {open_count} 筆還掛在外面。"
                  + (f"　{len(problems)} 個帳戶查不到。" if problems else ""))
        if problems:
            show_warning(self.root, "有帳戶查不到",
                         "這幾個帳戶沒查到委託：\n\n" + "\n".join(problems))

    def _fill_pending(self):
        """整份重建，跟更新分頁 fill_sync_tree()／下單分頁執行預覽同一個做法。"""
        for item in self.pending_tree.get_children():
            self.pending_tree.delete(item)

        for row in self.pending_rows:
            # 還掛在外面的用買賣底色（跟下單分頁的執行預覽同一組顏色，也跟網站
            # 本身買紅賣綠一致）；已經成交／取消／失敗的淡化——它們是「已經結束
            # 的事」，不該看起來跟還能取消的那幾筆一樣醒目。
            tag = {"B": "buy", "S": "sell"}.get(row["side"], "") if row["open"] else "done"
            # 順序跟 ui_layout._build_pending_tab 的 columns 一模一樣，那邊是照
            # 網站「委託查詢」那張表排的（帳戶、單別這兩欄是程式多的——「單別」
            # 是 2026/09/02 加的，預約單開始跟委託單混在同一張表之後，不標出來
            # 人分不出「委託書號」那一欄印的到底是 ordno 還是 preordno）。
            self.pending_tree.insert("", "end", values=(
                row["sheet"], _kind_text(row), row["ordered_at"], row["work_date"], row["ordno"],
                row["code"], row["trade_text"], row["side_text"], row["price_text"],
                show(row["qty"]), show(row["cancelled"]), show(row["matched"]),
                row["flag_text"], row["status"], show(row["left"]),
            ), tags=(tag,) if tag else ())

        self._resize_pending_columns()
        self._update_pending_ui()

    def _resize_pending_columns(self):
        """
        每一欄都量過再給寬度，不用固定數字猜。

        2026/08/31 出過的問題：欄寬是猜的，字級／字型／DPI 只要有一項跟猜的
        時候不一樣就切字——「委託日期」被切成「2026/08/31 09:00:4.」、
        「有效交易日」剩「2026/08/3」。所以這裡直接量 Treeview 現在用的那組
        字型：表頭是粗體、內容是一般體，兩邊取大的（`col_width` 已經把左右
        留白算進去了）。

        **minwidth 一定要跟著設**：欄寬總和比表格寬的時候，Tk 會去壓
        stretch 的那一欄，壓到 minwidth 為止（預設只有 20 像素）——只設
        width 擋不住切字。設好之後 Tk 就不壓了，改用下面那條水平捲軸。

        量的是畫面上真正那幾格的字（`tree.set`），不是 pending_rows 裡的原始
        值：兩邊要是哪天不一致，會切字的是畫面上那一份。
        """
        tree = self.pending_tree
        items = tree.get_children()
        for key in tree["columns"]:
            need = max(
                col_width(self.family, [tree.heading(key)["text"]], bold=True),
                col_width(self.family, [str(tree.set(item, key)) for item in items],
                          minimum=wide(self.pending_widths[key])),
            )
            # 「委託狀態」放的是網站回的失敗原因（errmsg 是自由文字，量不出
            # 上限），照量的給會把整張表撐爆，給一個上限就好——它是 stretch
            # 的那一欄，畫面有空間的時候本來就會自己變寬。
            if key == "status":
                need = min(need, wide(300))
            tree.column(key, width=need, minwidth=need)

    def _update_pending_ui(self):
        """查詢按鈕、右上角的時間戳，以及三顆取消按鈕的亮／灰。"""
        for kind, button in self.pending_cancel_buttons.items():
            # 亮的條件：這一類真的有東西可取消，而且現在沒有別的事在跑。
            # 「沒查過就不能取消」是順帶做到的（沒查過就沒有列）——人是看著這張
            # 表按下去的，不能讓按鈕比表新（10.3 第一點）。
            can = bool(self._pending_cancel_rows(CANCEL_SIDES[kind]))
            free = not (self.busy or self.pending_busy or self.pending_cancel_active)
            button.configure(state="normal" if (can and free) else "disabled")

        if self.pending_busy:
            self.pending_button.configure(text="查詢中…", state="disabled")
        else:
            self.pending_button.configure(
                text="查詢掛單", state="disabled" if self.busy else "normal")

        if self.pending_at is None:
            self.pending_stamp.configure(text="還沒查過")
        else:
            self.pending_stamp.configure(text=f"最後查詢 {self.pending_at:%H:%M:%S}")
