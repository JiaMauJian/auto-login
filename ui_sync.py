"""更新分頁的顯示與操作：左邊名單、右邊常駐狀態列與訊息框。"""

import datetime

import ledger as ledger_mod
import planner
from util import same_number, show, to_num
from ui_common import (WHEN_TODAY, ask_opening_balance, cash_method_toggle_enabled, stock_title,
                        within)


def _row_of(cell):
    """
    從 "E5" 這種格子名稱取出列號。股票併行時要照 Excel 列的順序排（見
    _format_today_events），不是照歷程檔裡實際寫入的先後。
    """
    return int(cell[1:])


def _cash_formula(event):
    """
    現金那一行的算式：「{算法名稱} = {算式} = {結果}」，兩種算法都是這個形狀
    （2026/08/22 使用者訂正：不管有沒有轉負都要看得到算法，不是只有轉負那一次）。

    planner._bank_note／_formula_note 算好的 note 長這樣：

        銀行餘額 + 淨收付(T+0) + 淨收付(T+1) = 893 - 238 + 0 = 655
        今日初始現金餘額 + 今日淨收付 = 893 - 238 = 655

    第一個 " = " 前面那段（"銀行餘額 + 淨收付(T+0) + …"）是標籤，換算法名稱
    打頭是重複資訊——算法名稱已經講了是哪一種算法，數字才是這一行真正要看的
    東西，所以只取第一個 " = " 之後的部分。

    `by == "adopt"` 的兩種（今天第一次登入設基準、「修改今日初始現金餘額」沒有
    格子要寫）在 _cash_line 就先分流掉了，不會走到這裡。「修改今日初始現金
    餘額」如果真的寫了格子（`by == "program"`），note 是 apply_cash_reset 算的，
    2026/08/24 起跟自動算出來那句共用 planner._formula_note，形狀一樣，這裡
    不用特別分流。最後的 `note or show(...)` 只是留給格式萬一對不上時的保底。
    """
    note = event.get("note") or ""
    if note.startswith("銀行餘額") and " = " in note:
        return f"{planner.METHOD_NAMES[planner.METHOD_BANK]} = {note.partition(' = ')[2]}"
    if note.startswith("今日初始現金餘額") and " = " in note:
        return f"{planner.METHOD_NAMES[planner.METHOD_OPENING]} = {note.partition(' = ')[2]}"
    return note or show(event.get("new"))


def _cash_line(time_text, event):
    """
    現金那一行。負的就整行標紅——不管是不是這一輪才由正轉負，只要結果是負的
    就紅（2026/08/22 使用者訂正：原本只有「這一輪由正轉負」那一次才標紅，
    導致連續好幾筆都是負的時候，除了第一筆，後面每一筆的「老早就負的」都沒
    紅字提醒，看起來像沒事——這裡跟名單上現金負的名字變紅字是同一個規矩，
    看「現在是不是負的」，不看「是不是這一輪才變負的」）。

    `by == "adopt"` 的兩種情況——今天第一次登入把 B8 收成今日起點
    （`planner.initialize`）、「修改今日初始現金餘額」算出來剛好等於 Excel
    上的數字所以沒有格子要寫（`planner.apply_cash_reset`）——都不是套公式
    算出來的餘額異動，Excel 那一格根本沒被寫過。原本兩者都套用 [餘額更新]
    標籤、只印一句沒帶數字的說明（例如「今日初始現金餘額（今天第一次登入時
    的 B8）」），語意不明：既不像更新（沒有算式、沒有結果數字），也讓人以為
    這一輪真的改了 B8。改用專屬標籤 [今日初始餘額] 並印出實際數字，跟真正
    套公式寫入的 [餘額更新] 分開。
    """
    new = to_num(event.get("new"), None)
    negative = new is not None and new < 0
    if event.get("by") == "adopt":
        return (f"{time_text} [今日初始餘額] {show(event.get('new'))}", "neg" if negative else None)
    return (f"{time_text} [餘額更新] {_cash_formula(event)}", "neg" if negative else None)


def _stock_lines(time_text, events):
    """
    同一輪、同一檔股票的股數／成本併成一行。歷程檔裡股數、成本是兩筆分開的
    事件，這裡再照股票名稱分一次（呼叫這裡之前已經先照 at 分過同一輪了）。
    """
    titles, order = {}, []
    for event in events:
        title = stock_title(event.get("label", ""))
        if title not in titles:
            titles[title] = {"row": _row_of(event.get("cell", "A0"))}
            order.append(title)
        which = "股數" if event.get("label", "").startswith("股數") else "成本"
        titles[title][which] = event.get("new")
    order.sort(key=lambda title: titles[title]["row"])

    return [
        (f"{time_text} [股票更新] {title}　"
         + "　".join(f"{key} {show(titles[title][key])}" for key in ("股數", "成本")
                    if key in titles[title]),
         None)
        for title in order
    ]


def _warning_line(time_text, text):
    """
    股票提醒那一行（見 UiSyncMixin._fill_notes）。2026/08/22 使用者要求：
    跟其他行統一格式——不再堆在訊息框最後面（會被今天累積的歷程往下擠、
    要捲很多才看得到，見 docs/更新分頁訊息框改版.md），改成掛時間戳、混進
    同一個時間序，跟 `_cash_line`／`_stock_lines` 同一種「時間 [標籤] 內容」
    形狀。字用深黃色（`warn` tag，`self.colors.warning`）——份量比一般歷程
    重，2026/08/22 使用者再次要求跟其他行顏色分開。現金被擋住的那幾句
    （`planner._cash` 的 `[現金]` 開頭）還在討論中，不走這裡。
    """
    return (f"{time_text} [警告] {text}", "warn")


def _error_line(time_text, text):
    """
    這一位帳號整組失敗（登入逾時、Excel 找不到分頁之類，見
    `ui_background._on_fetched` 的 `problem_of`）那一行。2026/08/22 使用者
    要求：不要全部堆在系統狀態列或所有人共用的提醒框裡，改成跟其他行同一種
    「時間 [標籤] 內容」形狀，只出現在這一組對應到的那位交易人自己的訊息框
    （見 UiSyncMixin._fill_notes 用 trader_of 反查是誰）。真的還查不出是誰
    的那極少數情況才會落到 fallback，一樣走這裡（見 _fill_notes）。字用紅色
    （`err` tag，`self.colors.danger`）——比警告更急，2026/08/22 使用者要求
    跟 [警告] 分開顏色。
    """
    return (f"{time_text} [異常] {text}", "err")


def _dedupe_cash_rows(rows):
    """
    現金那幾筆事件，結果（`new`）跟同一種算法上一筆顯示過的一樣就丟掉，不進
    訊息框（2026/08/22 使用者要求；2026/08/24 訂正成「同一種算法分開比」，
    見下）。

    現金每次讀取一定會覆蓋 Excel、也一定會記進歷程（見
    docs/現金餘額兩種算法.md「B8 無條件覆蓋」）——這裡動的只是訊息框要不要
    重複顯示同一個結果，不影響 Excel 寫入或歷程檔，「歷程」分頁還是看得到
    每一筆。比對的是算出來的結果，不是整句算式：算式裡淨收付的細節可能每次
    都有一點出入，但只要最後的餘額沒變，對使用者來說就是「沒事」。

    2026/08/24 訂正：原本不分算法，`by == "adopt"` 的今日初始餘額跟
    `by == "program"` 的正常寫入全部混在一起比。切換現金算法（初始餘額累加
    ↔ 銀行餘額推算）之後，新算法第一次算出的結果只要數字剛好跟切換前最後
    一筆一樣，就會被當成「沒變」吞掉，使用者看不到新算法真的重算過一次。
    現在「銀行餘額推算」自成一組 key，只跟自己上一筆比；「今日初始餘額」
    跟「初始餘額累加」的寫入還是同一組——它們本來就是同一個基準（baseline
    剛設好、還沒加淨收付；或加了淨收付但淨收付是 0），數字相同時本來就該
    只顯示一次，這條沒有變。

    「今天第一筆」（每一組 key 各自的第一筆）一定留著（`last_shown` 用
    `None` 當哨兵，`same_number` 對 None 一律回 False）。

    照時間先後比對，不是照畫面顯示的順序（畫面是新的在最上面）。
    """
    ordered = sorted(rows, key=lambda row: row.get("at") or "")
    kept, last_shown = [], {}
    for row in ordered:
        if row.get("label") != "現金餘額":
            kept.append(row)
            continue
        note = row.get("note") or ""
        key = "bank" if note.startswith("銀行餘額") else "opening"
        new = row.get("new")
        if not same_number(last_shown.get(key), new):
            kept.append(row)
            last_shown[key] = new
    return kept


def _format_today_events(rows):
    """
    把某一位「今天」的歷程事件（含 `label == "警告"` 的股票提醒、
    `label == "異常"` 的整組帳號失敗合成列，見 UiSyncMixin._fill_notes），
    排成訊息框要顯示的 (文字, 標籤) 清單。

    同一輪內（同一個 at）異常排最前面、警告第二、現金第三、股票照 Excel
    列的順序排在最後——不管歷程檔裡實際寫入的先後（planner.commit 是股票先、
    現金最後才 append）。整體新的在最上面，跟「歷程」分頁
    ui_history._fill_history 同一個慣例。

    現金那三筆裡如果 [今日初始餘額]（`by == "adopt"`，來自 planner.initialize）
    跟 [餘額更新]（`by == "program"`，來自 planner.commit）同一個 at 都有——
    今天第一次讀取這個人時最常見——[餘額更新] 排在上面：它是這一輪實際
    套用公式寫進 Excel 的結果，時間上也確實晚於「先把 B8 收成起點」那一步，
    「整體新的在最上面」這個慣例套到同一個 at 裡也該一樣（2026/08/24 使用者
    訂正；原本兩者維持歷程檔裡的寫入先後，[今日初始餘額] 反而排在上面）。

    [今日初始餘額] 只有在「這一輪」也用「銀行餘額推算」寫出結果時才拿掉——
    銀行餘額推算用不到這個基準，兩個一起印會讓人以為基準是算式的一部分
    （2026/08/23 使用者要求）。判斷依據是這一輪 cash_events 裡有沒有算法是
    銀行餘額推算的 program 事件，不是「現在選的是哪個算法」：2026/08/24 之前
    是看現在選哪個算法，切換算法之後，連當天稍早、還在用另一種算法時留下的
    [今日初始餘額] 也會被現在的算法選擇連坐一起拿掉（使用者訂正）。
    """
    groups, order = {}, []
    for row in rows:
        at = row.get("at") or ""
        if at not in groups:
            groups[at] = []
            order.append(at)
        groups[at].append(row)
    order.sort(reverse=True)

    lines = []
    for at in order:
        time_text = at[11:19] if len(at) >= 19 else at
        group = groups[at]
        error_events = [row for row in group if row.get("label") == "異常"]
        warning_events = [row for row in group if row.get("label") == "警告"]
        cash_events = [row for row in group if row.get("label") == "現金餘額"]
        # [餘額更新]（program）排在 [今日初始餘額]（adopt）前面，見上方 docstring。
        cash_events.sort(key=lambda row: row.get("by") == "adopt")
        if any(row.get("by") == "program" and (row.get("note") or "").startswith("銀行餘額")
               for row in cash_events):
            cash_events = [row for row in cash_events if row.get("by") != "adopt"]
        stock_events = [row for row in group
                        if row.get("label", "").startswith(("股數（", "成本（"))]
        lines.extend(_error_line(time_text, event["text"]) for event in error_events)
        lines.extend(_warning_line(time_text, event["text"]) for event in warning_events)
        lines.extend(_cash_line(time_text, event) for event in cash_events)
        lines.extend(_stock_lines(time_text, stock_events))
    return lines


class UiSyncMixin:
    # ---------- 更新分頁 ----------

    def replan(self):
        """重新計算提案並重畫。plan 是純函式，可以隨便重跑。"""
        self.proposals, self.warnings = {}, {}
        for name, record in self.records.items():
            data = self.sheet_data.get(name)
            if not data:
                continue
            book = self.ledger.sheet(name)
            book["account_code"] = record.get("account_code", "")
            items, warns = planner.plan(data, record, book, self.today, self.cash_method.get())
            self.proposals[name] = items
            self.warnings[name] = warns

        self.fill_sync_tree()

    def fill_sync_tree(self):
        """左邊的名單與右邊的狀態列＋訊息框一起重畫。"""
        self._fill_people()
        self._fill_right()

    def _summary(self, name):
        """
        一位交易人的濃縮狀態：(要不要標記 ⚠, 現金顯示值, 現金是不是負的)。

        標記只看訊息框裡會不會出現 [警告]／[異常]（planner warnings、整組帳號
        失敗，見 `_fill_notes`），跟這輪有沒有格子要寫無關——單純寫入是正常
        狀況，訊息框本來就會印 [股票更新]／[餘額更新]，不需要在名單上另外
        標記（2026/08/22 使用者要求：⚠ 只保留給訊息框真的有警告/異常的時候，
        避免使用者看到 ⚠ 卻在訊息框找不到對應的警告字樣）。
        """
        items = self.proposals.get(name, [])
        flagged = bool(self.warnings.get(name)) or any(
            self.trader_of.get(problem["order"]) == name for problem in self.problems) or (
            name in self.cash_baseline_errors and self.cash_method.get() == planner.METHOD_OPENING)

        cash = next((item for item in items if item["kind"] == "cash"), None)
        if cash is None:
            return flagged, "", False
        value = cash["proposed"] if cash["will_write"] else cash["current"]
        number = to_num(value, None)
        return flagged, show(value), number is not None and number < 0

    def _shown(self):
        """名單上現在真的有誰。上一位／下一位都照這個走，才會跟眼睛看到的一致。"""
        return list(self.people.get_children())

    def _fill_people(self):
        self.people.delete(*self.people.get_children())

        names = list(self.proposals)
        # 正在看的那位不見了（換了檔、重讀、改了篩選）就退回第一位。
        # 右邊不能停在一個名單上已經沒有的人身上。
        #
        # 名單一有人就一定要選中一位，不能等使用者自己點：名單是網頁資料長出來的
        # （沒讀到的人不會在上面），讀完卻誰都沒選中的話，右邊會整片空白
        # （見 _fill_right），明明剛讀完、左邊也看得到人，右邊卻像什麼都沒讀到。
        # 剛開程式什麼都還沒讀的時候名單是空的，這行自己會落到 None，
        # 不必另外擋「還沒選過人」。
        if self.current_sheet not in names:
            self.current_sheet = names[0] if names else None

        need = 0
        for name in names:
            # 現金餘額現在是隱藏欄（見 _build_people），值照樣填 —— 隱藏的意思是
            # 「先不顯示」，不是「不算」，那一欄要回來的時候不必再回頭補這裡。
            flagged, cash, negative = self._summary(name)
            tags = []
            if flagged:
                need += 1
                flag = "⚠"
                tags.append("attention")
            else:
                flag = "✓"
            if negative:
                tags.append("negative")
            self.people.insert(
                "", "end", iid=name,
                text=name + ("（模擬）" if name in self.fake_sheets else ""),
                values=(cash, flag),
                tags=tuple(tags),
            )

        total = len(self.proposals)
        if not total:
            self.people_count.configure(text="")
        elif need:
            self.people_count.configure(text=f"{need} / {total} 位要處理")
        else:
            self.people_count.configure(text=f"共 {total} 位，都一致")

        if self.current_sheet:
            self.people.selection_set(self.current_sheet)
            self.people.see(self.current_sheet)

    def _on_person_selected(self, _event=None):
        picked = self.people.selection()
        # selection_set 自己也會觸發這個事件。選的還是同一位就別再畫一次，
        # 否則每次重建名單都要多畫一次右邊。
        if not picked or picked[0] == self.current_sheet:
            return
        self.current_sheet = picked[0]
        self._sync_scope_to_person()
        self._fill_right()

    def _step_person(self, delta):
        """換上一位／下一位。名單怎麼排就怎麼走，到頭了就停住（不繞回去）。"""
        names = self._shown()
        if not names:
            return
        index = names.index(self.current_sheet) if self.current_sheet in names else 0
        index = max(0, min(len(names) - 1, index + delta))

        # 自己更新，不靠 <<TreeviewSelect>> —— 那個事件是排進事件佇列才送的，
        # 「按鈕按了、右邊要立刻換」不該押在佇列什麼時候輪到它上面。
        # selection_set 之後那個事件還是會來，但它看到選的是同一位就直接跳過。
        self.current_sheet = names[index]
        self.people.selection_set(self.current_sheet)
        self.people.see(self.current_sheet)
        self._sync_scope_to_person()
        self._fill_right()

    def _fill_right(self):
        """右邊：常駐狀態列＋訊息框，都是選中的那位交易人的。"""
        name = self.current_sheet
        # 訊息框的標題掛上名字（「交易人A　訊息」），不塞進內容第一行——標題貼在
        # 框邊、不隨內容捲動，換人換得快的時候一眼就看得到是誰，比塞進會被捲走
        # 的第一行穩妥（2026/08/22 使用者要求；原本的表格標頭見 ui_layout._build_detail）。
        if name:
            who = name + ("（模擬）" if name in self.fake_sheets else "")
            self.msg_frame.configure(text=f"{who}　訊息")
        else:
            self.msg_frame.configure(text="訊息")
        self._fill_status(name)
        self._fill_notes(name)

    def _fill_status(self, name):
        """
        右上角常駐狀態列：現金算法、今日初始現金餘額、現在的現金餘額、
        「✓ 與網頁一致」小提示、修改按鈕——不管這一輪有沒有異動都在，跟
        訊息框（今天發生了什麼）分開（見 docs/更新分頁訊息框改版.md）。

        「現金餘額」這一項固定顯示目前算出來的數字，2026/08/22 使用者要求：
        訊息框那幾行 [餘額更新] 之後會去重（同樣的結果不再重複印，見
        _cash_line），這裡就不能只靠「訊息框最後一行是多少」來回答「現在是
        多少」——那一行可能是很久以前印的。負的比照名單上現金負的名字變紅字，
        同一個規矩。
        """
        item = self._cash_item(name)
        asked = self.path in self.cash_method_asked

        self.method_label.configure(
            text=f"現金算法：{planner.METHOD_NAMES[self.cash_method.get()]}" if asked else "")
        cursor = "hand2" if asked and cash_method_toggle_enabled() else ""
        if self.method_label.cget("cursor") != cursor:
            self.method_label.configure(cursor=cursor)

        opening_method = self.cash_method.get() == planner.METHOD_OPENING
        if opening_method:
            self.opening_label.configure(text=f"今日初始現金餘額：{self._opening_text(name, item)}")
        self._show_opening_row(opening_method)

        if item is None:
            self.balance_label.configure(text="現金餘額：—", foreground="")
        else:
            balance = item["proposed"] if item["will_write"] else item["current"]
            number = to_num(balance, None)
            self.balance_label.configure(
                text=f"現金餘額：{show(balance)}",
                foreground=self.colors.danger if number is not None and number < 0 else "")

        # 「✓ 與網頁一致」小提示：這一位讀過（或修改過今日初始現金餘額）就有
        # self.round_at，時間戳掛在這裡；被擋住（現金還沒設基準）或有提醒
        # （例如「網頁庫存已無此檔」）在跑的時候不顯示——那時候不算「一致」，
        # 提醒本身已經在講了，這裡不能講反話（見 docs/更新分頁訊息框改版.md）。
        at = self.round_at.get(name)
        quiet_ok = bool(at) and not self.warnings.get(name) and (item is None or not item["blocked"])
        if quiet_ok:
            time_text = at[11:19] if len(at) >= 19 else at
            self.quiet_label.configure(text=f"✓ 與網頁一致（{time_text}）")
        else:
            self.quiet_label.configure(text="")

        self.opening_ready = (self.ledger is not None and item is not None
                              and not item["blocked"] and opening_method)
        self._sync_buttons()

    def _opening(self, name):
        """某一位現在的現金基準（紀錄檔裡的今日初始現金餘額）。沒有就 None。"""
        cash = self.ledger.sheet(name)["cash"] if (self.ledger is not None and name) else None
        return ledger_mod.opening_balance(cash) if cash is not None else None

    def _opening_text(self, name, item):
        """今日初始現金餘額那一格要寫什麼字。"""
        opening = self._opening(name)
        # 一位都還沒選（還沒讀過網頁資料）時寫破折號而不是「還沒設定」——
        # 那時候是「不知道要看誰」，不是「這個人沒有基準」。
        text = "—" if not name else ("(還沒設定)" if opening is None else show(opening))
        if item is not None and item["reset_to"] is not None:
            text = f"{text} → {show(round(item['reset_to'] - item['net'], 2))}"
        return text

    def _on_method_click(self, _event=None):
        """
        點一下「現金算法」的名字：換成另一種。只有已經問過、名字露臉之後才點得到
        ——這次執行還沒問過時 method_label 是空字串，點了也不該有反應。
        """
        if self.path in self.cash_method_asked:
            self._toggle_cash_method()

    def _show_opening_row(self, show_it):
        """
        用銀行餘額推算的日子，「今日初始現金餘額」與「修改」一起收起來，不占畫面。

        它們講的是初始餘額累加那一種算法的基準，今天既然不用那種算法，擺著只是
        佔位置。基準本身不會消失，只是沒顯示：每天照樣設，明天切回初始餘額累加
        就要用它。
        """
        if show_it == bool(self.opening_button.winfo_manager()):
            return
        if show_it:
            self.opening_label.grid()
            self.opening_button.grid()
        else:
            self.opening_label.grid_remove()
            self.opening_button.grid_remove()

    def edit_opening(self):
        """
        「修改」按下去：問一個新的開盤前現金，套進提案，然後照這一輪的規矩落實。

        套用走的是 planner.apply_cash_reset —— 跟程式自己跳的那個對話框同一段
        程式碼，兩個入口不會算出不一樣的結果。

        落實直接寫進 Excel 並落帳（沒有格子要寫也一樣要落帳，理由見 _commit_round）。
        """
        name = self.current_sheet
        item = self._cash_item(name)
        # 用銀行餘額推算的日子這顆按鈕根本不在畫面上，這裡再擋一次是因為
        # 「看不到」跟「按不到」是兩件事 —— 它算的是另一種算法的答案。
        if (self.ledger is None or item is None or item["blocked"] or self.busy
                or self.cash_method.get() != planner.METHOD_OPENING):
            return

        cash = self.ledger.sheet(name)["cash"]
        opening = ask_opening_balance(self.root, self.family, name,
                                      ledger_mod.opening_balance(cash))
        if opening is None:
            return

        planner.apply_cash_reset(item, opening)
        # 這顆按鈕動到的只有眼前這一位，寫入與落帳的範圍就跟著縮到他身上
        # —— 名單上別人那些「要寫」的格子是上一輪算的，不該被這一下順手寫出去。
        self.round_scope = {name}
        self.round_at[name] = datetime.datetime.now().isoformat(timespec="seconds")
        self.fill_sync_tree()

        writes, total = self._collect_writes()
        if total:
            self._begin_write(writes, total)
            return

        # 算出來剛好等於 Excel 上的數字，一格都不必寫 —— 但基準確實被改掉了，
        # 這一筆不落帳就等於沒按過（見 _commit_round）。
        recorded = self._commit_round()
        # 順序不能反，理由同 ui_background._on_fetched／_on_written：
        # refresh_history() 要先把剛落帳的這一筆讀回 self.history_rows，
        # replan() 的 fill_sync_tree() 畫訊息框才看得到。
        self.refresh_history()
        self.replan()
        self._say(f"{name} 的今日初始現金餘額改成 {show(opening)}，"
                  "跟 Excel 上的數字剛好一樣，沒有格子要寫。"
                  + (f"紀錄檔更新了 {recorded} 筆（見歷程）。" if recorded else ""))

    def _fill_notes(self, name):
        """
        訊息框內容：今天的歷程——直接篩 self.history_rows（sheet == name 且
        at 是今天）——跟這一輪的股票提醒（`self.warnings[name]`，見
        planner.plan）、這一位所屬組別的整組失敗原因（`self.problems`，見
        `ui_background._on_fetched` 的 `problem_of`）混在一起，一起照時間
        重排（見 _format_today_events），整體新的在最上面。

        提醒、失敗原因都不堆在固定位置（2026/08/22 再訂正）：原本兩者都排在
        所有歷程後面，帳號測久了訊息框累積很多行之後會被擠出這個小框的可視
        範圍，要往下捲才看得到，等於沒提醒。改成掛時間戳、跟其他行一樣照
        時間排序，混進同一個時間序（見 _warning_line／_error_line）——不會
        永遠釘在哪裡，但也不會被埋掉，跟其他行一樣「新的擠掉舊的」。

        **失敗原因現在按交易人分流**（2026/08/22 使用者要求，不要塞進系統
        狀態列，也不要 20 個人共用同一份提醒）：`self.problems` 每一項帶著
        `order`（第幾組），這裡用 `self.trader_of.get(order)` 反查是哪一位，
        等於這個人就只把屬於自己那幾組的失敗原因併進自己的訊息框，換人看到
        的是換人自己的。極少數「連是誰都查不出來」的情況（`trader_of` 也沒有
        這個 order，通常是這組帳號從來沒有成功登入過、根本沒有名字可以歸戶）
        沒有任何交易人的分頁掛得上，退回舊的做法：不管選中誰都顯示，前綴
        補回 `TBB_ID_{order}` 自報身分（下面 fallback 那段；2026/08/22
        使用者要求用這個名字而不是「第 N 組」——連不上人的時候，這個名字
        才是能讓人直接回頭去 .env 對到是哪一行的線索）。

        現金被擋住的那幾句（`planner._cash` append 的 `[現金] ...`）還在
        討論中，暫時維持原樣排在最後面，不進這個時間序。

        「這一輪跟網頁一致、不用更新」這件事也不在這裡講（見
        docs/更新分頁訊息框改版.md）：搬到右上角常駐狀態列的「✓ 與網頁一致」
        小提示（見 _fill_status），原地更新不往下疊，跟這裡「今天發生了
        什麼事」的定位分開。

        原本第一行是「這份資料是誰的、什麼時候讀的」（`簡嘉懋　讀取於
        19:54:23`），2026/08/22 使用者要求拿掉——每一行歷程自己就帶著時間戳，
        再加一行整份資料的讀取時間是重複資訊。少了它，「這份資料是不是這一輪
        剛讀的，還是上一輪留著的舊資料」這件事目前沒有畫面上的信號提醒（原本
        靠這行的讀取時間分辨，見 ui.py 開頭「只更新一位有兩件事要守住」那段）
        ——真的需要分辨時只能回頭看「歷程」分頁或訊息框裡最新一行歷程的時間。

        現金結果連續沒變的那幾筆會被 `_dedupe_cash_rows` 收掉，不然每次讀取
        都硬要覆蓋 B8（見 docs/現金餘額兩種算法.md）會讓這裡洗版——現在算
        出來的實際數字，改看右上角常駐狀態列的「現金餘額」（見 _fill_status），
        不必在這裡找。
        """
        lines = []
        if name:
            today = datetime.date.today()
            rows = [row for row in self.history_rows
                   if row.get("sheet") == name and within(row.get("at"), WHEN_TODAY, today)]
            # 股票提醒、這一位的失敗原因都包成跟歷程事件同一種形狀
            # （label 固定 "警告"／"異常"），才能混進 _dedupe_cash_rows／
            # _format_today_events 同一套排序，不必另外寫一套合併邏輯。
            # [現金] 開頭的還在討論中，不動它；失敗原因不套「今天」的篩選——
            # 那是要留到「這一組自己再讀一次成功為止」的持續狀態，不是今天
            # 才發生、過午夜就該消失的一次性事件（見 _refresh_problems）。
            at = self.round_at.get(name) or ""
            warning_rows = [{"at": at, "label": "警告", "text": warn}
                           for warn in self.warnings.get(name, [])
                           if not warn.startswith("[現金] ")]
            error_rows = [{"at": problem["at"], "label": "異常", "text": problem["text"]}
                         for problem in self.problems
                         if self.trader_of.get(problem["order"]) == name]
            # B8 空白設不成今日初始餘額只在「初始餘額累加」算法底下要緊——銀行餘額
            # 推算根本不讀 B8，這則提醒對那個算法只是噪音（2026/08/25 使用者要求）。
            baseline_error = self.cash_baseline_errors.get(name)
            if baseline_error and self.cash_method.get() == planner.METHOD_OPENING:
                error_rows.append({"at": baseline_error["at"], "label": "異常",
                                   "text": baseline_error["text"]})
            lines.extend(_format_today_events(_dedupe_cash_rows(rows) + warning_rows + error_rows))

        lines += [(warn, None) for warn in self.warnings.get(name, []) if warn.startswith("[現金] ")]
        # fallback：這個 order 連是哪一位都反查不出來，沒有分頁掛得上，
        # 不管選中誰都顯示。自報身分用 `TBB_ID_{order}`（2026/08/22 使用者
        # 要求）而不是「第 N 組」——.env 就是用這個變數名稱編號的
        # （見 login.load_accounts），連不上人的時候，這個名字才是能讓人
        # 直接回頭去 .env 對到是哪一行的線索；「第 N 組」單看畫面猜不出
        # 是哪一組帳密。一樣帶時間戳、跟 [異常] 同一個標籤，格式要統一，
        # 不能因為是 fallback 就少了時間。
        lines += [_error_line(problem["at"][11:19] if len(problem["at"]) >= 19 else problem["at"],
                              f"TBB_ID_{problem['order']} {problem['text']}")
                 for problem in self.problems
                 if self.trader_of.get(problem["order"]) is None]

        self.warn_box.configure(state="normal")
        self.warn_box.delete("1.0", "end")
        for index, (text, tag) in enumerate(lines):
            if index:
                self.warn_box.insert("end", "\n")
            self.warn_box.insert("end", text, (tag,) if tag else ())
        self.warn_box.configure(state="disabled")

    def _sync_buttons(self):
        """上面那幾顆能不能按。畫面上會變灰的按鈕就剩它們。"""
        # 「更新」要 Excel 開著：讀完要拿它的現值算提案，寫入更是直接改它。
        # not self._excel_in_use() 而不是 not self.busy：下單分頁那幾條路也在用
        # COM 動同一份活頁簿（見 ui_background._excel_in_use）。
        self.fetch_button.configure(
            state="normal" if self.excel_open and not self._excel_in_use() else "disabled")

        # 「登入」不看 Excel —— 跟「全部登出」同一個規矩（那一顆本來就是這樣，
        # 見 ui_layout._build_toolbar）。這兩顆是一對，管的都是瀏覽器 session：
        # 登入的對象來自 .env（login.load_accounts 一路數 TBB_ID_1、TBB_ID_2…），
        # 不是 Excel 分頁；背景那條路也明講「登入按鈕只做登入，不碰 Excel」
        # （見 ui_background._browser_worker 的 cmd == "login"）。
        #
        # 2026/08/30 之前這一顆跟「讀取」共用同一道 excel_open 關卡，理由寫的是
        # 「後面每一步都要 Excel」。那句話在只有更新分頁的時候是對的，掛單分頁
        # 長出來之後就不是了 —— ui_pending.py 一格 Excel 都不碰，卻因為交易人
        # 姓名只能靠登入拿到、而登入被 Excel 擋著，變成非先開 Excel 不可。拆掉
        # 這道關卡不會開洞：真正要 Excel 的兩條路各自有守門員（start_fetch、
        # ui_order.refresh_order_data 都自己叫 _require_excel()）。
        #
        # 留下來的 not self.busy 擋的不是 Excel 是 cookie：整個瀏覽器只有一組
        # cookie，下單那一輪借的就是這顆鎖（見 ui_order_exec.py 開頭那段「送錯
        # 帳戶」的說明），這時候按登入會把手上這組 cookie 換掉。
        self.login_button.configure(state="normal" if not self.busy else "disabled")
        # 「全部登出」不跟 excel_open 掛勾 —— 它管的是瀏覽器 session，
        # 跟有沒有選 Excel 檔無關；沒有瀏覽器可登出時按下去只會被 start_logout_all
        # 自己擋下來、在狀態列講一句，這裡不必先幫它擋。
        self.logout_button.configure(state="normal" if not self.busy else "disabled")
        self._apply_scope_state()
        # 「修改」不看 Excel 開著沒 —— 它改的是紀錄檔裡的基準，要寫 Excel 的時候
        # 寫入那邊自己會把檔案開起來。能不能按只看「這一位有沒有網頁資料」，
        # 那是 _fill_status 判的。用銀行餘額推算的日子整顆被藏起來
        # （_show_opening_row），這裡照樣設 state —— 藏著的按鈕設得動，
        # 放回來的時候就已經是對的狀態了。
        self.opening_button.configure(
            state="normal" if self.opening_ready and not self.busy else "disabled")
