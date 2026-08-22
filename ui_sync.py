"""同步分頁的顯示與操作：左邊名單、右邊明細、現金那張表。"""

import ledger as ledger_mod
import planner
from util import show, to_num, values_match
from ui_common import ask_opening_balance, fit_to_content


def _neg_tag(value):
    """負的數字才上紅字。空的、不是數字的都不算。"""
    number = to_num(value, None)
    return "neg" if number is not None and number < 0 else None


def stock_title(label):
    """從「股數（2059 川湖）」取出「2059 川湖」。取不到就原樣顯示。"""
    inside = label.partition("（")[2].rstrip("）")
    return inside or label


class UiSyncMixin:
    # ---------- 同步分頁 ----------

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

    def _value_text(self, name, item, compare_web=True):
        """
        一格在畫面上要寫什麼字，外加它是不是這一輪剛被寫過。

        沒事就只寫現在的數字 —— 一格擺出舊值、網頁值、新值三個數字，真正要看的
        那一個反而被埋掉了。有話要說的時候才寫兩個：

            1,000 → 2,000        等著寫的，或這一輪剛寫進去的
            1,000（網頁 2,000）  跟網頁不一樣，但這一輪沒被寫（例如現金被擋）
            1,000                跟網頁一致，這一輪也沒動過

        舊值取的是這批網頁資料讀進來時 Excel 上原本的數字（self.before）。
        寫入成功之後 item["current"] 就是新數字了，沒有這份快照的話，畫面上
        再也看不到「被蓋掉之前是多少」。

        現金要 compare_web=False：它的 web 是今日淨收付、不是餘額，
        拿去跟 Excel 的餘額比是兩件不同的東西。
        """
        before = self.before.get((name, item["cell"]), item["current"])
        if item["will_write"]:
            return f"{show(before)} → {show(item['proposed'])}", False
        if not values_match(before, item["current"]):
            return f"{show(before)} → {show(item['current'])}", True
        if compare_web and item["web"] is not None and not values_match(item["current"], item["web"]):
            return f"{show(item['current'])}（網頁 {show(item['web'])}）", False
        return show(item["current"]), False

    def _cash_turned(self, item, before):
        """
        餘額這一輪是不是由正翻負。

        紅字是一個數字一個數字上的（負的才紅），所以這裡只剩「由正變負」這件事
        要另外講 —— 那是這一條上最該被看見的一次變化，而兩個各自上色的數字
        只說得出「現在是負的」，說不出「今天才變負的」。
        """
        old = to_num(before, None)
        new = to_num(item["proposed"] if item["will_write"] else item["current"], None)
        return old is not None and old >= 0 and new is not None and new < 0

    def fill_sync_tree(self):
        """左邊的名單與右邊的明細一起重畫。"""
        self._fill_people()
        self._fill_detail()

    def _summary(self, name):
        """一位交易人的濃縮狀態：(要寫幾格, 幾條提醒, 現金顯示值, 現金是不是負的)。"""
        items = self.proposals.get(name, [])
        writes = sum(1 for item in items if item["will_write"])
        warns = len(self.warnings.get(name, []))

        cash = next((item for item in items if item["kind"] == "cash"), None)
        if cash is None:
            return writes, warns, "", False
        value = cash["proposed"] if cash["will_write"] else cash["current"]
        number = to_num(value, None)
        return writes, warns, show(value), number is not None and number < 0

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
        # （見 _fill_detail），明明剛讀完、左邊也看得到人，右邊卻像什麼都沒讀到。
        # 剛開程式什麼都還沒讀的時候名單是空的，這行自己會落到 None，
        # 不必另外擋「還沒選過人」。
        if self.current_sheet not in names:
            self.current_sheet = names[0] if names else None

        need = 0
        for name in names:
            # 現金餘額現在是隱藏欄（見 _build_people），值照樣填 —— 隱藏的意思是
            # 「先不顯示」，不是「不算」，那一欄要回來的時候不必再回頭補這裡。
            writes, warns, cash, negative = self._summary(name)
            tags = []
            if writes:
                need += 1
                flag = f"要寫 {writes}"
                tags.append("attention")
            elif warns:
                flag = "⚠"
                tags.append("warned")
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
        # 否則每次重建名單都要多畫一張明細。
        if not picked or picked[0] == self.current_sheet:
            return
        self.current_sheet = picked[0]
        self._sync_scope_to_person()
        self._fill_detail()

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
        self._fill_detail()

    def _grouped(self, name):
        """
        把一格一列的提案併成畫面上的列：一檔股票一列，現金另外拿出來。

        以 Excel 的列號分組而不是股號 —— 畫面上的一列就是檔案上的同一列，
        對照的時候不必在心裡再翻譯一次。

        一個分頁只有第 4~8 列五個位置，一列一檔、股號不會重複。真的重複了的話
        這裡照樣會畫成兩列，但底下撐不住：兩列都查到同一檔的網頁值，
        股數/成本會整個複製貼到兩列，不會照原本的比例拆開。
        所以那不是一種支援的排法，是一種要避開的排法。
        """
        groups, order, cash = {}, [], None
        for item in self.proposals.get(name, []):
            if item["kind"] == "cash":
                cash = item
                continue
            if item["row"] not in groups:
                groups[item["row"]] = {}
                order.append(item["row"])
            groups[item["row"]][item["which"]] = item
        return [groups[row] for row in order], cash

    def _fill_detail(self):
        """右邊那張表：選中的交易人有哪幾檔、哪幾格要動。"""
        self.tree.delete(*self.tree.get_children())
        name = self.current_sheet
        groups, cash = self._grouped(name)

        # 高度照這位有幾檔算，但至少留五列。下限原本是八列 —— 那時候現金貼在這張表
        # 底下，表格一跳高度現金就跟著上下移動，寧可空著幾列也不要它動。現金搬到
        # 左邊並排之後（2026/08/21），它不再跟著這張表的高度走，下限就跟著改成
        # 「左邊那一欄有多高」：現金表三列加底下那顆按鈕，差不多就是這裡五列，
        # 兩邊收在同一段高度裡。上限只是防呆（真有人塞了三十檔，那就讓它捲）。
        self.tree.configure(height=max(5, min(len(groups), 18)))

        for group in groups:
            qty, cost = group.get("qty"), group.get("cost")
            both = [item for item in (qty, cost) if item is not None]
            if not both:
                continue

            texts, written = {}, False
            for which, item in (("qty", qty), ("cost", cost)):
                if item is None:
                    texts[which] = ""
                    continue
                texts[which], done = self._value_text(name, item)
                written = written or done

            tags = []
            if any(item["will_write"] for item in both):
                tags.append("write")
            elif written:
                tags.append("done")
            # 前景色一列只給一個 —— 同時掛兩個管顏色的標籤，最後誰贏是 Tk 的
            # 內部順序決定的，看起來就會時橘時灰。
            if any(item.get("missing") for item in both):
                tags.append("missing")

            # 說明欄拿掉了（2026/08/21，見 ui_layout.DETAIL_COLUMNS）——股票這邊
            # 唯一寫得出來的那句「網頁庫存已無此檔」，訊息框裡本來就有一條講得更
            # 清楚的提醒（第幾列、哪一檔、要不要刪）。這裡只剩「灰字＝網頁沒有它」。
            self.tree.insert(
                "", "end",
                values=(
                    stock_title(both[0]["label"]),
                    texts["qty"], texts["cost"],
                ),
                tags=tuple(tags),
            )

        # 欄寬照這一位的內容重算：股票名字有長有短（「2059 川湖」對「006208 富邦台50」），
        # 值有時是一個數字、有時是「104.6 → 10,400」，固定欄寬一定有人被切到。
        fit_to_content(self.tree)

        self._fill_cash(name, cash)
        self._fill_notes(name)

    def _refill_cash(self):
        """只重畫現金那張表。人跟提案都從現在的狀態拿，換算法時用得到。"""
        name = self.current_sheet
        self._fill_cash(name, self._cash_item(name))
        # 現金的說明現在寫在訊息框裡（見 _fill_cash），換算法整句話都會變，
        # 只重畫表格的話那句話會留在上一種算法的版本。
        self._fill_notes(name)

    def _fill_cash(self, name, item):
        """
        股票表旁邊那張現金表：一件事一列，欄位跟上面那張表對齊。

            現金                     金額
            現金餘額            893 → 655
            今日初始現金餘額          893
            現金算法             銀行餘額推算

        說明欄 2026/08/21 拿掉了（見 ui_layout.CASH_COLUMNS）：這幾列的說明
        ——「銀行餘額 + 淨收付(T+0) + 淨收付(T+1) = 893 - 238 + 0 = 655」——是整個畫面上
        最長的句子，擠在半個畫面寬的表格裡一定被切尾巴。它們收在 self.cash_notes
        裡，由 _fill_notes 寫進底下的訊息框，那裡會自動換行。

        沒資料的列就不畫（不是留一列空的）—— 換到還沒讀過的人時留著上一位的餘額，
        是這畫面上最危險的一種殘影。負的數字整列上紅字（理由見 ui_layout 建這張表
        那一段）。

        今日初始現金餘額直接讀紀錄檔的現金基準（見 ledger.opening_balance），不是
        提案算出來的 —— 它講的是「今天從多少錢開始」，跟這一輪要不要寫哪一格無關，
        就算這位今天一格都不必動也要看得到。基準每天由當天第一次登入設成 B8，
        所以正常情況它就是今天早上的那個數字。剛按過「修改」還沒落帳的時候寫成
        「舊 → 新」，跟餘額那一列同一個寫法：按完卻還顯示舊數字，看起來就像沒按到。

        現金算法這次執行還沒問過就不畫那一列。算法是這次程式開起來、第一次按
        「讀取全部帳戶」時跳視窗問的，在那之前寫一個名字上去，看起來就像已經選好了
        —— 而那時候顯示的其實是上次沿用下來的預設值。
        """
        rows = []

        if item is not None:
            before = self.before.get((name, item["cell"]), item["current"])
            after = item["proposed"] if item["will_write"] else item["current"]
            value = (show(after) if values_match(before, after)
                     else f"{show(before)} → {show(after)}")

            # 「餘額轉負」擺在說明最前面 —— 後面那句「今日淨收付…」每天都在，
            # 由正變負卻是難得一次，排在後面會被當成例行文字滑過去。
            note = item["note"]
            if self._cash_turned(item, before):
                note = f"餘額轉負；{note}" if note else "餘額轉負"
            rows.append(("balance", "現金餘額", value, note, _neg_tag(after)))

        if self.cash_method.get() == planner.METHOD_OPENING:
            rows.append(("opening", "今日初始現金餘額", self._opening_text(name, item), "",
                         _neg_tag(self._opening(name))))

        if self.path in self.cash_method_asked:
            # 不寫「（點一下換）」：那個入口本來就是給測試用的（.env 關得掉），
            # 在正式畫面上等於一句每天都在、卻幾乎不會用到的提示。滑鼠移過去
            # 游標會變手指（見 _on_cash_motion），要用的人找得到。
            rows.append(("method", "現金算法",
                         planner.METHOD_NAMES[self.cash_method.get()], "", "method"))

        # 說明留給訊息框，表格只放名字跟值。空的說明不進去 —— 訊息框裡一行
        # 「今日初始現金餘額：」後面什麼都沒有，比不寫還難懂。
        self.cash_notes = [(title, note) for _key, title, _value, note, _tag in rows if note]

        self.cash_tree.delete(*self.cash_tree.get_children())
        for key, title, value, note, tag in rows:
            self.cash_tree.insert("", "end", iid=key, values=(title, value),
                                  tags=(tag,) if tag else ())
        # 高度照實際有幾列給，多的空白留給底下的訊息框。一列都沒有的時候還是留 1
        # ——Treeview 高度 0 在畫面上只剩一條表頭，看起來像壞掉。
        self.cash_tree.configure(height=max(len(rows), 1))
        # 欄寬照內容重算。這張表的列數會變（今日初始那一列只有一種算法有、
        # 現金算法那一列問過才有），最長的字跟著換，所以每次填完都要重算一次。
        fit_to_content(self.cash_tree)

        self.opening_ready = (self.ledger is not None and item is not None
                              and not item["blocked"]
                              and self.cash_method.get() == planner.METHOD_OPENING)
        self._show_opening_row(self.cash_method.get() == planner.METHOD_OPENING)
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

    def _on_cash_click(self, event):
        """
        現金表上點到「現金算法」那一列就換另一種算法（會先跳確認視窗）。
        點別列不做事 —— 這張表其他列是純顯示。
        """
        if self.cash_tree.identify_row(event.y) == "method":
            self._toggle_cash_method()

    def _on_cash_motion(self, event):
        """滑過「現金算法」那一列時游標變手指，不然沒有東西看得出那一列點得動。"""
        want = "hand2" if self.cash_tree.identify_row(event.y) == "method" else ""
        # 每次滑鼠移動都設一次的話，Tk 會不停重設游標，滑過表格時會閃。
        if self.cash_tree.cget("cursor") != want:
            self.cash_tree.configure(cursor=want)

    def _show_opening_row(self, show_it):
        """
        用銀行餘額推算的日子，「修改今日初始現金餘額」那顆收起來，不占畫面。

        它改的是初始餘額累加那一種算法的基準，今天既然不用那種算法，
        擺著只是佔位置。基準本身不會消失，只是沒顯示：每天照樣設，
        明天切回初始餘額累加就要用它。
        """
        if show_it == bool(self.opening_row.winfo_manager()):
            return
        if show_it:
            self.opening_row.grid()
        else:
            self.opening_row.grid_remove()

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
                                      ledger_mod.opening_balance(cash), item)
        if opening is None:
            return

        planner.apply_cash_reset(item, opening)
        # 這顆按鈕動到的只有眼前這一位，寫入與落帳的範圍就跟著縮到他身上
        # —— 名單上別人那些「要寫」的格子是上一輪算的，不該被這一下順手寫出去。
        self.round_scope = {name}
        self.fill_sync_tree()

        writes, total = self._collect_writes()
        if total:
            self._begin_write(writes, total)
            return

        # 算出來剛好等於 Excel 上的數字，一格都不必寫 —— 但基準確實被改掉了，
        # 這一筆不落帳就等於沒按過（見 _commit_round）。
        recorded = self._commit_round()
        self.replan()
        self.refresh_history()
        self._say(f"{name} 的今日初始現金餘額改成 {show(opening)}，"
                  "跟 Excel 上的數字剛好一樣，沒有格子要寫。"
                  + (f"紀錄檔更新了 {recorded} 筆（見歷程）。" if recorded else ""))

    def _head_line(self, name):
        """
        「這份資料是誰的、什麼時候讀的」那一句：`簡嘉懋　讀取於 19:54:23`。

        原本是表格上方一行獨立的粗體標頭，2026/08/21 使用者要求收進訊息框
        當第一行（見 _fill_notes、ui_layout._build_detail）。名字一定要跟資料
        走在一起 —— 換人的時候第一眼要確認「換對了沒」，而畫面上另一個寫著名字
        的地方是左邊名單那個反白，那講的是「我點了誰」，不是「右邊這些數字是誰的」。

        讀取時間精準到秒，不到分鐘——分鐘不夠精準的話，剛讀完跟半小時前讀的
        兩個時間點會長得一樣。

        還沒讀過任何東西的時候整句留白（2026/08/21）——空的名單加空的表格，
        自己就已經在說「還沒有資料」了，再寫一句只是把同一件事講第二次；
        「然後要按哪顆」開機時狀態列已經講了（見 ui.py）。
        """
        if not name:
            return ""

        parts = [name + ("（模擬）" if name in self.fake_sheets else "")]
        read = self.read_at.get(name)
        parts.append(f"讀取於 {read:%H:%M:%S}" if read else "尚未讀取")
        return "　".join(parts)

    def _fill_notes(self, name):
        """
        第一行是「這份資料是誰的、什麼時候讀的」（`簡嘉懋　讀取於 19:54:23`，
        見 _head_line）：2026/08/21 使用者要求把原本表格上方那行標頭也收進來，
        以後有讀取動作就固定放在這個框的第一行。往下是原本兩張表「說明」欄的
        內容（同一天整批搬過來，見 ui_layout.DETAIL_COLUMNS），最後才是提醒，
        而且提醒只在「有事」的時候出聲。
        現金那幾句排在提醒前面：「這個數字是怎麼算出來的」是每天都要看一眼的事，
        提醒則是有事才有，排前面會把每天都在的那句話往下擠。已經被 planner 當成
        提醒送出來的那一句（被擋住不寫的理由，「[現金] …」）不再重複一次。

        同一天拿掉的是另一行「簡嘉懋　現金 655　跟網頁一致」（也是使用者要求）：
        現金在現金表第一列、要寫幾格是左邊名單上那個旗標（「要寫 3」／「✓」，
        見 _fill_people），同一件事講兩次，只會讓每天都要讀的那幾句算式往下掉。

        20 個人的提醒全部堆在同一個框裡等於沒有提醒 —— 別人的事左邊名單上
        已經用 ⚠ 標出來了，要看就換過去看。整組失敗（problems）例外，
        那跟選中誰無關，一定要講。

        「操作中不要改 Excel 的現金餘額」不擺在這裡 —— 這個框在讀取完、寫入完
        都還留著上一輪的內容，常駐在這裡等於讀完老半天還在講「正在做的時候
        別碰」，反而誤導。那句話改成讀取中的忙碌訊息（見 start_fetch），
        跟著狀態列一起出現、一起被下一句蓋掉。
        """
        text = []
        warns = list(self.warnings.get(name, []))
        if name:
            head = self._head_line(name)
            if head:
                text.append(head)
            text += [f"{title}：{note}" for title, note in self.cash_notes
                     if f"[現金] {note}" not in warns]

        text += warns
        for problem in self.problems:
            text.append(f"⚠ {problem}")

        self.warn_box.configure(state="normal")
        self.warn_box.delete("1.0", "end")
        self.warn_box.insert("1.0", "\n".join(text))
        self.warn_box.configure(state="disabled")

    def clear_notes(self):
        """
        把訊息框清空。換現金算法的時候呼叫（見 _set_method）。

        裡面每一句都是點「讀取」那一刻算出來的，而換算法換掉的正是「現金那個
        數字怎麼來」：算式、被擋住的理由都可能不再成立，而新的那一種可能根本還沒
        去查資料（銀行餘額與交割金額只在選了那種算法的那一輪才查）。留著舊的
        比什麼都不寫更危險 —— 那幾句看起來跟畫面上現在的數字是一套的。

        下一次讀取（或者在名單上換一位）就會連同新算法重新填好。
        """
        self.cash_notes = []
        self.warn_box.configure(state="normal")
        self.warn_box.delete("1.0", "end")
        self.warn_box.configure(state="disabled")

    def _sync_buttons(self):
        """上面那兩顆能不能按。畫面上會變灰的按鈕現在只剩它們。"""
        # Excel 沒開著就不給登入，也不給讀取 —— 讀取自己會順便登入，只擋登入的話
        # 這道關卡按另一顆按鈕就繞過去了。擋在最前面的理由是後面每一步都要 Excel：
        # 讀完要拿它的現值算提案，寫入更是直接改它。
        ready = self.excel_open and not self.busy
        self.login_button.configure(state="normal" if ready else "disabled")
        self.fetch_button.configure(state="normal" if ready else "disabled")
        self._apply_scope_state()
        # 「修改」不看 Excel 開著沒 —— 它改的是紀錄檔裡的基準，要寫 Excel 的時候
        # 寫入那邊自己會把檔案開起來。能不能按只看「這一位有沒有網頁資料」，
        # 那是 _fill_cash 判的。用銀行餘額推算的日子整顆被藏起來
        # （_show_opening_row），這裡照樣設 state —— 藏著的按鈕設得動，
        # 放回來的時候就已經是對的狀態了。
        self.opening_button.configure(
            state="normal" if self.opening_ready and not self.busy else "disabled")
