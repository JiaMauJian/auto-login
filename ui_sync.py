"""同步分頁的顯示與操作：左邊名單、右邊明細、現金那一條。"""

import tkinter as tk

import ledger as ledger_mod
import planner
from util import show, to_num, values_match
from ui_common import ask_opening_balance


def _neg_tag(value):
    """負的數字才上紅字。空的、不是數字的都不算。"""
    number = to_num(value, None)
    return "neg" if number is not None and number < 0 else None


def stock_title(label):
    """從「股數（2059 川湖）」取出「2059 川湖」。取不到就原樣顯示。"""
    inside = label.partition("（")[2].rstrip("）")
    return inside or label


def group_note(qty, cost):
    """
    一列的說明。兩格講的是同一件事就寫一次，否則各自標明是哪一格。

    只有一格有話說的時候也要標 —— 「網頁庫存已無此檔」沒說是股數還是成本，
    等於沒說。
    """
    notes = [(which, item["note"])
             for which, item in (("股數", qty), ("成本", cost))
             if item is not None and item["note"]]
    if not notes:
        return ""
    present = len([item for item in (qty, cost) if item is not None])
    if len(notes) == present and len({note for _which, note in notes}) == 1:
        return notes[0][1]
    return "；".join(f"{which}：{note}" for which, note in notes)


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
        if self.current_sheet not in names:
            self.current_sheet = names[0] if names else None

        need = 0
        for name in names:
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
                values=(cash, flag), tags=tuple(tags),
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

        # 高度照這位有幾檔算，但至少留八列 —— 每換一個人表格就跳一次高度、
        # 底下的現金跟著上下移動，比空幾列還難看，所以下限抓在「手上大概會有
        # 幾檔」而不是「這一位現在有幾檔」；上限只是防呆（真有人塞了三十檔，
        # 那就讓它捲）。
        self.tree.configure(height=max(8, min(len(groups), 18)))

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

            self.tree.insert(
                "", "end",
                values=(
                    " ".join(item["cell"] for item in both),
                    stock_title(both[0]["label"]),
                    texts["qty"], texts["cost"],
                    group_note(qty, cost),
                ),
                tags=tuple(tags),
            )

        self._fill_head(name)
        self._fill_cash(name, cash)
        self._fill_opening(name, cash)
        self._fill_notes(name)

    def _fill_cash(self, name, item):
        """
        表格底下那一條現金：現金餘額 B8　舊值 → 新值　說明：…

        一段一段插，負的數字自己上紅字。沒資料就整條清掉 ——
        換到還沒讀過的人時留著上一位的餘額，是這畫面上最危險的一種殘影。
        """
        segments = []
        if item is not None:
            before = self.before.get((name, item["cell"]), item["current"])
            after = item["proposed"] if item["will_write"] else item["current"]
            turned = self._cash_turned(item, before)

            segments.append((f"現金餘額 {item['cell']}", "dim"))
            segments.append(("　", None))
            segments.append((show(before), _neg_tag(before)))
            if not values_match(before, after):
                segments.append(("　→　", None))
                segments.append((show(after), _neg_tag(after)))

            # 「餘額轉負」擺在說明最前面 —— 後面那句「今日淨收付…」每天都在，
            # 由正變負卻是難得一次，排在後面會被當成例行文字滑過去。
            note = item["note"]
            if note or turned:
                segments.append(("　　說明：", "dim"))
                if turned:
                    segments.append(("餘額轉負", "turned"))
                    if note:
                        segments.append(("；", None))
                if note:
                    segments.append((note, None))

        self.cash_line.configure(state="normal")
        self.cash_line.delete("1.0", "end")
        for text, tag in segments:
            self.cash_line.insert("end", text, (tag,) if tag else ())
        self.cash_line.configure(state="disabled")
        # 高度要等版面算完寬度才知道換不換行，所以排到 idle 再量。
        self.cash_line.after_idle(self._fit_cash_line)

    def _fit_cash_line(self):
        """
        讓那一條剛好包住內容。說明偶爾會長到換行，固定一列會被切掉。

        高度沒變就什麼都不做 —— 這個函式也綁在 <Configure> 上，每改一次高度就是
        一次新的 Configure，照改不誤的話兩個值會互相觸發、來回抖個不停。
        """
        try:
            lines = self.cash_line.count("1.0", "end", "displaylines")[0]
        except (tk.TclError, TypeError, IndexError):
            lines = 1
        height = max(1, min(lines, 3))
        if height != int(self.cash_line["height"]):
            self.cash_line.configure(height=height)

    def _fill_opening(self, name, item):
        """
        現金那一條底下那一行：今日初始現金餘額，加一顆改它的按鈕。

        數字直接讀紀錄檔的現金基準（見 ledger.opening_balance），不是提案算出來的
        —— 它講的是「今天從多少錢開始」，跟這一輪要不要寫哪一格無關，就算這位
        今天一格都不必動也要看得到。基準每天由當天第一次登入設成 B8，所以正常
        情況它就是今天早上的那個數字。

        剛按過「修改」還沒落帳的時候寫成「舊 → 新」，跟上面那一條同一個寫法：
        按完卻還顯示舊數字，看起來就像沒按到。

        整組（數字＋「修改」）只在用「初始餘額累加」時顯示（見 _show_opening_row）
        —— 銀行餘額推算的日子每次讀取直接算好寫回 B8，這組今天用不到。
        """
        cash = self.ledger.sheet(name)["cash"] if (self.ledger is not None and name) else None
        opening = ledger_mod.opening_balance(cash) if cash is not None else None

        # 一位都還沒選（還沒讀過網頁資料）時寫破折號而不是「還沒設定」——
        # 那時候是「不知道要看誰」，不是「這個人沒有基準」。
        text = "—" if not name else ("(還沒設定)" if opening is None else show(opening))
        if item is not None and item["reset_to"] is not None:
            text = f"{text} → {show(round(item['reset_to'] - item['net'], 2))}"

        number = to_num(opening, None)
        self.opening_value.configure(
            text=text, foreground=self.colors.danger if number is not None and number < 0 else "")
        self.opening_ready = (self.ledger is not None and item is not None
                              and not item["blocked"]
                              and self.cash_method.get() == planner.METHOD_OPENING)
        self._show_opening_row(self.cash_method.get() == planner.METHOD_OPENING)
        self._sync_buttons()

    def _show_opening_row(self, show_it):
        """
        用銀行餘額推算的日子，基準數字跟「修改」整組收起來，不占畫面。

        它們講的是初始餘額累加那一種算法的基準，今天既然不用那種算法，
        擺著只是佔位置。數字本身不會消失，只是沒顯示：基準每天照樣設，
        明天切回初始餘額累加就要用它。
        """
        if show_it == bool(self.opening_row.winfo_manager()):
            return
        if show_it:
            self.opening_row.pack(side="left")
        else:
            self.opening_row.pack_forget()

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
                  f"{item['cell']} 上的數字剛好一樣，沒有格子要寫。"
                  + (f"紀錄檔更新了 {recorded} 筆（見歷程）。" if recorded else ""))

    def _fill_head(self, name):
        """表格上方那一行：是誰、第幾位、現金多少、這次要寫幾格、資料是幾點讀的。"""
        if not name:
            self.detail_head.configure(text="還沒有資料 —— 按上面的「讀取網頁資料」")
            return

        writes, _warns, cash, _negative = self._summary(name)
        names = self._shown()
        parts = [name + ("（模擬）" if name in self.fake_sheets else "")]
        if name in names:
            parts.append(f"第 {names.index(name) + 1} / {len(names)} 位")
        if cash:
            parts.append(f"現金 {cash}")
        parts.append(f"要寫 {writes} 格" if writes else "跟網頁一致")
        # 一次只更新一位之後，畫面上每個人的資料新舊不一 —— 沒有這個時間，
        # 半小時前讀的數字跟剛剛讀的長得一模一樣。
        read = self.read_at.get(name)
        if read:
            parts.append(f"讀取於 {read:%H:%M}")
        self.detail_head.configure(text="　".join(parts))

    def _fill_notes(self, name):
        """
        提醒只講選中的這一位，而且只在「有事」的時候出聲。

        20 個人的提醒全部堆在同一個框裡等於沒有提醒 —— 別人的事左邊名單上
        已經用 ⚠ 標出來了，要看就換過去看。整組失敗（problems）例外，
        那跟選中誰無關，一定要講。

        沒有警告的時候框就留空 —— 「還沒有資料」「已經一致」這種話下面
        狀態列都講過（見 ui.py／ui_background.py 的 _say），這裡再寫一次
        只是同一件事說兩遍，不是真正的提醒。
        """
        text = [f"• {warning}" for warning in self.warnings.get(name, [])]
        for problem in self.problems:
            text.append(f"⚠ {problem}")

        # 每天都成立的規矩，擺在最後一行。這個框只有五行高又沒有捲軸，
        # 常駐的字排在前面就會把當天真正的警告推出視線 —— 排最後的話，
        # 沒事的日子看得到（框是空的），有事的日子被擠掉的正好是它。
        if self.proposals:
            text.append("操作中不要改 Excel 的現金餘額")

        self.warn_box.configure(state="normal")
        self.warn_box.delete("1.0", "end")
        self.warn_box.insert("1.0", "\n".join(text))
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
        # 那是 _fill_opening 判的。用銀行餘額推算的日子整組被藏起來
        # （_show_opening_row），這裡照樣設 state —— 藏著的按鈕設得動，
        # 放回來的時候就已經是對的狀態了。
        self.opening_button.configure(
            state="normal" if self.opening_ready and not self.busy else "disabled")
