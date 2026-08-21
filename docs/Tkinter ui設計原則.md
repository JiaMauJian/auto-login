# Python tkinter UI 設計原則參考文件

> 用途：作為 AI 助理在協助設計 / 撰寫 tkinter 介面程式碼時的參考準則。

> **使用說明**：本文件為「原則庫」，請依專案規模與需求選用，不需要每次全部套用。
>
> - 小工具／單一視窗小程式：優先參考第一~四章（版面配置、視覺體驗、進階套件、架構建議）即可
> - 涉及耗時操作、多視窗、需打包發布等情境：再參考對應章節（第五~十章）
>   目標是寫出乾淨可維護的介面，而不是為了套用規則而增加不必要的複雜度。

---

## 一、版面配置（Layout）

### 1. 優先使用 `grid`，避免使用 `pack`

- `grid` 適合較複雜的表單、多欄位介面，控制精確
- `pack` 適合簡單的堆疊排列
- **同一個容器（Frame/Window）內不要同時混用 `pack` 和 `grid`**，會造成版面錯亂或程式錯誤

### 2. 善用 `sticky` 與 `weight` 讓視窗可正常縮放

```python
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)
widget.grid(row=0, column=0, sticky="nsew")
```

- `weight` 決定該欄/列在視窗縮放時分配到的比例
- `sticky="nsew"` 讓元件隨容器縮放而延展（north/south/east/west）

### 3. 用 Frame 分區塊

將介面拆成多個 `Frame`，各自負責一個區域，避免全部元件塞在同一層：

- 頂部工具列（toolbar）
- 左側選單（sidebar）
- 主內容區（main content）
- 底部狀態列（status bar）

### 4. 視窗最小尺寸

- 用 `root.minsize(width, height)` 設定最小可縮放尺寸，避免使用者把視窗縮太小導致版面跑掉

---

## 二、視覺與體驗

| 項目     | 建議做法                                                                                                                          |
| -------- | --------------------------------------------------------------------------------------------------------------------------------- |
| 間距     | `padx`、`pady` 全專案統一用固定幾組數值（例如 5、10、20），避免每個元件間距不一致                                                 |
| 字型     | 用 `tkinter.font` 定義好幾組字型物件（標題／內文／按鈕），全域套用，避免各處硬寫字型字串                                          |
| 元件選用 | 盡量用 **ttk**（`tkinter.ttk`）取代原生 tkinter 元件（Button、Entry、Combobox 等），外觀較現代                                    |
| 主題     | 可套用 ttk 的 Style/Theme，或系統主題，避免介面過於「90 年代感」                                                                  |
| DPI 縮放 | 高解析度螢幕下字型/元件可能顯示過小，需留意 DPI awareness 處理（尤其 Windows 上常見），必要時搭配作業系統設定或程式內縮放係數調整 |

---

## 三、進階美化套件（原生 tkinter 不夠用時）

- **ttkbootstrap**：基於 ttk，提供 Bootstrap 風格的現代化主題，社群推薦度高
- **customtkinter**：提供圓角、深色模式等現代 UI 元件

> 使用第三方套件前，需確認專案環境（Python 版本、是否可安裝額外套件）是否允許。

---

## 四、架構建議

### 1. 分層設計（類似 MVC）

- UI 層只負責「顯示畫面」與「收集使用者輸入」
- 商業邏輯（資料處理、資料庫存取等）不要寫在按鈕的 callback 函式裡，應獨立成邏輯層/服務層
- 方便日後維護、測試，也方便 UI 改版不影響邏輯

### 2. 表單輸入驗證

- 建議做「即時提示」，例如：
  - 欄位資料格式錯誤時邊框變紅
  - 顯示提示文字告知錯誤原因
- 能提升使用者體驗，減少送出後才發現錯誤的情況

---

## 五、執行緒與介面卡頓（Responsiveness）

- tkinter 是**單執行緒事件迴圈**，長時間運算（檔案處理、網路請求、資料庫查詢等）**不能直接在主執行緒執行**，否則畫面會凍結無回應
- 建議做法：
  - 用 `threading` 開背景執行緒處理耗時工作，搭配 `queue.Queue` 回傳結果給主執行緒
  - 主執行緒用 `after()` 定時輪詢 queue，取得結果後再更新 UI
  - **不要在 worker thread 直接操作 tkinter 元件**（tkinter 元件非執行緒安全），所有 UI 更新都必須在主執行緒進行

---

## 六、選單列與快捷鍵

- 使用 `Menu` 元件建立頂層選單列（`root.config(menu=menubar)`）與右鍵選單（`Menu(tearoff=0)` + `bind("<Button-3>", ...)`）
- 常用操作建議綁定快捷鍵：
  - 用 `bind_all()` 綁定全域快捷鍵（例如 Ctrl+S 存檔、Esc 關閉視窗）
  - 注意 Windows／macOS 快捷鍵慣例不同（Ctrl vs Cmd），跨平台專案需留意

---

## 七、圖片處理

- tkinter 原生 `PhotoImage` 只支援 GIF/PGM/PPM，若要用 **PNG/JPG** 需搭配 `Pillow`（`from PIL import Image, ImageTk`）
- **常見坑**：圖片物件（`PhotoImage` / `ImageTk.PhotoImage`）若沒有保留參照（reference），會被 Python 垃圾回收，導致畫面上圖片消失。做法：
  - 把圖片物件存成 `self.img = ...`（掛在物件上），或存進一個 list/dict 統一管理，避免變成區域變數被回收

---

## 八、錯誤處理與使用者提示

- 統一用 `messagebox`（`showerror` / `showwarning` / `showinfo`）呈現錯誤或提示，不要用 `print` 或讓程式直接崩潰
- 可能出錯的操作（檔案讀寫、網路連線、資料庫存取等）建議用 `try/except` 包裝，並在 `except` 中給使用者清楚、口語化的錯誤訊息（避免直接把例外訊息原封不動丟給使用者）

---

## 九、視窗生命週期管理

- 開啟子視窗（`Toplevel`）時，注意 **modal**（用 `grab_set()` 讓子視窗獨佔操作焦點）與**非 modal**（可與主視窗同時操作）的差異，依需求選用
- 關閉視窗（`WM_DELETE_WINDOW` 事件）時，記得處理資源釋放：
  - 關閉資料庫連線
  - 停止背景執行緒（例如設定停止旗標，等待 thread join）
  - 儲存未存檔的變更前先提示使用者

---

## 十、打包發布

- 常見打包工具：`PyInstaller`、`cx_Freeze`
- 注意事項：
  - 跨平台打包差異（Windows / macOS / Linux 需個別打包，不能一次打包多平台）
  - 圖片、設定檔等資源檔的路徑處理，打包後路徑會改變，建議用 `sys._MEIPASS`（PyInstaller）或相對於執行檔的路徑方式讀取資源，而非寫死開發環境的路徑

### ttkbootstrap + PyInstaller 常見坑

1. **主題資源檔沒被打包進去**：PyInstaller 預設只分析 import 的 `.py` 程式碼，不會自動抓套件內的非 Python 資料檔（主題定義、圖示等），導致打包後主題跑掉或 `FileNotFoundError`
   - 解法：打包時加 `--collect-all ttkbootstrap`
     ```bash
     pyinstaller --collect-all ttkbootstrap your_app.py
     ```
   - 或在 `.spec` 檔加：
     ```python
     from PyInstaller.utils.hooks import collect_all
     datas, binaries, hiddenimports = collect_all('ttkbootstrap')
     ```

2. **Hidden Imports 問題**：ttkbootstrap 內部有動態載入的模組（如 `ttkbootstrap.dialogs`、`ttkbootstrap.widgets`、`ttkbootstrap.icons`），若程式碼未明確 import，靜態分析會漏掉，導致執行期 `ModuleNotFoundError`
   - 解法：通常 `--collect-all` 會一併解決，或手動在 spec 檔加 `hiddenimports=['ttkbootstrap']`

3. **`--onefile` vs `--onedir` 的取捨**：
   - `--onefile`：單一執行檔方便發布，但啟動較慢（每次要先解壓到暫存目錄），主題資源多時更明顯
   - `--onedir`：啟動快，但發布時是整個資料夾
   - 常駐執行的後台管理工具，建議先用 `--onedir` 觀察，真的需要單檔再換 `--onefile`

4. **Windows 防毒誤判**：`--onefile` 打包的執行檔常被防毒軟體誤判為惡意程式（解壓行為模式類似封裝惡意程式），內部工具部署前建議先跟資安/IT 溝通白名單，或改用 `--onedir` 降低誤判機率

**建議打包指令範例**：

```bash
pyinstaller --name MyApp --onedir --collect-all ttkbootstrap --icon=app.ico your_app.py
```

---

## 十一、ttkbootstrap 自訂樣式名稱的命名規則

- ttkbootstrap 元件（`ttk.Label`、`ttk.Treeview`…）建構子收到 `style="Xxx.TLabel"` 這種自訂名稱時，會先檢查這個名稱是否「已登記」在它自己的樣式清單（`Style.style_exists_in_theme`）：
  - 沒登記過 → 當成 bootstyle 字串重新解析，解析不出已知色系/家族時，**會整個丟回該元件的預設樣式**（`Xxx` 前綴整個消失），不會報錯也不會警告
  - 已登記過 → 直接沿用你給的名稱
- **只呼叫 `style.map(name, ...)` 不會登記這個名稱**；`style.map` 只是把 state 對應的顏色寫進 Tcl 樣式表，不會讓 ttkbootstrap 知道這個名字存在
- **只有 `style.configure(name, ...)` 會登記**（哪怕不帶任何參數，`style.configure(name)` 空呼叫也算數）
- **规则**：自訂樣式名稱要在建立對應元件之前，先呼叫一次 `style.configure(name)`（即使沒有要 configure 的屬性），再呼叫 `style.map(name, ...)` 設定 state 相關顏色（例如選取列底色）：

  ```python
  # 錯誤：只 map，元件建構時會被悄悄換回預設的 Treeview 樣式（灰底選取）
  ttk.Style().map("History.Treeview",
                  background=[("selected", colors.primary)])
  tree = ttk.Treeview(frame, style="History.Treeview")

  # 正確：先 configure() 登記名稱，再 map() 設定 state 顏色
  ttk.Style().configure("History.Treeview")
  ttk.Style().map("History.Treeview",
                  background=[("selected", colors.primary)],
                  foreground=[("selected", colors.selectfg)])
  tree = ttk.Treeview(frame, style="History.Treeview")
  ```

- 這個坑不容易發現，因為：
  - 不會報錯、不會警告，畫面看起來「有套用但顏色不對」，很容易誤以為是焦點（focus）造成的灰色，或以為顏色設錯
  - `widget.cget("style")` 可以拿來驗證：如果印出來的不是你設的名稱（例如印出 `"Treeview"` 而不是 `"History.Treeview"`），就是這個問題
- 專案裡其他自訂樣式（`Hint.TLabel`、`Method.TLabel`、`Auto.TLabel`、`Manual.TLabel`、`Choice.TRadiobutton`）都是用 `style.configure(...)` 建立，本來就沒有這個問題；只有需要「依 state（例如選取列）變色」而只呼叫 `style.map()` 的情境才會踩到

---

## 十二、Treeview 要畫格線（Excel 那種灰線）

ttk 的 `Treeview` **沒有**任何「畫格線」的選項。實測（Tk 8.6.12 + ttkbootstrap 2.2.2）確認以下三條路都不通：

| 試過的做法                                    | 結果                                                     |
| --------------------------------------------- | -------------------------------------------------------- |
| `style.layout("Treeview.Separator", …)`       | `Layout Treeview.Separator not found`，Tk 根本沒這個 layout |
| `style.configure("X.Treeview.Cell", relief="solid", borderwidth=1)` | 設得下去，畫面上什麼都沒有（padding 元素不畫框）        |
| `fieldbackground` 設成線色                    | 只是整片底色，被每一列的 row 底色蓋掉                    |

唯一可行的是 **image element**：做一張 1 像素寬的透明 `PhotoImage`，把要當線的那一條塗成灰色，用 `element_create(..., "image", ...)` 掛進 `Row` / `Cell` / `Item` 的 layout。

```python
line = style.colors.border                      # 用主題自己的灰，跟外框同一個顏色
row_img = tk.PhotoImage(width=1, height=rowheight)
row_img.put(line, to=(0, rowheight - 1, 1, rowheight))   # 只塗最底下那一列 → 橫線
col_img = tk.PhotoImage(width=1, height=rowheight)
col_img.put(line, to=(0, 0, 1, rowheight))               # 整條塗滿 → 直線
self._grid_line_images = (row_img, col_img)     # ★ 一定要留參照，見第七章

style.element_create("grid.row", "image", row_img, border=(0, 0, 0, 1), sticky="nswe")
style.element_create("grid.col", "image", col_img, border=(0, 1, 0, 1), sticky="ns")

style.layout("Treeview.Row", [
    ("Treeitem.row", {"sticky": "nswe", "children": [("grid.row", {"sticky": "nswe"})]})])
style.layout("Treeview.Cell", [
    ("grid.col", {"side": "right", "sticky": "ns"}),
    ("Treedata.padding", {"sticky": "nswe", "children": [
        ("Treeitem.text", {"sticky": "nswe"})]})])
```

幾個一定要注意的點：

- **`border` 是 9-patch 的固定邊**（left, top, right, bottom）。橫線給 `(0, 0, 0, 1)`：拉伸時只拉中間，最底下那 1 像素原樣不動。直線給 `(0, 1, 0, 1)`，高度才跟著格子拉伸（表頭比資料列矮，不給的話壓不到底）。
- **要掛成 `Treeitem.row` 的「子元素」，不要取代它**。`Treeitem.row` 負責畫底色（tag 的底色、選取列的藍），取代掉就整片不見了。圖除了那條線以外是透明的，底色照樣透出來。
- **`show="tree headings"` 的第一欄（`#0`）走的是 `Treeview.Item` 不是 `Treeview.Cell`**，要另外補一份，不然只有那一欄沒有右邊那條線。
- **表頭（`Treeview.Heading`）的直線補不上去，別試**。Tk 給表頭跟給資料格的方框差 4 像素，兩段線接不起來，看起來是一節一節歪的；線排在 layout 的哪個位置、`Treeview.Cell` 的 `padding` 怎麼調都改不掉（實測 `padding=0`、`(4,0,0,0)`、`(4,2,0,2)` 全都差 4）。量法：`tree.bbox(iid, col)` 拿資料格的方框，再用 `ImageGrab` 抓圖數線落在哪個 x。表頭本來就有自己的底線跟資料分開，直線從第一列才開始並不奇怪。
- 改的是 `Treeview.Row` / `Treeview.Cell` 這種**基底樣式的 layout**，全部 Treeview 一次到位，不必每張表各自設定。
- 空白列（資料列數少於表格高度）畫不出線 —— Treeview 只畫存在的列，這點跟 Excel 不一樣，改不了。

> **這個專案實際跑過之後又拿掉了（2026/08/21）**：畫面對，但**效能很差** —— 每一列、每一格都多一個 image element 要合成，捲動與重畫明顯變鈍（持股同步的表格不大都感覺得出來，20 個帳號的名單更明顯）。ttk 的 image element 沒有快取，每次重畫都是一次縮放合成。要格線的話，先確認表格夠小，或改用**斑馬紋**：`tree.tag_configure("stripe", background="#f2f2f2")`，填資料時奇數列掛上這個 tag （本專案改用這個做法，見 `ui_common.stripe`）。它只是換底色，沒有額外的元素要畫，重畫成本跟沒有它一樣；ttkbootstrap 的 `Tableview` 內建的 `stripecolor=("#f2f2f2", None)` 也是同一件事 —— 但 `Tableview` 是另一個元件（自帶搜尋列、分頁、排序，欄寬與列數也由它管），已經用 `ttk.Treeview` 寫好的表格不值得為了斑馬紋整組換過去，自己掛 tag 就有一樣的效果。

> 一個要注意的地方：**已經有底色的列不要再加斑馬紋**。那些底色通常在講事情（這一列要寫、剛寫過、這一位要處理），被斑馬紋蓋掉就沒了；而一列同時掛兩個管底色的 tag，誰贏是 Tk 的內部順序決定的，看起來會時有時無。

---

## 十三、Treeview 沒有「照內容自動調欄寬」

`ttk.Treeview` 的欄位只有三個旋鈕，三個都跟內容無關：

| 選項                | 實際作用                                                     |
| ------------------- | ------------------------------------------------------------ |
| `width`             | 目前欄寬，寫死的數字，不會因為內容變長而變                   |
| `minwidth`          | 使用者拖欄位邊界時的下限，同樣跟內容無關                     |
| `stretch`           | 表格變寬時這一欄要不要分到多出來的空間，不是「照字長大」     |

沒有 `resizeColumnsToContents` 這種東西（Tk 8.6.12 + ttkbootstrap 2.2.2 實測）。字比欄位寬的時候 **ttk 直接切，不補省略號**，而 Treeview 又沒有橫向捲軸，切掉的部分捲也捲不出來 —— 畫面上只會看到一個斷在半個字中間的數字，沒有任何跡象說它被切過。

要照內容調就得自己量。本專案的做法在 `ui_common.fit_to_content` / `fit_columns`（同步分頁那兩張表），有三個地方是踩過才知道的：

**1. 欄寬要補的不只是 padding，還有 ttk 自己吃掉的邊**

一欄要多寬才不切字：

```
欄寬 = font.measure(要顯示的字) + 2 × (Treeview.Cell 的 padding + 4)
```

那個 `+4` 是每一格左右各自被 `Treeitem.text` 吃掉的邊，樣式調不掉。實測（12 級 Microsoft JhengHei UI、`padding=(8, 0)`）：`"104.6 → 10,400"` 量出來 113px，欄寬要 **137px** 字才畫得完整，而 113 + 8 + 8 只有 129 —— 少的 8px 剛好切掉最後一個字的一半。

量法：`tree.bbox(iid, col)` 拿格子的方框，`ImageGrab.grab` 抓那一塊，數哪幾個 x 有墨跡。欄寬夠的時候墨跡寬度會停在一個定值（等於 `measure` 減掉字的邊距），不夠的時候會跟著欄寬一起縮 —— 那個轉折點就是最小不切欄寬。

**2. 量測不能放進 `<Configure>`**

`font.measure()` 是一次 Tcl 呼叫，**一個字串約 0.2ms**（跟字長幾乎無關，貴的是跨語言呼叫本身）。拖分隔線時 `<Configure>` 一秒噴幾十次，量測放進去就是每次重畫都重量整張表。

正確的切法是把兩件事分開：**內容變了才量**（填完表叫一次，結果記在 widget 上），`<Configure>` 只拿現成的數字重攤。實測本專案：填完表重量 3.13ms／次，`<Configure>` 那條路 0.001ms／次。

| 表格                            | 要量幾個字串 | 一次要多久 |
| ------------------------------- | ------------ | ---------- |
| 5 檔 × 3 欄（持股明細）         | 15           | ~3ms       |
| 60 列 × 6 欄（歷程「今天」）    | 360          | ~70ms      |
| 2000 列 × 6 欄（歷程「全部」）  | 12,000       | ~2.4 秒    |

列數沒有上限的表格別直接套 —— 要嘛不量（照固定比重攤），要嘛先用 Python 的 `len()`（中日文字算 2）挑出每欄最長的幾個候選，只對候選呼叫 `measure`，這樣不管幾千列都是固定的十幾次。

**3. 量出來的寬度加起來一定會有塞不下的時候**

自動量只是把「理想寬度」算準，不會變出寬度。所以每一欄要有兩個數字：**理想**（裝得下現在的內容）與**下限**（還讀得出來的底線）。放得下就從理想起跳、多的照比重分；放不下就從理想往下限等比縮，縮到下限才開始切字。只有理想寬度、沒有下限的話，視窗一縮最右邊那一欄會被擠出畫面外，而那一欄不會留下任何存在過的痕跡。

---

## 十四、參考資源

| 資源                           | 說明                                                    |
| ------------------------------ | ------------------------------------------------------- |
| [TkDocs](https://tkdocs.com)   | 官方推薦教學網站，涵蓋 tkinter / ttk 完整用法，範例清楚 |
| Real Python – tkinter 系列文章 | 偏向實務案例，適合參考設計模式                          |
| ttkbootstrap 官方文件          | 提供大量現成主題範例，可直接參考排版                    |

---

## 十五、給 AI 的實作提醒

當協助使用者撰寫 / 修改 tkinter 介面程式碼時：

1. 先確認容器內排版方式（grid 或 pack），不可混用
2. 縮放需求務必設定 `columnconfigure` / `rowconfigure` 的 `weight`，並考慮設定 `minsize()`
3. 間距與字型從專案既有設定取用，不要每次隨意訂新數值
4. 元件優先使用 `ttk` 版本，除非有特殊需求需用原生 tkinter
5. UI callback 函式應盡量精簡，僅呼叫外部邏輯函式，不直接寫商業邏輯
6. 若專案已引入 `ttkbootstrap` 或 `customtkinter`，優先沿用既有套件風格，避免混搭不同美化套件
7. 遇到耗時操作（I/O、網路、大量運算），一律評估是否需要背景執行緒 + queue，避免主執行緒卡住
8. 涉及圖片顯示時，務必確認圖片物件有被保留參照，避免圖片消失的常見坑
9. 錯誤情境一律用 `messagebox` 呈現，並包在 `try/except` 中，不可讓例外直接中斷程式或用 `print` 呈現給使用者
10. 子視窗、背景執行緒、資料庫連線等資源，需在視窗關閉事件中正確釋放
