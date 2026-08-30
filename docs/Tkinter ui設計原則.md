# Python tkinter UI 設計原則參考文件

> 用途：作為 AI 助理在協助設計 / 撰寫 tkinter 介面程式碼時的參考準則。

> **使用說明**：本文件為「原則庫」，請依專案規模與需求選用，不需要每次全部套用。
>
> - 小工具／單一視窗小程式：優先參考第一~四章（版面配置、視覺體驗、進階套件、架構建議）即可
> - 涉及耗時操作、多視窗、需打包發布等情境：再參考對應章節（第五~十章）
> - **用了 ttkbootstrap，或介面覺得鈍：第十一~十三章一定要看。**那三章跟前面性質不同 ——
>   前面是通則，那三章是本專案實際踩過、量過數字的坑，而且全都是「不會報錯、只是靜靜地
>   不對或很慢」的那一類，不知道就不會想到要查
>
> 目標是寫出乾淨可維護的介面，而不是為了套用規則而增加不必要的複雜度。

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

> 「元件選用」那一條有例外，而且本專案已經踩到兩次：ttk／美化套件的元件不一定比較好。
> 歷程分頁刻意改用原生 `tk.Text` 而不是 `Treeview`（見第十三節），捲軸也刻意把繪製元件
> 換回 Tk 內建的（見第十二節）。「外觀較現代」是預設值不是鐵律，遇到效能或可控性的
> 問題時，退回原生元件是正當選項。

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

> **畫面鈍不一定是執行緒問題。** 這一節講的是「主執行緒被佔住」，但那只是卡頓的一種
> 成因。背景執行緒都做對了、畫面還是鈍的話，下一個要查的是**重畫成本** —— 見第十二節。
> 本專案 2026/08/29 實測到的卡頓（拖視窗一格要 1.6 秒）就完全不是這一節的問題，
> 執行緒的部分本來就是對的。

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

## 十二、ttkbootstrap 圖片元件的重畫成本（捲軸拉不動、進度條吃 CPU）

**症狀**：拖視窗邊框改變大小時，畫面重繪明顯延遲、整個拖曳過程都在頓。元件愈多
的分頁愈明顯，但兇手不是元件多，是**畫面上有幾條捲軸**。

**成因**：ttkbootstrap 的捲軸滑塊是一張帶透明邊的 9-slice 圖片（見它的
`style/builders/scrollbar.py`），Tk 每次重畫都要把那張圖拉伸到滑塊的實際長度、
逐像素做 alpha 合成。成本跟**滑塊的像素長度**成正比，所以「沒東西可捲」（滑塊
滿格）反而最貴 —— 而那正是多數面板平常的狀態。

2026/08/29 在這個專案實測（拉動視窗寬度，一次 resize 的中位數）：

| 情境                        | 修正前  | 修正後 |
| --------------------------- | ------- | ------ |
| 每多一條捲軸                | +250 ms | +8 ms  |
| 下單分頁（畫面上 4 條捲軸） | 1268 ms | 206 ms |
| 更新分頁（2 條）            | 537 ms  | 113 ms |
| 歷程分頁（1 條）            | 251 ms  | 65 ms  |
| 憑證分頁（0 條，對照組）    | 73 ms   | 73 ms  |

補充事實（都實測過，可以省掉重走一遍的功夫）：

- 換 ttkbootstrap 的別的主題救不了（cosmo / litera / darkly 一樣慢）；Tk 內建主題
  （vista / clam / default / alt）本來就沒這個問題，它們的滑塊是直接畫矩形
- 成本是真的 CPU 在畫，不是等視窗管理員：`update_idletasks()`（算版面）只佔 5ms，
  其餘都在 `update()` 的重繪，而且 cProfile 看得到 100% 在 Tcl 層，Python callback
  一毫秒都沒有 —— 所以「把 `<Configure>` 綁定拿掉」這類方向完全無效
- 跟 `Panedwindow`、`Canvas`、`Treeview`、元件數量都無關，純粹是捲軸條數

**做法**：用 `element_create(..., "from", "clam", ...)` 從內建主題複製一份真正
「不是圖片」的 trough / thumb 進來，再把捲軸的 layout 指過去（本專案的實作見
`ui_layout.UiLayoutMixin._use_cheap_scrollbars`，套一次全視窗的捲軸都生效）：

```python
for src in ("trough", "thumb"):
    style.element_create(f"Flat.Scrollbar.{src}", "from", "clam", src)
for axis, sticky in (("Vertical", "ns"), ("Horizontal", "we")):
    style.layout(f"{axis}.TScrollbar", [
        ("Flat.Scrollbar.trough", {"sticky": sticky, "children": [
            ("Flat.Scrollbar.thumb", {"expand": "1", "sticky": "nswe"})]})])
    style.configure(f"{axis}.TScrollbar", arrowsize=8, gripcount=0, ...)
```

**三個踩過的坑**：

- **只改 layout、把元件名字寫成 `Scrollbar.thumb` 是不夠的**。Tk 找元件會沿著樣式
  名往上找，最後還是找回 ttkbootstrap 註冊的那個圖片元件（它註冊的名字是
  `Vertical.TScrollbar.thumb`）。結果是「快了但畫壞」：圖片不再被拉伸，只剩頭尾
  兩個端帽、中間空一段。一定要 `element_create` 換成別的元件
- **這個畫壞特別難抓**，因為 widget 的行為完全正確：`identify()` 量得到的滑塊範圍
  是對的，錯的只有畫出來的像素。驗證捲軸外觀只能靠截圖，不能靠 `identify()`
- **`arrowsize` 是唯一給得動粗細的選項，而且一定要給**。換掉 layout 之後捲軸的
  粗細本來是箭頭撐出來的，沒有箭頭又沒有 `arrowsize` 就會縮成 1 像素 —— 一條看不見
  也點不到的線，不會報錯（`width`、`thickness` 這兩個名字看起來比較像，實測完全
  沒作用）。另外 `gripcount=0` 是關掉 clam 畫在滑塊正中間的那三條紋

### 同一個成因的第二處：進度條的動畫間隔

`Progressbar` 的 pbar / trough 也是 ttkbootstrap 的圖片元件，所以**動畫每跑一格
就是一次重畫**。`start()` 的間隔給太小，就等於叫它整天重畫這張圖。

本專案原本寫 `start(12)`（一秒 83 格，人眼根本分不出來），2026/08/29 實測讓程式
在「忙碌中」的時候持續吃掉 **74.9%** 的 CPU —— 而登入 20 組帳號要跑好幾分鐘，
等於整個介面在最忙的時候反而最鈍。改成 `start(100)` 之後降到 **6.8%**，看起來
一樣在動（見 `ui_background._set_busy`）。

| 做法                      | 動畫期間 CPU | 動畫期間一次 resize |
| ------------------------- | ------------ | ------------------- |
| `Progressbar`，每 12ms    | 52%          | 44 ms               |
| `Progressbar`，每 100ms   | 7%           | 29 ms               |
| 純文字轉圈四格，每 100ms  | 5%           | 13 ms               |
| 純文字轉圈四格，每 150ms  | 1%           | 12 ms               |

主題也有影響（同樣 12ms 一格：cosmo 42%、vista 17%、clam 5%），但主因是間隔，
不是主題。真的要再省，就別用 Progressbar、改成一個 `Label` 輪流換 `-` `\` `|` `/`
這四個字元 —— 但那樣字元寬度不一會左右抖動，要指定等寬字型。

**一般規則**：動畫間隔不要小於 ~80ms。人眼看不出 12ms 和 100ms 的差別，CPU 看得出來。

---

## 十三、Treeview 不好用，除非不得已別用

畫不出格線（唯一能畫的 image element 做法實測效能太差，捲動明顯變鈍）、欄寬也
不會照內容自動調（要自己用 `font.measure()` 量，還要避開在 `<Configure>` 裡重複
量）——這個專案已經盡量不用它，歷程分頁後來改用 `tk.Text` 就是因為這兩個坑。
完整細節（試過哪些做法、量出來的數字）留在 git 歷史裡，真的不得已要用時再查。

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
