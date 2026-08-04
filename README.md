# 資料對照彙整工具 (table-lookup)

載入兩份資料表,用一份的 key 清單去篩選 / 對照另一份,組出自訂欄位的新表。
支援 **xlsx / xls / csv / json**。

## 用途

- **來源 1**:要彙整的資料(欄位較多的母檔,輸出內容多半來自這裡)
- **來源 2**:篩選條件(提供 key 清單)
- 處理邏輯:用「來源 2」某欄的值,到「來源 1」找對應列,組出新表

### 比對選項

key 以字串完全比對(自動去頭尾空白、`123.0` 視為 `123`、區分英文大小寫);
來源 1 有重複 key 時取第一筆。

唯一的勾選是**「篩選沒有對應資料」**:預設不勾 = 抓兩邊都有的;
勾了 = 反過來抓「來源 1 有、來源 2 沒有」的列。

勾「篩選沒有對應資料」時,這些列在來源 2 本來就沒有對應,輸出欄位(含自訂模板的
「插入欄位」)會自動鎖成只能挑來源 1;取消勾選時原本挑的來源 2 欄位會還原。

## 介面(四個分頁)

1. **檔案** — 選兩個來源檔(xlsx / xls 可選工作表,也支援 csv / json)
2. **比對與欄位** — 設定 key 欄位、比對選項(上表)與輸出欄位(「來源資料」直接複製欄位、「自訂」用 `{欄名}` 組合)
3. **預覽 / 執行** — 挑輸出格式與路徑,先預覽再執行
4. **日誌** — 執行記錄

設定可從「設定」選單存檔,下次自動帶入。

## 結構

```
table-lookup/
├── run.py            # 開發時直接執行
├── main_frame.py     # 供 MyTools Launcher 嵌入載入的入口
├── build.bat         # PyInstaller 打包
├── requirements.txt  # 依賴:openpyxl、xlrd(讀 .xls)
├── app/
│   ├── core.py       # 資料處理核心(讀檔 / 比對 / 輸出,stdlib + openpyxl + xlrd)
│   └── main.py       # Tkinter UI
└── tests/smoke.py
```

## 執行方式

**透過 MyTools Launcher**(建議):在 launcher 工具清單安裝後直接開啟。

**獨立執行**:
```powershell
pip install -r requirements.txt
python run.py
```

需要 Python 3.8+(內建 tkinter)。
