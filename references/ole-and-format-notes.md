# Clean Word Converter — OLE 結構與 Excel 格式處理筆記

## OLE 嵌入物件結構（.docx 入面）

- `word/document.xml` 入面嘅 `<o:OLEObject r:id="rIdX">`（namespace `urn:schemas-microsoft-com:office:office`），通常包喺 `<w:object>` / `<w:r>` 入面
- `word/_rels/document.xml.rels` 將 rId 映射到 `word/embeddings/oleObjectN.bin`
- `.bin` 係 **OLE2 Compound File Binary (CFB)**，signature `D0 CF 11 E0`

### 兩代 Excel 嘅存放方式

| ProgID | 結構 | 處理 |
|---|---|---|
| `Excel.Sheet.12`（.xlsx） | CFB 入面嘅 `\x01Ole10Native` stream：4-byte little-endian 長度 + xlsx zip bytes（`PK\x03\x04`） | 解出 zip → `openpyxl` |
| `Excel.Sheet.8`（.xls） | **成個 CFB 本身就係個 .xls**（入面有 `Workbook` / `Book` stream） | 成個 bin 直接俾 `xlrd`（`file_contents=`） |
| 其他（Word.Document、Chart、Package 等） | 唔係試算表 | 留 `[TABLE EXTRACTION FAILED]` placeholder |

抽取邏輯（`_extract_ole_payload`）：先搵 `\x01Ole10Native` / `CONTENTS` stream（新式），搵唔到但有 `Workbook`/`Book` stream 就當舊式 xls 處理。

## Number format 處理（`_fmt_number`）

- Excel 格式最多四段，以 `;` 分隔：**正;負;零;文字**。每段按數值正負零揀用
- 段內規則：
  - `[Red]` / `[$-804]` 等 bracket token 剝走
  - `_x`（跟住任意一字元）= 預留 x 寬度嘅空白；`*x` = 填充 — 都剝走
  - 有 `0` / `#` → 真數字格式，套用千分位/小數位/百分比
  - 冇 `0` / `#`（得字面如 `"-"`）→ **literal 輸出**，常見於零值段 `_(* "-"??_)`
  - `?` = 對齊預留位（唔係數字內容）— 每個 `?` 補一個 **figure space U+2007**（同數字等寬），還原 Accounting dash 嘅位置
- 負數段 `(#,##0.00)` → 括弧顯示；負段缺省時補 `-` 前綴
- `General` → `Decimal(str(v))` 原值輸出，避免 float 誤差

## 視覺還原原則（好重要）

**目標係同原 Word 文件嘅 OLE 預覽一致，唔係同 Excel 開住個檔一致。**

- OLE 預覽跟**打印邏輯**：Excel 格線（gridlines）預設**唔打印** → 表格底層框線**永遠 nil**，只畫真實 cell 邊框（`xlrd` border / `openpyxl` border 逐邊映射：thin/hair→single sz=4、medium→sz=8、thick→sz=12、double/dashed/dotted 對應）
- 逐格擷取：字體名、磅數（xls `font.height` 係 twips，`/20` 得 pt）、粗體、斜體、底線、水平對齊（xls `xf.alignment.hor_align`：1=left/2=center/3=right）
- 數字/日期格預設右對齊（同 Excel）；明確對齊設定優先
- 欄寬：Excel 字元單位 → dxa（`1 char ≈ 7px + 5px padding, 1px = 15 dxa`），再按**文件實際版心**（頁寬 − margins）等比拉滿；fixed layout 防止 Word 自動拉均
- 行高：xls `rowinfo.height`（twips）、xlsx `row_dimensions.height`（pt → ×20 twips），`hRule=atLeast`
- Hidden 行/欄：略過（同 Excel 顯示一致）
- Merged cells：xlsx `ws.merged_cells`、xls `sh.merged_cells`（r2/c2 exclusive），經 hidden 行列 remap 後對應 Word `gridSpan`/`vMerge`

## Overflow 橫向合併（`_overflow_merges`）

Excel 顯示：文字格內容長過欄寬 + 右邊格**空** → 文字溢出到隔籬格顯示（唔係真 merge）。Word 表格冇溢出概念，會迫住換行。

還原方法：掃每行文字格，顯示寬度（CJK 當 2 字元）> 欄寬，而右邊**連續空格** → 喺 Word 橫向合併（gridSpan）到文字擺得落為止。規則：

- 淨係文字格（數字格唔搞 — Excel 入面數字太長出 `###`）
- 遇到非空格/已有 merge 就停
- 每行獨立計算，唔影響其他行嘅欄位結構

## 已知限制

1. OLE 物件原本嘅 floating/文字環繞定位 → 一律變 inline 表格（IRD Tagging Tool 本來就只讀 inline table）
2. Excel 顏色、conditional formatting、圖表、篩選唔會帶入
3. 公式抽 cached value（`data_only=True`），唔保留公式本身
4. OLE 物件喺頁首/頁尾/巢狀表格入面 → 唔支援（會 log warning）
5. 欄寬字元→像素換算係近似（Calibri 11 基準）；serif 字體會有少許誤差，但按版心拉滿後整體比例仍貼近
6. 極端大表截斷至 2000 行 × 100 欄（會 log warning）
