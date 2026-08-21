# Clean Word Converter · iXBRL 前置處理

將 `.docx` 入面嘅 **embedded Excel OLE 物件**原位轉換為原生 Word 表格，係 IRD iXBRL Data Preparation Tools 嘅前置處理器。

支援新式 `.xlsx` OLE 同舊式 `.xls` OLE（Excel 97‑2003 / Excel.Sheet.8），視覺上高度還原原 Excel：
- 欄寬、框線（間線）、粗體/斜體/底線、字體名/磅數、行高、對齊
- merged cells、hidden 行列、overflow 橫向合併
- **數值 0 顯示實數 `0`**（amendment 1，唔再出 Accounting 零值 dash `-`，方便 tagging）
- **描述欄配對橫向合併**（amendment 2，見 `AMENDMENTS.md`）
- **負數保證顯示負數**（amendment 3，唔會再畀貨幣符號擋返括號而變正數）
- 抽取失敗嘅物件原位留 `[TABLE EXTRACTION FAILED — manual review required]`

## 快速開始

```bash
# 1. 相依
python -m venv .venv && source .venv/bin/activate
pip install -r scripts/requirements.txt

# 2. 命令行轉換（單檔）
python scripts/clean_word_converter.py <input.docx> -o <output_clean.docx>

# 3. 自我回歸測試（自動生成 fixture → 轉換 → 斷言視覺還原）
python scripts/final_visual_test.py
```

## 檔案結構

```
clean-word-converter/
├── README.md                 # 本文件
├── AMENDMENTS.md             # 同事回饋嘅行為變更記錄
├── SKILL.md                  # 原始 skill spec（用於 AI agent port 過去）
├── scripts/
│   ├── clean_word_converter.py   # 主程式（*已包 amendment 1、2 & 3*）
│   ├── final_visual_test.py     # 自包含視覺回歸測試
│   ├── make_fixture.py          # 生成「新式 .xlsx OLE」fixture
│   ├── build_xls_fixture.py     # 生成「舊式 .xls OLE」fixture
│   └── requirements.txt
├── references/
│   └── ole-and-format-notes.md  # OLE 解碼 + number format 實現筆記
└── assets/sample-data/          # 真正嘅測試樣本 .docx
```

> `final_visual_test.py` 係自包含：會用 make_fixture / build_xls_fixture 自己生成測試檔，再 call 主程式轉換並斷言。喺新機器 clone 後可直接跑，唔使另備 fixture。

## 主程式圖示（port 時 reference）

```
.docx ZIP
 └─ word/embeddings/oleObjectN.bin
      └─ 判斷 Compound File (CFB) →
           ├─ Excel.Sheet.8 (Storage)  → 讀 xl/ 目錄（_read_xlsx）
           └─ Excel 97-2003             → 由 CFB 抽出 Workbook stream（_read_xls）
                ↓
   SheetData（rows / merges / col_widths / num_cells / cell_styles / row_heights）
                ↓
   _build_sheet_elements → native Word table（含 overflow merge + description merge）
```
