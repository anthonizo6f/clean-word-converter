---
name: clean-word-converter
description: 將 .docx 入面嘅 embedded Excel OLE 物件原位轉換為原生 Word 表格（IRD iXBRL Data Preparation Tools 前置處理）。支援新式 .xlsx 同舊式 .xls OLE，視覺上高度還原原 Excel：欄寬、框線（間線）、粗體/斜體/底線、字體名/磅數、行高、對齊、merged cells、overflow 橫向合併、零值 display 0、描述欄配對合併、hidden 行列。當用戶要將內嵌 Excel 嘅 Word 文件轉做 IRD iXBRL tagging 用得嘅 clean docx 時使用。
compatibility: Created for Zo Computer
metadata:
  author: anthoni.zo.computer
---

# Clean Word Converter

將 `.docx` 入面嘅 embedded Excel OLE 物件（雙擊先開到嗰啲）原位換成**原生 Word 表格**，令文件可以餵入 IRD iXBRL Data Preparation Tools（佢讀唔到 OLE，只讀 native table）。

## 幾時用

- 用戶要交財務報表/稅務計算俾 IRD（iXBRL），而原文件係 Word + 內嵌 Excel
- 用戶講「clean word」「OLE 轉表格」「iXBRL 前置」「embedded Excel 轉 native table」

## 點用

**首選 — 控制台 UI**：審計自動化控制台（https://audit-automation-console-anthoni.zocomputer.io）→ 側欄「🧹 Clean Word」→ 上傳 .docx → 下載轉換後檔案（可批量 + zip）。

**CLI（批量/自動化）**：

```bash
python "Skills/clean-word-converter/scripts/clean_word_converter.py" \
  --input <file.docx | 資料夾> --output-dir <輸出資料夾>
```

- 輸出：每個輸入檔產生 `<原名>_clean.docx`；log 寫喺 `<output-dir>/clean_word_converter.log`
- 相依：`pip install -r "Skills/clean-word-converter/scripts/requirements.txt"`（python-docx、olefile、pandas、openpyxl、xlrd；zo.computer 已預裝）

## 轉換保證

- 段落/標題/文字順序完全保留；OLE 原位換表；輸出零 OLE 殘留（`word/embeddings` 同 `o:OLEObject` 清除）
- 視覺高度還原原 OLE 預覽：欄寬按版心拉滿、真實 cell 邊框（**唔畫格線**，OLE 預覽跟打印邏輯）、粗體/斜體/底線、字體名/磅數、行高、水平對齊、merged cells、hidden 行列略過
- 數值顯示跟 Excel number format（正;負;零 段式）：負數括弧、千分位、百分比；數值 0 一律顯示實數 `0`（唔再出 Accounting 零值 dash `-`，因為 dash 唔適合 tagging）
- 文字長過欄寬 + 右邊空格 → 自動橫向合併還原 Excel 溢出顯示
- 描述欄配對（預設頭兩欄 A、B）：同一行只有 1 欄有內容 → 橫向合併跨兩欄；第 1、第 2 欄都有內容 → 唔合併，兩格各自自動換行
- 抽唔到嘅物件 → 原位留 `[TABLE EXTRACTION FAILED — manual review required]`，唔會靜默略過

## 測試

```bash
python "Skills/clean-word-converter/scripts/final_visual_test.py"
```

自包含：生成新/舊 OLE fixture → 轉換 → 斷言全部視覺還原點。改完代碼必行。

## 深入文件

- `references/ole-and-format-notes.md` — OLE/CFB 結構、xls vs xlsx 抽取、number format 段式規則、視覺還原原則、已知限制
- `assets/sample-data/` — 兩個測試樣本（新式 .xlsx OLE、舊式 .xls OLE）

## 重要：single source of truth

控制台 `Audit Automation/app.py` 係**由呢個 skill 度 import** 主程式。改代碼只改 `Skills/clean-word-converter/scripts/clean_word_converter.py`，改完行 `final_visual_test.py`，再重啟 console service（`svc_WCp_An20uwg`）。
