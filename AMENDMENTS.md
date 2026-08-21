# AMENDMENTS — 同事回饋行為變更

本次（2026‑08‑21）按同事對真實文件嘅回饋，改咗兩個視覺/語義行為。兩項都已經喺 `final_visual_test.py` 加咗斷言並全數通過。

## Amendment 1 — 數值 0 顯示實數 `0`

**問題**：同一個 `0` 喺原生 Word 表格度，之前依 Excel number format 嘅零值段顯示為 Accounting dash（例：`"`-`"` 或 `"-"??`）。同事指出 dash 唔適合被 tagging，要顯示返實數。

**改法**（`clean_word_converter.py` → `_fmt_number`）：當 `v == 0` 且所選零值段嘅 **字面內容全係 dash 字元**（`-`、`–`、`—`，連埋 figure space U+2007）時，直接回傳 `"0"`。其他字面 zero‑section（例如 `"N/A"`）一律保留原樣，唔會誤改。

```python
if v == 0 and all(ch in "-–—\u2007" for ch in lit):
    return "0"
```

- 影響：`#,##0;(#,##0);"-"` 嘅 0 → `0`
- 不影響：負數（`(#,##0)`）、正數、以及 `"N/A"` 呢類真字面零值

## Amendment 2 — 描述欄配對橫向合併

**問題**：真實文件中，描述有啲放喺第 1 欄、有啲放喺第 2 欄，轉出嚟唔一致。同事要求：**得 1 欄有內容就合併成 1 欄（跨兩欄）**；**兩欄都有內容就自動換行、唔合併**。

**改法**（`clean_word_converter.py`，新常數 + `_description_merges`）：

```python
DESCRIPTION_MERGE = True       # 關閉可立即還原舊行為
DESCRIPTION_MERGE_PAIR = (0, 1)  # 描述欄配對，0-based（預設 A、B 兩欄）
```

規則：
- 同一行只有左欄或只有右欄有文字 → 橫向合併 `(r, A)‑(r, B)` 跨兩欄；若內容喺右欄，會先搬去左欄再合併，確保文字保留。
- 同一行 A、B 都有內容 → **唔合併**，兩格各自自動換行（Word 固定表格內文字本身就會 wrap）。
- **唔郁**：兩欄都空、欄位係數字格、或該位置已經有 `sd.merges` / overflow merge 覆蓋（避免同 overflow 合併衝突）。

### 點樣調校欄位

喺 `clean_word_converter.py` 頂部改 `DESCRIPTION_MERGE_PAIR`（0‑based 欄 index）即可。例如改為第 2、3 欄：`DESCRIPTION_MERGE_PAIR = (1, 2)`。

## 測試

`python scripts/final_visual_test.py` 新增/更新斷言：
- `t0.cell(8, 2).text == "0"` 同 `t0.cell(8, 3).text == "0"`（零值）
- 描述欄 merge：`gridSpan == 2` 於單欄有內容嘅行、兩欄都有內容嘅行唔設 gridSpan、overflow 合併優先唔受影響、right‑only 搬左合併、數字格唔郁。
