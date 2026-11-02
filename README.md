# 📊 活動結案報告自動化產出工具 (Automated Event Report Generator)

這是一個專為各大專院校社團、社課、工作坊與交流活動打造的**一鍵式問卷數據分析與 Word 結案報告自動化排版工具**。

使用者下載本專案後，**只需直接對 AI 助手說一句話**：
> **「幫我完成 {課程名稱} 的結案報告」**

AI 便會自動識別資料、進行交叉分析、繪製高解析度圖表，並產出符合高規格規範的 Word 結報文件！

---

## 🌟 核心特色

1. **對 AI 極致友善（AI-First）**：
   - 內建 `AGENTS.md`，任何 AI 工具（Antigravity、Cursor、Claude Code、Copilot 等）讀取專案後皆能自動辨識並執行相應指令。
2. **泛用動態自適應引擎 (`report_engine.py`)**：
   - **完全不寫死題目**：自動探測 1–6 分或 1–5 分評分題目，自動擷取真實題目名稱與各環節評價。
   - **自動辨識活動類型**：
     - **技能學習型**：自動配對報名表與回饋表之相同能力題目，計算「前後測能力平均增長幅」與「進步學員比例」。
     - **交流體驗型**：自動統計人脈拓展人數分布、跨領域合作意願與社交成效。
   - **跨表關聯去重**：自動透過 Email 或學號對報名與回饋名單去重，精確計算到場率、媒合留存率與現場 Walk-in 人數。
3. **超高解析度統計圓餅圖 (`generate_charts.py`)**：
   - 採用 PIL 4x 超取樣技術繪製學院分布與年級結構圖，兼具視覺美感與高清晰度，且無需安裝龐大的 `matplotlib`。
4. **精準 Word 專業規格排版**：
   - 全文件統一微軟正黑體、版面邊界上下左右 1.0 cm。
   - 第一頁獨立保留**【活動簽到名冊 貼附處】**標準灰框。
   - 第二頁起呈現執行摘要（6 大 KPI 指標卡）、質性回饋整理與精進策略。

---

## 📁 專案目錄結構

```text
result-report/
├── data/                                             # 放置問卷 Excel（報名表、回饋表）
├── output/                                           # 自動產出之 Word 結案報告 (.docx，全英文命名)
├── template.docx                                     # 專業排版規範基底範本（含簽到名冊框線）
├── main.py                                           # 主執行程式（CLI 命令列入口）
├── report_engine.py                                  # 動態自適應分析與 Word 排版引擎
├── generate_charts.py                                # 高解析圖表繪製模組
├── requirements.txt                                  # Python 依賴清單
├── README.md                                         # 使用者指南
└── AGENTS.md                                         # AI Agent 執行規範指南
```

---

## 🚀 快速開始

### 1. 安裝必要套件

專案維持極簡依賴，僅需安裝 `python-docx` 與 `Pillow`：

```bash
pip install -r requirements.txt
```

### 2. 放入問卷 Excel

將 Google 表單或其他問卷系統匯出的 Excel 檔案（`.xlsx`）放入 `data/` 資料夾（或專案根目錄）：
- 報名表檔名請包含「**報名**」字樣與**課程名稱**。
- 回饋表檔名請包含「**回饋**」或「**回覆**」字樣與**課程名稱**。

---

## 💬 兩種使用方式

### 方式一：直接對 AI 下指令（最推薦 ✨）

若你在 Antigravity、Cursor、Claude Code 或其他支援終端執行的 AI 編輯器中打開本專案，只需在對話框直接輸入：

```text
幫我完成 Idea Lab 的結案報告
```
或
```text
幫我完成 Coffee Chill and Chat 的結案報告
```

AI 會自動呼叫 `main.py` 處理，並在完成後向您回報報告儲存位置與核心成效數據！

---

### 方式二：手動在終端機執行

您也可以隨時在命令列自行執行：

```bash
# 1. 依課程關鍵字產出單一報告
python main.py "Idea Lab"
python main.py "Coffee"

# 2. 自動產出目前目錄下所有偵測到的課程報告
python main.py --all

# 3. 查看目前已偵測到哪些課程問卷
python main.py --list

# 4. 自訂日期或講師資訊（若不提供，日期會自動從問卷時間戳記偵測）
python main.py "Idea Lab" --date "2026 年 10 月 13 日" --lecturer "王小明講師"
```

產出的 Word 報告將自動儲存於 `output/` 資料夾中。
