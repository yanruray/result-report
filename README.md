# 📊 活動結案報告自動化產出工具 (Event Report Generator)

一個專為活動、社課與工作坊打造的**自動化問卷分析與 Word 結報排版工具**。只需放入問卷並下達指令，即可一鍵完成資料清洗、指標計算、高解析度圖表繪製與標準規格 Word 排版。

---

## 🌟 核心特色

- **AI-First 設計**：內建 `AGENTS.md`，對 AI 助手說一句話即可自動完成報告。
- **泛用自適應分析**：自動探測量表評分題、前後測能力增長及人脈交流成效，不寫死任何問卷題目。
- **跨表去重對齊**：自動依 Email 或學號計算出席留存率與現場參與人數。
- **印刷級超清圖表**：以 Pillow 4x 超取樣繪製學院與年級圓餅圖，零大型套件依賴。
- **標準 Word 公文排版**：微軟正黑體、A4 四周邊界 1.0 cm、首頁預留簽到名冊框線、次頁呈現 KPI 卡片與成效表格。

---

## 📁 專案架構

```text
result-report/
├── data/              # 存放問卷 Excel 檔案（報名表、回饋表）
├── output/            # 自動產出之 Word 結案報告 (.docx)
├── template.docx      # 標準排版規範基底範本
├── main.py            # 主執行程式（CLI 入口）
├── report_engine.py   # 動態分析與排版引擎
├── generate_charts.py # 高解析圖表繪製模組
├── requirements.txt   # Python 套件依賴清單
├── README.md          # 專案說明文件
├── AGENTS.md          # AI Agent 執行指南
└── LICENSE            # MIT 開源授權條款
```

---

## 🚀 快速開始

### 1. 安裝套件

```bash
pip install -r requirements.txt
```

### 2. 準備問卷

將問卷 Excel 檔案（`.xlsx`）放入 `data/` 資料夾：
- 報名表檔名請包含「**報名**」與**課程名稱**。
- 回饋表檔名請包含「**回饋**」或「**回覆**」與**課程名稱**。

---

## 💬 使用方式

### 方式一：對 AI 下指令（推薦 ✨）

在 AI 工具（如 Antigravity、Cursor、Claude Code 等）中直接輸入：

> **「幫我完成 {課程名稱} 的結案報告」**

AI 將自動調用主程式分析並產出報告，完成後自動回報成效數據與存檔路徑。

### 方式二：終端機執行

```bash
# 產出指定課程報告
python main.py "{課程名稱}"

# 一鍵產出所有偵測到的課程報告
python main.py --all

# 查看目前偵測到的問卷清單
python main.py --list
```

產出之 Word 檔案將存放於 `output/` 資料夾。

---

## 👤 開發者 (Developer)

- **Ray Lin**

---

## 📄 授權條款 (License)

本專案採用 [MIT License](LICENSE) 開源授權。
