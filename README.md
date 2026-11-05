# 📊 活動結案報告自動化產出工具 (Event Report Generator)

一個專為活動、社課與工作坊打造的**自動化問卷分析與 Word 結報排版工具**。支援命令列 (CLI)、AI Agent 自動化指令，以及 **MCP (Model Context Protocol)** 協定。只需放入問卷並下達指令，即可一鍵完成資料清洗、指標計算、高解析度圖表繪製與標準規格 Word 排版。

---

## 🌟 核心特色

- **AI-First & MCP 深度整合**：提供標準 MCP 伺服器，無縫串接 Claude Desktop、Cursor、Windsurf 等 AI 客戶端。
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
├── main.py            # 主執行程式（CLI 入口）
├── mcp_server.py      # MCP 伺服器（供 Claude Desktop / Cursor 等工具介接）
├── report_engine.py   # 動態分析與排版引擎
├── generate_charts.py # 高解析圖表繪製模組
├── requirements.txt   # Python 套件依賴清單
├── README.md          # 專案說明文件
├── AGENTS.md          # AI Agent 執行指南
└── LICENSE            # MIT 開源授權條款
```

---

## 🚀 快速開始

### 1. 安裝環境與套件

```bash
pip install -r requirements.txt
```

### 2. 準備問卷

將問卷 Excel 檔案（`.xlsx`）放入 `data/` 資料夾：
- 報名表檔名請包含「**報名**」與**課程名稱**。
- 回饋表檔名請包含「**回饋**」或「**回覆**」與**課程名稱**。

---

## 💬 使用方式

本工具提供三種使用管道，您可依工作習慣自由選擇：

### 方式一：對 AI 下指令（推薦 ✨）

若您使用 Antigravity、Cursor、Claude Code 等具備工作區環境存取能力的 AI 助手，直接於對話框輸入：

> **「幫我完成 {課程名稱} 的結案報告」**

AI 將自動調用主程式分析並產出報告，完成後自動回報成效數據與存檔路徑。

---

### 方式二：終端機執行 (CLI)

```bash
# 產出指定課程報告
python main.py "{課程名稱}"

# 一鍵產出所有偵測到的課程報告
python main.py --all

# 查看目前偵測到的問卷清單
python main.py --list
```

產出之 Word 檔案將自動存放於 `output/` 資料夾。

---

### 方式三：透過 MCP 伺服器 (Model Context Protocol)

本專案支援標準的 MCP 協定，可作為 MCP 工具整合至各式支援 MCP 的 AI 客戶端（例如 **Claude Desktop**、**Cursor**、**Windsurf** 等）。

#### 🛠️ MCP 提供之工具一覽

| 工具名稱 | 說明 | 參數 |
| :--- | :--- | :--- |
| `list_available_courses` | 掃描並列出目前 `data/` 中所有成對的課程問卷清單 | 無 |
| `generate_event_report` | 分析指定課程資料並產出 Word 結報 | `course_name` (必填)<br>`event_date` (選填)<br>`lecturer` (選填)<br>`custom_output_filename` (選填) |
| `generate_all_reports` | 一鍵產出所有偵測到的課程結案報告 | 無 |

#### ⚙️ 客戶端設定教學

##### 1. Claude Desktop
在 Claude Desktop 設定檔（Windows: `%APPDATA%\Claude\claude_desktop_config.json`，macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`）中加入本伺服器：

```json
{
  "mcpServers": {
    "event-report-generator": {
      "command": "python",
      "args": [
        "c:/path/to/result-report/mcp_server.py"
      ],
      "env": {
        "PYTHONIOENCODING": "utf-8"
      }
    }
  }
}
```
> 💡 *請將路徑替換為您本機中 `mcp_server.py` 的絕對路徑。若使用虛擬環境，請將 `command` 指向虛擬環境的 `python` 執行檔。*

##### 2. Cursor
1. 開啟 Cursor，進入 **Settings** > **Features** > **MCP**。
2. 點選 **+ Add New MCP Server**。
3. 填寫設定：
   - **Name**: `event-report-generator`
   - **Type**: `command`
   - **Command**: `python c:/path/to/result-report/mcp_server.py`
4. 儲存後即可在 Cursor Composer / Chat 中以自然語言調用工具產出報告！

---

## 👤 開發者 (Developer)

- **Ray Lin**

---

## 📄 授權條款 (License)

本專案採用 [MIT License](LICENSE) 開源授權。
