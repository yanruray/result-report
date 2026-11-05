# -*- coding: utf-8 -*-
"""
mcp_server.py
活動結報自動化產出工具 - MCP (Model Context Protocol) 伺服器

提供 Anthropic Claude Desktop、Cursor、Windsurf、Zed 等支援 MCP 的客戶端
呼叫標準化的結報產出與課程查詢工具。
"""

import os
import sys
import contextlib

try:
    from mcp.server.mcpserver import MCPServer
except ImportError:
    try:
        from mcp.server.fastmcp import FastMCP as MCPServer
    except ImportError:
        raise ImportError("請先安裝 mcp 套件：pip install mcp")

from main import (
    scan_available_courses,
    find_matching_course,
    generate_english_filename,
    read_xlsx,
    detect_event_date,
)
from report_engine import generate_report


@contextlib.contextmanager
def capture_internal_logs():
    """
    將內部 print 輸出重定向至 stderr，
    避免干擾 MCP 在 stdio 上的 JSON-RPC 通訊協定。
    """
    old_stdout = sys.stdout
    sys.stdout = sys.stderr
    try:
        yield
    finally:
        sys.stdout = old_stdout


mcp = MCPServer(
    name="event-report-generator",
    instructions="活動結報自動化產出工具，支援問卷分析、高解析圖表繪製與 Word 報告產出。"
)


@mcp.tool()
def list_available_courses() -> str:
    """
    列出目前專案中所有已偵測到的問卷課程名稱，以及對應的報名表與回饋表檔案路徑。
    """
    courses = scan_available_courses()
    if not courses:
        return "⚠️ 目前未在 data/ 或專案目錄中偵測到任何配對成功的問卷資料（需同時具備「報名」與「回饋」Excel 檔案）。"

    lines = ["### 📂 目前已偵測到的問卷課程清單：\n"]
    for idx, (title, info) in enumerate(courses.items(), 1):
        lines.append(f"{idx}. **{title}**")
        lines.append(f"   - 報名表：`{info['reg']}`")
        lines.append(f"   - 回饋表：`{info['fb']}`\n")
    return "\n".join(lines)


@mcp.tool()
def generate_event_report(
    course_name: str,
    event_date: str = "",
    lecturer: str = "活動籌備團隊",
    custom_output_filename: str = ""
) -> str:
    """
    根據課程/活動名稱關鍵字，自動配對問卷進行交叉分析，繪製統計圖表並產出 Word 結案報告。

    :param course_name: 課程或活動名稱關鍵字（例如：'工作坊'、'講座'、'迎新'）
    :param event_date: 自訂活動日期（選填，例如 '2026 年 10 月 13 日'；若未指定將自動從問卷時間戳記偵測）
    :param lecturer: 講師姓名或團隊名稱（選填，預設為 '活動籌備團隊'）
    :param custom_output_filename: 自訂輸出英文檔名（選填，例如 'Custom_Report.docx'）
    :return: 報告產出結果與核心成效摘要
    """
    courses = scan_available_courses()
    matched = find_matching_course(course_name, courses)
    if not matched:
        available = "、".join([f"「{t}」" for t in courses.keys()]) if courses else "無"
        return f"❌ 找不到與「{course_name}」相符的課程問卷。\n目前可用的課程：{available}\n請確認問卷 Excel 檔案已置於 data/ 或根目錄中。"

    title, info = matched
    reg_path = info['reg']
    fb_path = info['fb']

    fb_rows = read_xlsx(fb_path)
    if not event_date:
        event_date = detect_event_date(fb_rows) or "2026 年"

    os.makedirs('output', exist_ok=True)
    if custom_output_filename:
        if not custom_output_filename.endswith('.docx'):
            custom_output_filename += '.docx'
        out_path = os.path.join('output', custom_output_filename)
    else:
        en_filename = generate_english_filename(title)
        out_path = os.path.join('output', en_filename)

    with capture_internal_logs():
        output_result = generate_report(
            reg_input=reg_path,
            fb_input=fb_path,
            output_path=out_path,
            event_name=title,
            event_date=event_date,
            lecturer_name=lecturer,
            return_data=True
        )

    if isinstance(output_result, tuple):
        saved_path, stats = output_result
    else:
        saved_path, stats = output_result, None

    # 清除暫存圖檔
    for tmp in ['chart_c.png', 'chart_g.png', 'chart_college_temp.png', 'chart_grade_temp.png']:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except Exception:
                pass

    abs_path = os.path.abspath(saved_path)
    res_lines = [
        f"🎉 **{title} 活動成果分析報告已成功產出！**\n",
        f"- 📁 **報告路徑**：`{abs_path}`",
        f"- 📅 **活動日期**：{event_date}",
        f"- 👨‍🏫 **講師團隊**：{lecturer}",
    ]
    if stats:
        res_lines.extend([
            f"- 👥 **報名人數**：{stats.get('reg_count', 0)} 人",
            f"- 📝 **回饋填寫**：{stats.get('fb_count', 0)} 人",
            f"- 🔗 **媒合留存率**：{stats.get('retention_rate', 0):.1f}%",
            f"- ⭐ **整體滿意度**：{stats.get('overall_mean', 0):.2f} / 6.00 分 ({stats.get('overall_pos_rate', 0):.1f}% 給予 4 分以上正向評價)",
        ])
    res_lines.append("\n您可直接開啟 Word 檔案檢閱，首頁已預留簽到名冊框線，所有頁面邊界均為標準 1.0 公分。")
    return "\n".join(res_lines)


@mcp.tool()
def generate_all_reports() -> str:
    """
    一鍵為專案中所有已偵測到的問卷課程自動產出 Word 結案報告。
    """
    courses = scan_available_courses()
    if not courses:
        return "❌ 未發現任何可分析的課程問卷。請確認 Excel 檔案已放入 data/ 或根目錄。"

    results = []
    for title in courses.keys():
        res = generate_event_report(course_name=title)
        results.append(res)

    return "\n\n---\n\n".join(results)


if __name__ == "__main__":
    # 預設以 stdio transport 啟動 MCP 伺服器
    mcp.run()
