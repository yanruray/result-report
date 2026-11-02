# -*- coding: utf-8 -*-
"""
main.py
活動結報自動化產出工具 - 主執行入口 (CLI)

使用方式：
    python main.py "課程名稱"
    python main.py --all
    python main.py --list
"""

import os
import sys
import re
import argparse
import datetime
from collections import Counter

# 設定終端機標準輸出為 UTF-8
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

from report_engine import generate_report, read_xlsx


def extract_course_title(filename):
    """從檔名自動提取乾淨的課程/活動名稱"""
    base = os.path.basename(filename)
    for ext in ['.xlsx', '.xls']:
        if base.endswith(ext):
            base = base[:-len(ext)]
    # 移除「－報名表單 (回覆)」、「－回饋表單 (回覆)」、「(回覆)」、「問卷」等後綴
    base = re.sub(r'[\s_\-－─]*(報名表單?|回饋表單?|問卷|表單)?\s*[\(\（][^\)\）]*[\)\）]', '', base)
    base = re.sub(r'[\s_\-－─]*(報名表單?|回饋表單?|問卷|表單)', '', base)
    # 移除純日期標註如 09_30 或 2026_09_30（避免與學期編號 115-1 混淆）
    base = re.sub(r'(?<!\d)(?:20\d{2}[-_/])?(?:0[1-9]|1[0-2])[-_/](?:0[1-9]|[12]\d|3[01])(?!\d)', '', base)
    base = re.sub(r'[\s_]+', ' ', base).strip()
    return base


def parse_date_from_val(val):
    """解析 Excel 儲存格日期（支援序列值或常見日期字串）"""
    if not val:
        return None
    val_str = str(val).strip()
    try:
        f = float(val_str)
        if 30000 < f < 60000:
            dt = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=f)
            return dt.date()
    except ValueError:
        pass
    m = re.search(r'(\d{4})[/-](\d{1,2})[/-](\d{1,2})', val_str)
    if m:
        return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None


def detect_event_date(fb_rows):
    """從回饋表的時間戳記自動偵測活動舉辦日期"""
    if not fb_rows or len(fb_rows) < 2:
        return None
    dates = []
    for r in fb_rows[1:]:
        if r and len(r) > 0:
            d = parse_date_from_val(r[0])
            if d:
                dates.append(d)
    if not dates:
        return None
    most_common_date, _ = Counter(dates).most_common(1)[0]
    return f"{most_common_date.year} 年 {most_common_date.month} 月 {most_common_date.day} 日"


def scan_available_courses(search_dirs=None):
    """掃描指定目錄下所有的報名與回饋 Excel 檔案並按課程分組"""
    if search_dirs is None:
        search_dirs = ['.', 'data']

    all_files = []
    for d in search_dirs:
        if os.path.exists(d):
            for root, _, files in os.walk(d):
                for f in files:
                    if f.endswith('.xlsx') and not f.startswith('~$'):
                        full_path = os.path.normpath(os.path.join(root, f))
                        all_files.append(full_path)

    # 去重
    all_files = list(dict.fromkeys(all_files))

    reg_files = [f for f in all_files if '報名' in os.path.basename(f)]
    fb_files = [f for f in all_files if '回饋' in os.path.basename(f) or ('回覆' in os.path.basename(f) and '報名' not in os.path.basename(f))]

    valid_courses = {}
    paired_reg = set()
    paired_fb = set()

    # 第 1 階段：精確標題配對
    reg_by_title = {}
    for rf in reg_files:
        t = extract_course_title(rf)
        reg_by_title[t] = rf

    fb_by_title = {}
    for ff in fb_files:
        t = extract_course_title(ff)
        fb_by_title[t] = ff

    for t in list(reg_by_title.keys()):
        if t in fb_by_title:
            rf = reg_by_title[t]
            ff = fb_by_title[t]
            valid_courses[t] = {'reg': rf, 'fb': ff}
            paired_reg.add(rf)
            paired_fb.add(ff)

    # 第 2 階段：模糊關鍵字與學期語意配對（處理檔名命名稍有出入之活動，如「迎新_社員大會」與「迎新活動」）
    unpaired_reg = [rf for rf in reg_files if rf not in paired_reg]
    unpaired_fb = [ff for ff in fb_files if ff not in paired_fb]

    for rf in unpaired_reg:
        rf_name = os.path.splitext(os.path.basename(rf))[0]
        rf_title = extract_course_title(rf)
        # 提取學期代號 (如 115-1)
        sem_m = re.search(r'(\d{3}[-_]\d)', rf_name)
        sem = sem_m.group(1).replace('_', '-') if sem_m else ''

        best_fb = None
        best_score = 0
        for ff in unpaired_fb:
            ff_name = os.path.splitext(os.path.basename(ff))[0]
            ff_title = extract_course_title(ff)
            ff_sem_m = re.search(r'(\d{3}[-_]\d)', ff_name)
            ff_sem = ff_sem_m.group(1).replace('_', '-') if ff_sem_m else ''

            if sem and ff_sem and sem != ff_sem:
                continue

            score = 0
            if sem and ff_sem and sem == ff_sem:
                score += 2

            # 關鍵詞重疊比對
            rf_clean = re.sub(r'[\s_\-－─0-9]', '', rf_title)
            ff_clean = re.sub(r'[\s_\-－─0-9]', '', ff_title)

            # 檢查是否有長度 >= 2 的共同子字串（如「迎新」）
            for i in range(len(rf_clean)):
                for j in range(i + 2, len(rf_clean) + 1):
                    sub = rf_clean[i:j]
                    if sub in ff_clean:
                        score += len(sub) * 2

            if score > best_score and score >= 4:
                best_score = score
                best_fb = ff

        if best_fb:
            ff_title = extract_course_title(best_fb)
            chosen_title = rf_title if len(rf_title) >= len(ff_title) else ff_title
            valid_courses[chosen_title] = {'reg': rf, 'fb': best_fb}
            unpaired_fb.remove(best_fb)

    return valid_courses


def find_matching_course(query, courses):
    """根據使用者輸入的關鍵字模糊比對最合適的課程"""
    if not query:
        return None

    q = re.sub(r'[\s_\-－─《》\(\)]', '', query.lower())

    matches = []
    for title, info in courses.items():
        clean_title = re.sub(r'[\s_\-－─《》\(\)]', '', title.lower())
        if q in clean_title:
            matches.append((title, info))

    if matches:
        # 取最長匹配或第一筆
        return matches[0]

    # 若關鍵字包含空格，嘗試分割關鍵字皆命中
    q_tokens = [t for t in query.lower().split() if t]
    for title, info in courses.items():
        clean_title = title.lower()
        if all(token in clean_title for token in q_tokens):
            return (title, info)

    return None


def generate_english_filename(course_title):
    """將課程名稱轉換為英文格式檔案名稱（例如 115-1_Idea_Lab_Report.docx、115-1_Orientation_Report.docx）"""
    tokens = re.findall(r'[A-Za-z0-9]+(?:-[A-Za-z0-9]+)*', course_title)
    
    # 檢查是否含有英文字母（若僅有學期數字如 115-1，則轉換中文活動名稱）
    has_alpha = any(re.search(r'[A-Za-z]', t) for t in tokens)

    if not has_alpha:
        chinese_en_map = [
            ('迎新', 'Orientation'),
            ('社員大會', 'General_Assembly'),
            ('幹部訓練', 'Officer_Training'),
            ('幹訓', 'Officer_Training'),
            ('工作坊', 'Workshop'),
            ('黑客松', 'Hackathon'),
            ('職涯', 'Career_Talk'),
            ('期初', 'Kickoff'),
            ('期末', 'Closing'),
            ('說明會', 'Info_Session'),
            ('成果展', 'Exhibition'),
        ]
        matched_en = []
        for zh, en in chinese_en_map:
            if zh in course_title and en not in matched_en:
                matched_en.append(en)

        if matched_en:
            semester_tokens = [t for t in tokens if re.search(r'\d', t)]
            if semester_tokens:
                en_base = f"{'_'.join(semester_tokens)}_{matched_en[0]}"
            else:
                en_base = matched_en[0]
        elif tokens:
            en_base = "_".join(tokens)
        else:
            en_base = "Course_Event"
    else:
        en_base = "_".join(tokens)

    return f"{en_base}_Report.docx"


def execute_report(course_title, course_info, event_date=None, lecturer="活動籌備團隊", out_path=None):
    """執行單一課程的結報產出"""
    reg_path = course_info['reg']
    fb_path = course_info['fb']

    # 讀取並自動判定日期
    fb_rows = read_xlsx(fb_path)
    if not event_date:
        detected_date = detect_event_date(fb_rows)
        event_date = detected_date if detected_date else "2026 年"

    # 確保輸出路徑為英文命名
    os.makedirs('output', exist_ok=True)
    if not out_path:
        en_filename = generate_english_filename(course_title)
        out_path = os.path.join('output', en_filename)

    print(f"\n============================================================")
    print(f"🚀 開始產生結案報告：{course_title}")
    print(f"============================================================")
    print(f"📋 報名檔案：{reg_path}")
    print(f"📋 回饋檔案：{fb_path}")
    print(f"📅 活動日期：{event_date}")
    print(f"👨‍🏫 講師團隊：{lecturer}")
    print(f"⏳ 正在分析問卷資料與繪製統計圖表...")

    output_result = generate_report(
        reg_input=reg_path,
        fb_input=fb_path,
        output_path=out_path,
        event_name=course_title,
        event_date=event_date,
        lecturer_name=lecturer,
        return_data=True
    )

    if isinstance(output_result, tuple):
        saved_path, stats = output_result
    else:
        saved_path, stats = output_result, None

    # 清除暫存圖檔
    for tmp in ['chart_college_temp.png', 'chart_grade_temp.png']:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except Exception:
                pass

    print(f"\n🎉 結案報告已成功產出！")
    print(f"📁 儲存路徑：{os.path.abspath(saved_path)}")
    if stats:
        print(f"👥 報名人數：{stats.get('reg_count', 0)} 人")
        print(f"📝 回饋填寫：{stats.get('fb_count', 0)} 人")
        print(f"🔗 媒合留存：{stats.get('retention_rate', 0):.1f}%")
        print(f"⭐ 整體滿意：{stats.get('overall_mean', 0):.2f} / 6.00 分 ({stats.get('overall_pos_rate', 0):.1f}% 給予 4 分以上正向評價)")
    print(f"============================================================\n")
    return saved_path


def main():
    parser = argparse.ArgumentParser(description="活動結報自動化產出工具")
    parser.add_argument("course", nargs="?", default=None, help="課程或活動名稱關鍵字（例如：Idea Lab、Coffee）")
    parser.add_argument("--all", action="store_true", help="自動產出所有已偵測到之課程報告")
    parser.add_argument("--list", action="store_true", help="列出目前偵測到的所有問卷課程")
    parser.add_argument("--date", default=None, help="自訂活動日期（預設自動從回饋表偵測）")
    parser.add_argument("--lecturer", default="活動籌備團隊", help="自訂講師名稱（預設：活動籌備團隊）")
    parser.add_argument("--out", default=None, help="自訂輸出 Word 檔案路徑")

    args = parser.parse_args()

    courses = scan_available_courses()

    if args.list:
        print("\n📂 目前已偵測到的課程與問卷：")
        if not courses:
            print("  （尚未發現同時具備報名表與回饋表的課程資料，請將 Excel 檔案置於 data/ 或根目錄）")
        for i, (title, info) in enumerate(courses.items(), start=1):
            print(f"  {i}. {title}")
            print(f"     報名表：{info.get('reg')}")
            print(f"     回饋表：{info.get('fb')}")
        print()
        return

    if args.all:
        if not courses:
            print("❌ 未發現任何可分析的課程問卷。請確認 Excel 檔案已放入 data/ 或根目錄。")
            return
        print(f"⚡ 準備一次產出 {len(courses)} 門課程的結報報告...")
        for title, info in courses.items():
            execute_report(title, info, event_date=args.date, lecturer=args.lecturer)
        return

    # 若指定課程名稱
    if args.course:
        matched = find_matching_course(args.course, courses)
        if not matched:
            print(f"❌ 找不到與「{args.course}」相符的課程問卷。")
            print("目前已存在的課程清單：")
            for title in courses:
                print(f"  - {title}")
            print("\n💡 提示：請確認 Excel 檔名包含該課程關鍵字，以及「報名」與「回饋」字樣。")
            sys.exit(1)
        title, info = matched
        execute_report(title, info, event_date=args.date, lecturer=args.lecturer, out_path=args.out)
        return

    # 若未指定課程名稱
    if len(courses) == 1:
        # 只有唯一一門課，自動直接執行
        title, info = list(courses.items())[0]
        print(f"💡 偵測到唯一的課程問卷：「{title}」，直接開始產出...")
        execute_report(title, info, event_date=args.date, lecturer=args.lecturer, out_path=args.out)
    elif len(courses) > 1:
        print("\n💡 偵測到多門課程問卷，請指定要完成結報的課程名稱：")
        for i, title in enumerate(courses, start=1):
            print(f"  {i}. {title}")
        print("\n使用指令範例：")
        first_title = list(courses.keys())[0]
        print(f'  python main.py "{first_title}"')
        print(f'  python main.py --all   (一次產出全部)\n')
    else:
        print("❌ 未發現任何問卷 Excel 檔案。")
        print("💡 請將報名表與回饋表 Excel 檔案放進 data/ 資料夾或專案根目錄中。")


if __name__ == '__main__':
    main()
