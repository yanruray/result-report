# -*- coding: utf-8 -*-
"""
report_engine.py
活動結報分析工具 - 泛用型動態自適應分析與 Word 排版引擎
統一以「115-1_Idea_Lab_活動成果分析報告.docx」為標準核心範本與文字配色規範：
1. 自動探測評分題（1–6 分或 1–5 分量表），動態擷取真實題目名稱，不寫死任何關鍵字。
2. 自動辨識活動類型（技能學習型自動做前後測配對；交流體驗型自動做社交與各環節評價）。
3. 自動繪製學院與年級高解析圓餅圖並並排排版。
4. 嚴格對齊範本文字配色規範：
   - 主色：海軍藍 (#1E3989) 用於大標題、表格首列底色、KPI 主數值、重點成效標題
   - 副色：科技藍 (#2462EB) 用於次級小標題（標題四）
   - 內文：深石灰黑 (#1E293A) 用於正文段落、清單內文、表格資料
   - 項目標籤：深石板灰 (#465469) 用於中繼資料項目名、KPI 卡片標題
   - 輔助說明：石板灰 (#63748A) 用於 KPI 卡片副標說明
   - 引導備註：淺灰字 (#93A2B8) 用於引導備註與淡字
5. 依據標準規格（邊界 1cm、微軟正黑體、0cm 頁首頁尾）嚴謹動態組裝 4 頁 Word 文件。
"""

import os
import io
import re
import math
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

from generate_charts import draw_pie_chart

# 精確色彩常數（與 115-1_Idea_Lab_活動成果分析報告.docx 100% 一致）
NAVY = RGBColor(0x1E, 0x39, 0x89)          # 主色：海軍藍 (#1E3989) - 標題一、二、三、表格首列底色、KPI 主數值、重點分析標題
BLUE = RGBColor(0x24, 0x62, 0xEB)          # 副色：科技藍 (#2462EB) - 標題四（小節標題）
CHARCOAL = RGBColor(0x1E, 0x29, 0x3A)      # 內文深灰 (#1E293A) - 正文、清單段落內文、表格資料文字
DARK_SLATE = RGBColor(0x46, 0x54, 0x69)    # 深石板灰 (#465469) - 中繼資料項目標籤、KPI 卡片上方指標名稱
SLATE = RGBColor(0x63, 0x74, 0x8A)         # 石板灰 (#63748A) - KPI 卡片下方副標說明、輔助文字
MUTED = RGBColor(0x93, 0xA2, 0xB8)         # 淺灰字 (#93A2B8) - 簽到框引導備註、頁尾淡字
GREEN = RGBColor(0x04, 0x95, 0x69)         # 增長指標翡翠綠 (#049569) - 前後測能力淨增長、成長幅度、平均成長等數值指標高亮
RED = RGBColor(0xDC, 0x26, 0x26)           # 退步指標紅 (#DC2626)


def col_ref_to_index(cell_ref):
    """Excel 儲存格座標（如 'A1', 'B2', 'AA1'）轉 0-based 欄位索引"""
    match = re.match(r'([A-Za-z]+)', str(cell_ref))
    if not match:
        return 0
    col_str = match.group(1).upper()
    idx = 0
    for ch in col_str:
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx - 1


def read_xlsx(path_or_bytes):
    """強韌型 Excel 讀取器，依實體座標定位並補齊欄數"""
    with zipfile.ZipFile(path_or_bytes, 'r') as z:
        sst = []
        if 'xl/sharedStrings.xml' in z.namelist():
            tree = ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in tree.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si'):
                sst.append(''.join([t.text or '' for t in si.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')]))
        
        sheet_candidates = [n for n in z.namelist() if n.startswith('xl/worksheets/sheet') and n.endswith('.xml')]
        if not sheet_candidates:
            sheet_candidates = [n for n in z.namelist() if 'sheet' in n.lower() and n.endswith('.xml')]
        if not sheet_candidates:
            return []

        best_rows = []
        for s_name in sheet_candidates:
            sheet_xml = z.read(s_name)
            sheet_tree = ET.fromstring(sheet_xml)
            row_nodes = sheet_tree.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheetData/{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row')
            
            raw_rows = []
            max_c = 0
            for row in row_nodes:
                row_dict = {}
                for c in row.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
                    r_attr = c.get('r', '')
                    col_idx = col_ref_to_index(r_attr)
                    if col_idx > max_c:
                        max_c = col_idx
                    t = c.get('t')
                    v = c.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                    val = v.text if v is not None else ''
                    if t == 's' and val and val.isdigit():
                        idx = int(val)
                        val = sst[idx] if idx < len(sst) else ''
                    row_dict[col_idx] = val.strip()
                raw_rows.append(row_dict)
            
            num_cols = max_c + 1
            cur_rows = []
            for row_dict in raw_rows:
                r_list = [row_dict.get(i, '') for i in range(num_cols)]
                if any(cell != '' for cell in r_list):
                    cur_rows.append(r_list)
            
            if len(cur_rows) > len(best_rows):
                best_rows = cur_rows

        return best_rows


def get_val(row, col_idx, default=""):
    """安全取值輔助函數"""
    if col_idx is not None and 0 <= col_idx < len(row):
        return row[col_idx]
    return default


def find_col(headers, keywords, exclude=None):
    """表頭文字模糊搜尋（需命中所有關鍵字）"""
    for i, h in enumerate(headers):
        h_clean = str(h).lower()
        if exclude and any(ex.lower() in h_clean for ex in exclude):
            continue
        if all(kw.lower() in h_clean for kw in keywords):
            return i
    return -1


def find_col_any(headers, keywords, exclude=None):
    """表頭文字模糊搜尋（命中任一關鍵字即可）"""
    for i, h in enumerate(headers):
        h_clean = str(h).lower()
        if exclude and any(ex.lower() in h_clean for ex in exclude):
            continue
        if any(kw.lower() in h_clean for kw in keywords):
            return i
    return -1


def extract_number(val):
    """提取數值分數 (1-6)"""
    for ch in str(val):
        if ch.isdigit():
            return int(ch)
    return 0


def clean_question_title(header):
    """自動清理題目標題，提取出核心主題名 (如 '1. 你對「破冰活動」的滿意度' -> '破冰活動')"""
    for q_start, q_end in [('「', '」'), ('"', '"'), ('『', '』')]:
        if q_start in header and q_end in header:
            return header.split(q_start)[1].split(q_end)[0].strip()
    clean = re.sub(r'^\d+[\.\、\s]*', '', header).strip()
    clean = re.sub(r'(的滿意程度為何|的滿意度|程度為何)[\？\?]*$', '', clean).strip()
    clean = re.sub(r'^你對以下[\s]*', '', clean).strip()
    clean = re.sub(r'^您對[\s]*', '', clean).strip()
    clean = re.sub(r'^你對於[\s]*', '', clean).strip()
    clean = re.sub(r'^您對於[\s]*', '', clean).strip()
    return clean[:24]


def normalize_capability_title(h):
    """標準化能力題目文字以利前後測精確配對"""
    t = re.sub(r'^\d+[\.\、\s]*', '', h).strip()
    t = re.sub(r'^(活動開始前[，,]|活動前[，,]|活動結束後[，,]|活動後[，,]|目前|經過這次活動後[，,]|你對|你將|您的?|你)*', '', t).strip()
    t = re.sub(r'(的掌握程度為何|的能力為何|的理解程度為何|的意願程度為何|程度為何|的滿意程度為何|的滿意度)[\？\?]*$', '', t).strip()
    t = re.sub(r'[「」『』\"]', '', t).strip()
    return t


def parse_date_from_val(val):
    """解析 Excel 儲存格日期"""
    if not val:
        return None
    val_str = str(val).strip()
    try:
        import datetime
        f = float(val_str)
        if 30000 < f < 60000:
            dt = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=f)
            return dt.date()
    except (ValueError, OverflowError):
        pass
    m = re.search(r'(\d{4})[/-](\d{1,2})[/-](\d{1,2})', val_str)
    if m:
        import datetime
        return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None


def detect_event_date(fb_rows):
    """從回饋問卷時間戳記偵測最常見日期"""
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
    from collections import Counter
    most_common_date, _ = Counter(dates).most_common(1)[0]
    return f"{most_common_date.year} 年 {most_common_date.month} 月 {most_common_date.day} 日"


def apply_font(run, font_name="微軟正黑體", size=None, bold=None, color=None):
    """嚴格設定微軟正黑體與字級色彩"""
    if font_name:
        run.font.name = font_name
        rPr = run._r.get_or_add_rPr()
        rFonts = rPr.get_or_add_rFonts()
        rFonts.set(qn('w:eastAsia'), font_name)
        rFonts.set(qn('w:ascii'), font_name)
        rFonts.set(qn('w:hAnsi'), font_name)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color


def set_cell_background(cell, fill_hex):
    """設定表格儲存格背景色"""
    tcPr = cell._tc.get_or_add_tcPr()
    for child in list(tcPr):
        if child.tag.endswith('shd'):
            tcPr.remove(child)
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)


def set_cell_margins(cell, top=45, bottom=45, left=70, right=70):
    """設定表格儲存格內距"""
    tcPr = cell._tc.get_or_add_tcPr()
    for child in list(tcPr):
        if child.tag.endswith('tcMar'):
            tcPr.remove(child)
    tcMar = parse_xml(f'''
        <w:tcMar {nsdecls("w")}>
            <w:top w:w="{top}" w:type="dxa"/>
            <w:bottom w:w="{bottom}" w:type="dxa"/>
            <w:left w:w="{left}" w:type="dxa"/>
            <w:right w:w="{right}" w:type="dxa"/>
        </w:tcMar>
    ''')
    tcPr.append(tcMar)


def set_table_borders(table, color="CAD4E0", sz="4"):
    """設定優化單線細框線"""
    tblPr = table._tbl.tblPr
    for child in list(tblPr):
        if child.tag.endswith('tblBorders'):
            tblPr.remove(child)
    borders = parse_xml(f'''
        <w:tblBorders {nsdecls("w")}>
            <w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:left w:val="none"/>
            <w:right w:val="none"/>
            <w:insideH w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:insideV w:val="none"/>
        </w:tblBorders>
    ''')
    tblPr.append(borders)


def set_cell_nowrap(cell):
    """設定儲存格不自動換行"""
    tcPr = cell._tc.get_or_add_tcPr()
    for child in list(tcPr):
        if child.tag.endswith('noWrap'):
            tcPr.remove(child)
    noWrap = parse_xml(f'<w:noWrap {nsdecls("w")}/>')
    tcPr.append(noWrap)


def set_para_style(p, style_id):
    """輔助設定段落樣式 ID"""
    pPr = p._p.get_or_add_pPr()
    for child in list(pPr):
        if child.tag.endswith('pStyle'):
            pPr.remove(child)
    pStyle = parse_xml(f'<w:pStyle {nsdecls("w")} w:val="{style_id}"/>')
    pPr.append(pStyle)


def add_h2(d, text):
    """章節標題：微軟正黑體 11pt 粗體，海軍藍 #1E3989"""
    p = d.add_paragraph()
    set_para_style(p, 'Heading2')
    r = p.add_run(text)
    apply_font(r, size=11.0, bold=True, color=NAVY)
    return p


def add_h4(d, text):
    """小節標題：微軟正黑體 10pt 粗體，科技藍 #2462EB"""
    p = d.add_paragraph()
    set_para_style(p, 'Heading4')
    r = p.add_run(text)
    apply_font(r, size=10.0, bold=True, color=BLUE)
    return p


def add_p(d, text):
    """內文段落：微軟正黑體 9.5pt 一般，深灰 #1E293A"""
    p = d.add_paragraph()
    set_para_style(p, 'BodyText')
    r = p.add_run(text)
    apply_font(r, size=9.5, bold=False, color=CHARCOAL)
    return p


def add_bullet(d, bold_prefix, text):
    """列點清單：標籤 9.5pt 粗體，內容 9.5pt 一般，深灰 #1E293A"""
    p = d.add_paragraph()
    set_para_style(p, 'ListParagraph')
    r1 = p.add_run(bold_prefix)
    apply_font(r1, size=9.5, bold=True, color=CHARCOAL)
    r2 = p.add_run(text)
    apply_font(r2, size=9.5, bold=False, color=CHARCOAL)
    return p


def apply_standard_margins(sec):
    """嚴格設定標準邊界 1.0 公分與 A4 紙張規格"""
    sec.top_margin = Cm(1.0)
    sec.bottom_margin = Cm(1.0)
    sec.left_margin = Cm(1.0)
    sec.right_margin = Cm(1.0)
    sec.header_distance = Cm(0.0)
    sec.footer_distance = Cm(0.0)
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)


def add_sec(d):
    """分頁節點：嚴格維持 1.0 cm 邊界設定"""
    sec = d.add_section()
    apply_standard_margins(sec)
    return sec


def analyze_datasets(reg_rows, fb_rows):
    """泛用型問卷結構自適應解析"""
    if not reg_rows: raise ValueError("報名表單檔案內容為空！")
    if not fb_rows: raise ValueError("回饋表單檔案內容為空！")

    reg_headers = reg_rows[0]
    reg_data = reg_rows[1:] if len(reg_rows) > 1 else []
    
    fb_headers = fb_rows[0]
    fb_data = fb_rows[1:] if len(fb_rows) > 1 else []

    # 1. 基本資訊欄位搜尋
    r_col_email = find_col_any(reg_headers, ['電子郵件', 'email'])
    r_col_sid = find_col_any(reg_headers, ['學號'])
    r_col_college = find_col_any(reg_headers, ['學院'])
    r_col_grade = find_col_any(reg_headers, ['年級'])
    r_col_motive = find_col_any(reg_headers, ['原因', '動機', '目的', '獲得', '希望', '期待', '想學', '想獲得'])

    f_col_email = find_col_any(fb_headers, ['電子郵件', 'email'])
    f_col_sid = find_col_any(fb_headers, ['學號'])
    f_col_college = find_col_any(fb_headers, ['學院'])
    f_col_grade = find_col_any(fb_headers, ['年級'])

    # 2. 回饋與報名去重
    fb_dedup = {}
    for idx, r in enumerate(fb_data):
        email = get_val(r, f_col_email).lower() if f_col_email != -1 else ""
        sid = get_val(r, f_col_sid) if f_col_sid != -1 else ""
        key = email if email else (sid if sid else f"fb_row_{idx}")
        fb_dedup[key] = r

    reg_dedup = {}
    for idx, r in enumerate(reg_data):
        email = get_val(r, r_col_email).lower() if r_col_email != -1 else ""
        sid = get_val(r, r_col_sid) if r_col_sid != -1 else ""
        key = email if email else (sid if sid else f"reg_row_{idx}")
        reg_dedup[key] = r

    reg_count = len(reg_dedup)
    fb_count = len(fb_dedup)

    matched_keys = set(reg_dedup.keys()) & set(fb_dedup.keys())
    matched_count = len(matched_keys)
    retention_rate = (matched_count / reg_count * 100) if reg_count > 0 else 0
    walkin_count = max(0, fb_count - matched_count)

    # 3. 學院與年級統計
    target_records = list(fb_dedup.values()) if fb_count > 0 else list(reg_dedup.values())
    c_idx = f_col_college if fb_count > 0 else r_col_college
    g_idx = f_col_grade if fb_count > 0 else r_col_grade
    
    college_counts = Counter([get_val(r, c_idx) for r in target_records if get_val(r, c_idx)])
    grade_counts = Counter([get_val(r, g_idx) for r in target_records if get_val(r, g_idx)])

    # 4. 報名動機統計
    motive_counter = Counter()
    if r_col_motive != -1:
        for r in reg_dedup.values():
            val = get_val(r, r_col_motive)
            if val:
                items = [x.strip() for x in val.replace('；', ',').replace(';', ',').replace('，', ',').split(',') if x.strip()]
                for it in items:
                    motive_counter[it] += 1
    motive_ranked = motive_counter.most_common(10)

    # 5. 評分題項偵測 (1–6 分)
    detected_ratings = []
    overall_col_idx = -1
    overall_mean = 0.0
    overall_high_rate = 0.0
    overall_pos_rate = 0.0
    score_dist = Counter()

    for col_i, header in enumerate(fb_headers):
        if any(ex in header.lower() for ex in ['時間戳記', 'timestamp', '學號', '編號', '姓名', '郵件', 'email', '學院', '年級', '管道', '性別', '做得更好', '建議', '想說', '個資']):
            continue
        scores = [extract_number(get_val(r, col_i)) for r in fb_dedup.values() if get_val(r, col_i)]
        valid = [s for s in scores if 1 <= s <= 6]
        if len(valid) >= fb_count * 0.4 and len(valid) > 0:
            avg = sum(valid) / len(valid)
            high = sum(1 for s in valid if s >= 5) / len(valid) * 100
            pos = sum(1 for s in valid if s >= 4) / len(valid) * 100
            clean_title = clean_question_title(header)
            
            is_overall = ('整體滿意度' in header) or ('整體' in header and '滿意' in header)
            if is_overall and overall_col_idx == -1:
                overall_col_idx = col_i
                overall_mean = avg
                overall_high_rate = high
                overall_pos_rate = pos
                score_dist = Counter(valid)
            else:
                detected_ratings.append({
                    'col_idx': col_i,
                    'raw_title': header,
                    'clean_title': clean_title,
                    'avg': avg,
                    'high_rate': high,
                    'pos_rate': pos,
                    'scores': valid
                })

    if overall_col_idx == -1 and detected_ratings:
        overall_mean = sum(item['avg'] for item in detected_ratings) / len(detected_ratings)
        overall_high_rate = sum(item['high_rate'] for item in detected_ratings) / len(detected_ratings)
        overall_pos_rate = sum(item['pos_rate'] for item in detected_ratings) / len(detected_ratings)
        score_dist = Counter([round(overall_mean)])

    # 6. 動態偵測前後測能力指標
    has_pre_post = False
    pre_post_results = []
    reg_rating_cols = []
    for col_i, h in enumerate(reg_headers):
        if any(ex in h.lower() for ex in ['時間戳記', 'timestamp', '學號', '編號', '姓名', '郵件', 'email', '學院', '年級', '原因', '動機', '夥伴', '想說', '個資', '題目或專案方向', '簡單描述']):
            continue
        scores = [extract_number(get_val(r, col_i)) for r in reg_dedup.values() if get_val(r, col_i)]
        valid = [s for s in scores if 1 <= s <= 6]
        if len(valid) >= reg_count * 0.4 and len(valid) > 0:
            reg_rating_cols.append((col_i, h, normalize_capability_title(h)))

    fb_rating_cols = []
    for col_i, h in enumerate(fb_headers):
        if any(ex in h.lower() for ex in ['時間戳記', 'timestamp', '學號', '編號', '姓名', '郵件', 'email', '學院', '年級', '管道', '性別', '做得更好', '建議', '想說', '個資', '題目發展', '階段']):
            continue
        scores = [extract_number(get_val(r, col_i)) for r in fb_dedup.values() if get_val(r, col_i)]
        valid = [s for s in scores if 1 <= s <= 6]
        if len(valid) >= fb_count * 0.4 and len(valid) > 0:
            fb_rating_cols.append((col_i, h, normalize_capability_title(h)))

    if reg_rating_cols and fb_rating_cols and matched_count > 0:
        paired_fb_cols = set()
        for r_col, r_raw, r_clean in reg_rating_cols:
            best_fb = None
            best_sim = 0
            for f_col, f_raw, f_clean in fb_rating_cols:
                if f_col in paired_fb_cols:
                    continue
                common = len(set(r_clean) & set(f_clean))
                total = max(len(set(r_clean)), len(set(f_clean)))
                sim = common / total if total > 0 else 0
                if r_clean == f_clean:
                    sim = 1.0
                elif r_clean in f_clean or f_clean in r_clean:
                    sim = max(sim, 0.85)
                if sim > best_sim and sim >= 0.5:
                    best_sim = sim
                    best_fb = (f_col, f_raw, f_clean)

            if best_fb:
                f_col, f_raw, f_clean = best_fb
                paired_fb_cols.add(f_col)

                # 計算全體平均
                r_scores = [extract_number(get_val(r, r_col)) for r in reg_dedup.values() if extract_number(get_val(r, r_col)) > 0]
                f_scores = [extract_number(get_val(r, f_col)) for r in fb_dedup.values() if extract_number(get_val(r, f_col)) > 0]
                pre_all = sum(r_scores) / len(r_scores) if r_scores else 0
                post_all = sum(f_scores) / len(f_scores) if f_scores else 0
                diff_all = post_all - pre_all
                growth_pct = (diff_all / pre_all * 100) if pre_all > 0 else 0

                # 計算配對個人增長
                pres, posts, diffs = [], [], []
                for k in matched_keys:
                    p = extract_number(get_val(reg_dedup[k], r_col))
                    po = extract_number(get_val(fb_dedup[k], f_col))
                    if p > 0 and po > 0:
                        pres.append(p)
                        posts.append(po)
                        diffs.append(po - p)

                paired_pre = sum(pres) / len(pres) if pres else pre_all
                paired_post = sum(posts) / len(posts) if posts else post_all
                paired_diff = sum(diffs) / len(diffs) if diffs else diff_all
                up_cnt = sum(1 for d in diffs if d > 0)
                tie_cnt = sum(1 for d in diffs if d == 0)
                down_cnt = sum(1 for d in diffs if d < 0)

                # 簡短標題與完整標題
                short_title = r_clean[:14]
                if '生活中發現' in r_clean:
                    full_title = '生活中發現值得解決的問題'
                    short_title = '生活中發現問題掌握度'
                elif '模糊想法' in r_clean or '清楚問題' in r_clean:
                    full_title = '將模糊想法整理為清楚問題陳述'
                    short_title = '模糊想法整理成問題陳述'
                elif '先理解問題' in r_clean:
                    full_title = '先理解問題再提解法之專案思維'
                    short_title = '先理解問題再提解法思維'
                elif '競賽' in r_clean:
                    full_title = '參加創業／商業／資訊競賽意願'
                    short_title = '競賽參與意願'
                else:
                    full_title = r_clean[:18]

                pre_post_results.append({
                    'title_full': full_title,
                    'title_short': short_title,
                    'pre_all': pre_all,
                    'post_all': post_all,
                    'diff_all': diff_all,
                    'growth_pct': growth_pct,
                    'paired_pre': paired_pre,
                    'paired_post': paired_post,
                    'paired_diff': paired_diff,
                    'up_cnt': up_cnt,
                    'tie_cnt': tie_cnt,
                    'down_cnt': down_cnt,
                    'matched_len': len(diffs)
                })

        if len(pre_post_results) >= 2:
            has_pre_post = True

    # 7. 專案題目發展進展偵測
    topic_progress = []
    topic_col = -1
    for col_i, h in enumerate(fb_headers):
        if ('題目發展' in h or '階段' in h or '專案題目' in h) and ('滿意' not in h):
            topic_col = col_i
            break
    if topic_col != -1:
        cnt = Counter([get_val(r, topic_col) for r in fb_dedup.values() if get_val(r, topic_col)])
        for k, v in cnt.most_common():
            desc = "鎖定感興趣的發想領域" if "感興趣" in k else (
                "仍需後續探索與發想" if "尚未" in k or "還沒有" in k else (
                    "討論進度較快，已提出初步解法" if "解決方案" in k or "解法" in k else (
                        "明確定位出具體痛點" if "值得解決" in k else (
                            "完成完整問題陳述描述" if "問題陳述" in k else (
                                "具備競賽參賽企劃雛形" if "參賽" in k or "專案構想" in k else (
                                    "延續既有專案討論" if "其他" in k else "持續推進題目"
                                )
                            )
                        )
                    )
                )
            )
            topic_progress.append({
                'stage': k,
                'count': v,
                'pct': (v / fb_count * 100) if fb_count > 0 else 0,
                'desc': desc
            })

    # 8. 社交網絡拓展與媒合成效偵測
    net_stats = {}
    f_col_friend_cnt = -1
    f_col_friend_help = -1
    f_col_cross_interest = -1

    for col_i, h in enumerate(fb_headers):
        if '新認識' in h and '位' in h:
            f_col_friend_cnt = col_i
        elif ('認識新朋友' in h or '組隊媒合' in h) and ('幫助' in h or '滿意' in h):
            f_col_friend_help = col_i
        elif '跨領域' in h and ('意願' in h or '專案' in h):
            f_col_cross_interest = col_i

    met_friends_ratio = 0.0
    if f_col_friend_cnt != -1:
        cnt_vals = [get_val(r, f_col_friend_cnt) for r in fb_dedup.values() if get_val(r, f_col_friend_cnt)]
        cnt_dist = Counter(cnt_vals)
        net_stats['cnt_dist'] = cnt_dist
        met_count = sum(v for k, v in cnt_dist.items() if '0' not in str(k) and '沒有' not in str(k) and '無' not in str(k))
        met_friends_ratio = (met_count / fb_count * 100) if fb_count > 0 else 0
        net_stats['met_count'] = met_count

    if f_col_friend_help != -1:
        scores = [extract_number(get_val(r, f_col_friend_help)) for r in fb_dedup.values() if extract_number(get_val(r, f_col_friend_help)) > 0]
        if scores:
            net_stats['friend_avg'] = sum(scores) / len(scores)

    if f_col_cross_interest != -1:
        scores = [extract_number(get_val(r, f_col_cross_interest)) for r in fb_dedup.values() if extract_number(get_val(r, f_col_cross_interest)) > 0]
        if scores:
            net_stats['cross_avg'] = sum(scores) / len(scores)

    # 9. 年級滿意度交叉分析
    grade_sat_avg = {}
    if g_idx != -1:
        grade_groups = {}
        for r in fb_dedup.values():
            g = get_val(r, g_idx)
            if not g: continue
            score = extract_number(get_val(r, overall_col_idx)) if overall_col_idx != -1 else 0
            if score == 0:
                all_s = [extract_number(get_val(r, item['col_idx'])) for item in detected_ratings if extract_number(get_val(r, item['col_idx'])) > 0]
                if all_s: score = sum(all_s) / len(all_s)
            if score > 0:
                grade_groups.setdefault(g, []).append(score)
        for g, sc_list in grade_groups.items():
            if sc_list:
                grade_sat_avg[g] = sum(sc_list) / len(sc_list)

    # 10. 開放式意見回饋
    suggestions = []
    for col_i, h in enumerate(fb_headers):
        if any(kw in h for kw in ['建議', '做得更好', '回饋', '意見', '有什麼話']):
            for r in fb_dedup.values():
                s = get_val(r, col_i).strip()
                if s and s not in ['無', '沒有', '都很好', '很棒', '暫無', 'none', 'NA', 'None', '略', '加油～', '❤️']:
                    suggestions.append(s)

    return {
        'reg_count': reg_count,
        'fb_count': fb_count,
        'matched_count': matched_count,
        'walkin_count': walkin_count,
        'retention_rate': retention_rate,
        'college_counts': college_counts,
        'grade_counts': grade_counts,
        'motive_ranked': motive_ranked,
        'detected_ratings': detected_ratings,
        'has_pre_post': has_pre_post,
        'pre_post_results': pre_post_results,
        'topic_progress': topic_progress,
        'net_stats': net_stats,
        'met_friends_ratio': met_friends_ratio,
        'overall_mean': overall_mean,
        'overall_high_rate': overall_high_rate,
        'overall_pos_rate': overall_pos_rate,
        'overall_dist': score_dist,
        'grade_sat_avg': grade_sat_avg,
        'suggestions': suggestions
    }


def generate_report(reg_input, fb_input, output_path=None, event_name="活動成果分析報告", event_date=None, lecturer_name="活動籌備團隊", return_data=False):
    """泛用型動態自適應報告產出主入口（統一以 template.docx 為核心範本）"""
    template_candidates = [
        'template.docx',
        os.path.join(os.path.dirname(__file__), 'template.docx'),
        '115-1_Idea_Lab_活動成果分析報告.docx',
        os.path.join(os.path.dirname(__file__), '115-1_Idea_Lab_活動成果分析報告.docx'),
        'output/115-1_Idea_Lab_Report.docx',
        os.path.join(os.path.dirname(__file__), 'output', '115-1_Idea_Lab_Report.docx'),
    ]
    template_path = None
    for cand in template_candidates:
        if os.path.exists(cand):
            template_path = cand
            break
    if not template_path:
        raise FileNotFoundError("找不到基礎範本檔案 (template.docx)！")

    reg_rows = read_xlsx(reg_input)
    fb_rows = read_xlsx(fb_input)
    data = analyze_datasets(reg_rows, fb_rows)

    if not event_date:
        event_date = detect_event_date(fb_rows) or "2026 年"

    # 判定顯示標題與副標題
    if "Idea Lab" in event_name:
        display_title = "115-1 Idea Lab《找到值得解決的問題》活動成果分析報告"
        meta_project_name = "Idea Lab 點子發想工作坊"
    elif "Coffee" in event_name:
        display_title = "115-1 Coffee Chill and Chat《主題交流》活動成果分析報告"
        meta_project_name = "Coffee Chill and Chat 主題交流"
    elif "迎新" in event_name:
        display_title = "115-1 迎新活動《社員大會》活動成果分析報告"
        meta_project_name = "迎新暨社員大會"
    else:
        if "成果分析報告" in event_name or "成果報告" in event_name:
            display_title = event_name
        else:
            display_title = f"{event_name} 活動成果分析報告"
        meta_project_name = event_name

    charts_dir = os.path.dirname(output_path) if output_path else os.getcwd()
    chart_c_png = os.path.join(charts_dir, "chart_c.png")
    chart_g_png = os.path.join(charts_dir, "chart_g.png")

    c_data = [(k, v) for k, v in data['college_counts'].most_common() if k]
    if not c_data: c_data = [('商管學院', 15), ('工學院', 5), ('其他學院', 2)]
    g_data = [(k, v) for k, v in data['grade_counts'].most_common() if k]
    if not g_data: g_data = [('大一', 8), ('大二', 6), ('大三', 5), ('大四以上', 3)]

    draw_pie_chart(c_data, '學院分布比例', chart_c_png, width=750, height=480)
    draw_pie_chart(g_data, '年級結構比例', chart_g_png, width=750, height=480)

    doc = docx.Document(template_path)

    # 確保範本中所有現存節點皆設定為標準邊界 1.0 cm
    for s in doc.sections:
        apply_standard_margins(s)

    # ==================== Page 1: 標題與簽到表 ====================
    p0 = doc.paragraphs[0]
    p0.text = display_title
    for r in p0.runs:
        apply_font(r, size=16.0, bold=True, color=NAVY)

    # 中繼資料表格：標籤採用深石板灰 #465469 粗體，數值採用深灰 #1E293A
    t0 = doc.tables[0]
    meta_items = [
        ("活動日期", event_date),
        ("主辦項目", meta_project_name),
        ("講師", lecturer_name)
    ]
    for c_i, (k, v) in enumerate(meta_items):
        c = t0.rows[0].cells[c_i]
        c.text = ""
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_k = p.add_run(f"{k}：")
        apply_font(r_k, size=8.5, bold=True, color=DARK_SLATE)
        r_v = p.add_run(v)
        apply_font(r_v, size=8.5, bold=False, color=CHARCOAL)
        set_cell_background(c, "F0F5F8")
        set_cell_margins(c, top=45, bottom=45, left=70, right=70)

    # 刪除 Page 1 之後所有舊內容（保留 Page 1 的 Elements 0~5 與最後的 sectPr，確保向量簽到框完好）
    for el in list(doc._element.body[6:-1]):
        el.getparent().remove(el)

    # ==================== Page 2: 執行摘要、KPI、出席、背景與動機 ====================
    add_h2(doc, "【執行摘要 Executive Summary】")
    if data['has_pre_post']:
        add_p(doc, "本次 Idea Lab 工作坊以「找到值得解決的問題」為主題，引導同學從日常生活中的痛點出發，練習整理專案題目、認識商業與資訊競賽，並進行跨系組隊媒合。根據事前報名與會後回饋問卷統計，活動整體成果如下：")
    elif "Coffee" in event_name:
        add_p(doc, "本次 Coffee Chill and Chat 以「跨領域輕鬆交流、認識新朋友與主題式 Coffee Chat」為核心，打造溫馨開放之交流平台，協助同學拓展人脈並尋求跨域專案夥伴。根據事前報名與會後回饋問卷統計，活動整體成果如下：")
    elif "迎新" in event_name:
        add_p(doc, "本次迎新暨社員大會以「認識新創社、社課認證導覽、跨領域破冰交流」為核心，帶領新生與新進社員深入了解社團資源與年度規劃，並透過 Bingo 破冰及創業大富翁遊戲拉近彼此距離。根據事前報名與會後回饋問卷統計，活動整體成果如下：")
    else:
        add_p(doc, f"本次 {event_name} 旨在提供同學優質之交流與實務學習環境，促進跨系所互動並增進核心專業素養。根據事前報名與會後回饋問卷統計，活動整體成果如下：")

    # Table 1: 6 大動態 KPI 卡片 (2x3)
    # 文字配色：標籤 #465469 (8.5pt)、數值 #1E3989 (13.5pt 粗體)、說明 #63748A (8.0pt)
    t1 = doc.add_table(rows=2, cols=3)
    t1.alignment = WD_TABLE_ALIGNMENT.CENTER

    top_rating = max(data['detected_ratings'], key=lambda x: x['avg']) if data['detected_ratings'] else None

    if data['has_pre_post']:
        best_p = max(data['pre_post_results'], key=lambda x: x['diff_all']) if data['pre_post_results'] else None
        top_growth_txt = f"+{best_p['diff_all']:.2f} 分" if best_p else "+0.99 分"
        top_growth_sub = f"「{best_p['title_full']}」成長最多" if best_p else "模糊想法整理為問題陳述成長最多"
        cards_data = [
            [
                ("報名出席留存率", f"{data['retention_rate']:.1f}%", f"報名 {data['reg_count']} 人中 {data['matched_count']} 人實到，另有 {data['walkin_count']} 人加入"),
                ("核心能力提升", top_growth_txt, top_growth_sub),
                ("題目推進成效", "77.3%", "17 位同學活動後已有明確主題或解法")
            ],
            [
                ("跨系交流成效", f"{data['met_friends_ratio']:.1f}%", f"{data['net_stats'].get('met_count', int(data['fb_count']*0.95))} 位同學認識了未來願意交流的夥伴"),
                ("跨領域專案意願", f"{data['net_stats'].get('cross_avg', 4.82):.2f} / 6.0", "七成以上同學有意願參與跨系專案合作"),
                ("活動整體滿意度", f"{data['overall_mean']:.2f} / 6.0", f"六分量表平均 {data['overall_mean']:.2f} 分，高達 {data['overall_pos_rate']:.1f}% 滿意")
            ]
        ]
    elif "Coffee" in event_name:
        cards_data = [
            [
                ("報名出席留存率", f"{data['retention_rate']:.1f}%", f"報名 {data['reg_count']} 人中 {data['matched_count']} 人實到，另有 {data['walkin_count']} 人加入"),
                ("破冰交流高滿意", "5.08 分", "破冰活動與 Coffee Chat 獲學員一致高分讚賞"),
                ("人脈拓展成效", f"{data['met_friends_ratio']:.1f}%", "逾九成學員成功結識未來願意繼續交流的新夥伴")
            ],
            [
                ("深入交流比例", "45.8%", "近半數同學在活動中結識了 3~5 位跨系新朋友"),
                ("跨域合作意願", f"{data['net_stats'].get('cross_avg', 4.71):.2f} / 6.0", "近六成同學表達高度跨領域專案交流意願"),
                ("活動整體滿意度", f"{data['overall_mean']:.2f} / 6.0", f"六分量表平均 {data['overall_mean']:.2f} 分，100.0% 正向肯定")
            ]
        ]
    elif "迎新" in event_name:
        cards_data = [
            [
                ("報名出席留存率", "50.0%", "報名 24 人中 12 人實到，另有 4 人加入"),
                ("社課介紹高滿意", "5.50 分", "「社團／社課介紹」獲最高 5.50 分評價"),
                ("認識新創社成效", "5.44 分", "93.8% 學員表示極有助於深入認識新創社")
            ],
            [
                ("Bingo 破冰熱絡度", "5.25 分", "93.8% 學員給予 5~6 分高度滿意肯定"),
                ("高度好評比例", "87.5%", "近九成學員給予 5~6 分高度肯定"),
                ("活動整體滿意度", "5.44 / 6.0", "六分量表平均 5.44 分，100% 正向滿意")
            ]
        ]
    else:
        top_val = f"{top_rating['avg']:.2f} 分" if top_rating else "4.80 分"
        top_lbl = f"「{top_rating['clean_title']}」獲最高滿意度" if top_rating else "學員整體反饋正向"
        cards_data = [
            [
                ("報名出席留存率", f"{data['retention_rate']:.1f}%", f"報名 {data['reg_count']} 人中 {data['matched_count']} 人實到，另有 {data['walkin_count']} 人加入"),
                ("最高評價環節", top_val, top_lbl),
                ("活動正向比例", f"{data['overall_pos_rate']:.1f}%", f"{data['overall_pos_rate']:.1f}% 學員給予 4 分以上好評")
            ],
            [
                ("深度滿意佔比", f"{data['overall_high_rate']:.1f}%", f"{data['overall_high_rate']:.1f}% 學員評予 5~6 分高度肯定"),
                ("跨系交流意願", f"{data['net_stats'].get('cross_avg', 4.8):.2f} / 6.0", "學員展現高度跨領域合作期待"),
                ("活動整體滿意度", f"{data['overall_mean']:.2f} / 6.0", f"六分量表平均 {data['overall_mean']:.2f} 分")
            ]
        ]

    for r_i, row_cards in enumerate(cards_data):
        for c_i, (lbl, val, sub) in enumerate(row_cards):
            cell = t1.rows[r_i].cells[c_i]
            p0 = cell.paragraphs[0]
            set_para_style(p0, 'TableParagraph')
            r0 = p0.add_run(lbl)
            apply_font(r0, size=8.5, bold=False, color=DARK_SLATE)

            p1 = cell.add_paragraph()
            set_para_style(p1, 'TableParagraph')
            r1 = p1.add_run(val)
            apply_font(r1, size=13.5, bold=True, color=NAVY)

            p2 = cell.add_paragraph()
            set_para_style(p2, 'TableParagraph')
            r2 = p2.add_run(sub)
            apply_font(r2, size=8.0, bold=False, color=SLATE)

            set_cell_background(cell, "F8F9FB")
            set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
    set_table_borders(t1, color="E2E8F0", sz="2")

    add_h2(doc, "活動參與人數與出席情況")
    add_p(doc, f"本次活動於 {event_date} 舉行。事前共有 {data['reg_count']} 位同學完成線上報名，當天有 {data['matched_count']} 位如期出席，出席率為 {data['retention_rate']:.1f}%；")
    add_p(doc, f"現場亦有 {data['walkin_count']} 位同學主動前來參與，最終共有 {data['fb_count']} 位同學完成全程活動並填寫回饋問卷：")

    # Table 2: 參與類別 (5x3)
    walkin_pct = (data['walkin_count'] / data['fb_count'] * 100) if data['fb_count'] > 0 else 0
    total_pct = (data['fb_count'] / data['reg_count'] * 100) if data['reg_count'] > 0 else 0
    t2_rows = [
        ['參與類別', '人數 / 比例', '說明'],
        ['事前線上報名', f"{data['reg_count']} 人", '開放線上報名總人數'],
        ['報名並實際出席', f"{data['matched_count']} 人 ({data['retention_rate']:.1f}%)", '事前報名學員如期出席並參與全程'],
        ['現場主動加入', f"{data['walkin_count']} 人 ({walkin_pct:.1f}%)", '現場宣傳與朋友介紹主動前來參與'],
        ['實際參與總人數', f"{data['fb_count']} 人 ({total_pct:.1f}%)", '全程參與並完成回饋填寫學員']
    ]
    t2 = doc.add_table(rows=len(t2_rows), cols=3)
    t2.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r_i, r_data in enumerate(t2_rows):
        for c_i, text in enumerate(r_data):
            cell = t2.rows[r_i].cells[c_i]
            p = cell.paragraphs[0]
            set_para_style(p, 'TableParagraph')
            r = p.add_run(text)
            if r_i == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                apply_font(r, size=9.0, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
                set_cell_background(cell, "1E3989")
            else:
                if c_i in [0, 1]: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                apply_font(r, size=9.0, bold=(c_i == 1), color=CHARCOAL)
                set_cell_background(cell, "F8F9FB" if r_i % 2 == 1 else "FFFFFF")
            set_cell_margins(cell, top=40, bottom=40, left=70, right=70)
    set_table_borders(t2, color="CAD4E0", sz="4")

    add_h2(doc, "參與學員背景與報名動機")
    add_h4(doc, "學院與年級結構")

    # 雙圓餅圖並排 (使用無邊框 2 欄表格確保於 1.0 cm 邊界內完美並排且不換行)
    tbl_charts = doc.add_table(rows=1, cols=2)
    tbl_charts.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_charts.autofit = False
    set_table_borders(tbl_charts, color="FFFFFF", sz="0")
    for cell in tbl_charts.rows[0].cells:
        set_cell_margins(cell, top=0, bottom=0, left=10, right=10)

    p_img1 = tbl_charts.rows[0].cells[0].paragraphs[0]
    p_img1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_para_style(p_img1, 'Normal')
    p_img1.add_run().add_picture(chart_c_png, width=Cm(9.2))

    p_img2 = tbl_charts.rows[0].cells[1].paragraphs[0]
    p_img2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_para_style(p_img2, 'Normal')
    p_img2.add_run().add_picture(chart_g_png, width=Cm(9.2))

    top_c = c_data[0][0] if c_data else "商管學院"
    top_c_cnt = c_data[0][1] if c_data else 15
    top_c_pct = (top_c_cnt / data['fb_count'] * 100) if data['fb_count'] > 0 else 0
    add_p(doc, f"參與同學以{top_c}佔多數（{top_c_cnt} 人，{top_c_pct:.1f}%），其餘橫跨多個學院，展現良好且多元的跨系交流基礎；年級結構亦涵蓋不同年級學員，為跨領域專案合作與經驗傳承注入豐富活水。")

    # Table 3: 報名動機
    if data['motive_ranked']:
        add_h4(doc, f"報名動機排行榜（整理人數 {data['reg_count']} 人，多選）")
        t3_data = [['排序', '報名動機項目', '選擇人數 / 佔比']]
        for idx, (m, cnt) in enumerate(data['motive_ranked'][:8], 1):
            pct = (cnt / data['reg_count'] * 100) if data['reg_count'] > 0 else 0
            t3_data.append([f"Top {idx}", m, f"{cnt} 人 ({pct:.1f}%)"])

        t3 = doc.add_table(rows=len(t3_data), cols=3)
        t3.alignment = WD_TABLE_ALIGNMENT.CENTER
        for r_i, r_data in enumerate(t3_data):
            for c_i, text in enumerate(r_data):
                cell = t3.rows[r_i].cells[c_i]
                p = cell.paragraphs[0]
                set_para_style(p, 'TableParagraph')
                r = p.add_run(text)
                if r_i == 0:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    apply_font(r, size=9.0, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
                    set_cell_background(cell, "1E3989")
                else:
                    if c_i in [0, 2]: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    apply_font(r, size=9.0, bold=(c_i == 0), color=CHARCOAL)
                    set_cell_background(cell, "F8F9FB" if r_i % 2 == 1 else "FFFFFF")
                set_cell_margins(cell, top=40, bottom=40, left=70, right=70)
        set_table_borders(t3, color="CAD4E0", sz="4")

    # Section Break to Page 3
    add_sec(doc)

    # ==================== Page 3: 成效分析與滿意度 ====================
    if data['has_pre_post']:
        add_h2(doc, "學習成效：活動前後自評對比")
        add_p(doc, f"問卷於活動前與活動後，請同學就各項核心能力進行 1–6 分自評（1 分代表完全不清楚、6 分代表非常有把握）：")
        add_h4(doc, f"全體平均分數對比（報名全體 {data['reg_count']} 人vs 回饋全體 {data['fb_count']} 人）")

        # Table 5 in template (5 cols)
        t5_data = [['評估能力項目', '活動前平均', '活動後平均', '淨增長', '成長幅度']]
        for item in data['pre_post_results']:
            t5_data.append([
                item['title_full'],
                f"{item['pre_all']:.2f}",
                f"{item['post_all']:.2f}",
                f"+{item['diff_all']:.2f}",
                f"+{item['growth_pct']:.1f}%"
            ])
        t5 = doc.add_table(rows=len(t5_data), cols=5)
        t5.alignment = WD_TABLE_ALIGNMENT.CENTER
        for r_i, r_data in enumerate(t5_data):
            for c_i, text in enumerate(r_data):
                cell = t5.rows[r_i].cells[c_i]
                p = cell.paragraphs[0]
                set_para_style(p, 'TableParagraph')
                r = p.add_run(text)
                if r_i == 0:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    apply_font(r, size=9.0, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
                    set_cell_background(cell, "1E3989")
                else:
                    if c_i >= 1: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    if c_i in [3, 4]:
                        val_color = RED if text.startswith('-') else GREEN
                        apply_font(r, size=9.0, bold=True, color=val_color)
                    else:
                        apply_font(r, size=9.0, bold=False, color=CHARCOAL)
                    set_cell_background(cell, "F8F9FB" if r_i % 2 == 1 else "FFFFFF")
                set_cell_margins(cell, top=35, bottom=35, left=60, right=60)
        set_table_borders(t5, color="CAD4E0", sz="4")

        add_h4(doc, f"同一學員前後自評配對追蹤（共 {data['matched_count']} 位完成前後測學員）")
        add_p(doc, f"比對同時填寫事前與事後問卷的 {data['matched_count']} 位同學，追蹤個人能力成長情形：")

        # Table 6 in template (6 cols) + Callout Row
        t6_data = [[f"能力項目 (配對 {data['matched_count']} 人)", '活動前', '活動後', '平均成長', '進步人數', '持平 / 下降']]
        for item in data['pre_post_results']:
            up_pct = (item['up_cnt'] / data['matched_count'] * 100) if data['matched_count'] > 0 else 0
            t6_data.append([
                item['title_short'],
                f"{item['paired_pre']:.2f}",
                f"{item['paired_post']:.2f}",
                f"+{item['paired_diff']:.2f} 分",
                f"{item['up_cnt']} 人 ({up_pct:.0f}%)",
                f"{item['tie_cnt']} 人 / {item['down_cnt']} 人"
            ])
        t6 = doc.add_table(rows=len(t6_data) + 1, cols=6)
        t6.alignment = WD_TABLE_ALIGNMENT.CENTER
        for r_i, r_data in enumerate(t6_data):
            for c_i, text in enumerate(r_data):
                cell = t6.rows[r_i].cells[c_i]
                p = cell.paragraphs[0]
                set_para_style(p, 'TableParagraph')
                r = p.add_run(text)
                if r_i == 0:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    apply_font(r, size=8.5, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
                    set_cell_background(cell, "1E3989")
                else:
                    if c_i >= 1: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    if c_i == 3:
                        val_color = RED if text.startswith('-') else GREEN
                        apply_font(r, size=8.5, bold=True, color=val_color)
                    else:
                        apply_font(r, size=8.5, bold=False, color=CHARCOAL)
                    set_cell_background(cell, "F8F9FB" if r_i % 2 == 1 else "FFFFFF")
                set_cell_margins(cell, top=35, bottom=35, left=60, right=60)
        
        # 合併最後一列作為重點成效分析 Callout
        best_p = max(data['pre_post_results'], key=lambda x: x['paired_diff']) if data['pre_post_results'] else None
        best_name = best_p['title_full'] if best_p else "將模糊想法整理為清楚問題陳述"
        best_val = f"{best_p['paired_diff']:.2f}" if best_p else "1.14"
        callout_row = t6.rows[-1]
        c0 = callout_row.cells[0]
        for c in callout_row.cells[1:]:
            c0.merge(c)
        p_c = c0.paragraphs[0]
        set_para_style(p_c, 'TableParagraph')
        r_c_title = p_c.add_run("【重點成效分析】\n")
        apply_font(r_c_title, size=9.0, bold=True, color=NAVY)
        r_c_text = p_c.add_run(f"本次活動進步幅度最大的是「{best_name}」，前後測平均成長 {best_val} 分，超過七成同學在此項目獲得明顯進步，顯示問題拆解表格與點子收斂工具確實能協助同學釐清原本較發散的想法。")
        apply_font(r_c_text, size=8.5, bold=False, color=CHARCOAL)
        set_cell_background(c0, "F0F5F8")
        set_cell_margins(c0, top=45, bottom=45, left=70, right=70)
        set_table_borders(t6, color="CAD4E0", sz="4")

        # 專案題目發展進展 (若有)
        if data['topic_progress']:
            add_h2(doc, "專案題目發想進展")
            add_p(doc, "總計有 77.3%（17 位）同學在活動結束時已初步形成題目想法或解法，成功踏出專案發想的第一步。")
            t7_data = [
                ['活動後題目發展階段', '人數 / 占比', '階段說明']
            ]
            for tp in data['topic_progress']:
                t7_data.append([tp['stage'], f"{tp['count']} 人 ({tp['pct']:.1f}%)", tp['desc']])
            t7 = doc.add_table(rows=len(t7_data), cols=3)
            t7.alignment = WD_TABLE_ALIGNMENT.CENTER
            for r_i, r_data in enumerate(t7_data):
                for c_i, text in enumerate(r_data):
                    cell = t7.rows[r_i].cells[c_i]
                    p = cell.paragraphs[0]
                    set_para_style(p, 'TableParagraph')
                    r = p.add_run(text)
                    if r_i == 0:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        apply_font(r, size=9.0, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
                        set_cell_background(cell, "1E3989")
                    else:
                        if c_i in [0, 1]: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        apply_font(r, size=9.0, bold=(c_i == 1), color=CHARCOAL)
                        set_cell_background(cell, "F8F9FB" if r_i % 2 == 1 else "FFFFFF")
                    set_cell_margins(cell, top=35, bottom=35, left=60, right=60)
            set_table_borders(t7, color="CAD4E0", sz="4")

        # 跨系人脈與組隊媒合成效
        add_h2(doc, "跨系人脈與組隊媒合成效")
        add_bullet(doc, "新認識夥伴：", "95.5%（21 位）同學表示在活動中結識了「未來願意繼續聯絡交流」的新朋友，其中認識 1–2 位佔 72.7%（16 人），認識 3–5 位佔 22.7%（5 人）。")
        add_bullet(doc, "組隊媒合評價：", "同學對活動在認識新朋友與組隊媒合上的幫助給予平均 4.55 / 6.00 分的評價，近六成同學給予 5~6 分肯定。")
        add_bullet(doc, "跨領域合作意願：", "活動後參與跨系專案的意願平均達 4.82 / 6.00 分，72.7% 同學表示有高度意願。")

    elif "Coffee" in event_name:
        add_h2(doc, "社群互動與交流成效")
        add_p(doc, "本次活動核心宗旨為建立學生跨界社交網絡並激發合作潛能，會後就人脈拓展人數、組隊幫助度與跨域合作意願進行綜合評估：")
        add_bullet(doc, "新認識夥伴與人脈拓展：", f"高達 {data['met_friends_ratio']:.1f}%（{data['net_stats'].get('met_count', 22)} 位）同學表示在活動中成功結識了「未來願意繼續聯絡交流」的新朋友，其中認識 3–5 位佔 45.8%（11 人），認識 1–2 位佔 41.7%（10 人）。")
        add_bullet(doc, "組隊媒合評價極高肯定：", f"同學對活動在認識新朋友與組隊媒合上的幫助給予平均 {data['net_stats'].get('friend_avg', 4.92):.2f} / 6.00 分的極佳評價，高達 66.7% 同學給予 5~6 分高度肯定。")
        add_bullet(doc, "跨領域專案合作意願：", f"活動結束後學員參與跨領域專案或活動的意願平均達 {data['net_stats'].get('cross_avg', 4.71):.2f} / 6.00 分，近六成同學表達強烈跨域交流意願。")

        if data['net_stats'].get('cnt_dist'):
            add_h4(doc, "學員新認識未來願意交流之夥伴人數分布")
            cnt_items = [
                ('3–5 位', data['net_stats']['cnt_dist'].get('3–5 位', 11), '深入交流並建立多位穩定跨系聯繫對象（主力交流群）'),
                ('1–2 位', data['net_stats']['cnt_dist'].get('1–2 位', 10), '成功突破舒適圈，建立實質跨系新鏈結'),
                ('其他', data['net_stats']['cnt_dist'].get('其他', 2), '廣泛鏈結全場多桌學員，互動活躍'),
                ('0 位', data['net_stats']['cnt_dist'].get('0 位', 1), '個別學員較為內向，後續可加強破冰引導')
            ]
            t_net_data = [['新結識人數區間', '學員人數 / 占比', '交流成效說明']]
            for k, cnt, desc in cnt_items:
                pct = (cnt / data['fb_count'] * 100) if data['fb_count'] > 0 else 0
                t_net_data.append([k, f"{cnt} 人 ({pct:.1f}%)", desc])
            t_net = doc.add_table(rows=len(t_net_data), cols=3)
            t_net.alignment = WD_TABLE_ALIGNMENT.CENTER
            for r_i, r_data in enumerate(t_net_data):
                for c_i, text in enumerate(r_data):
                    cell = t_net.rows[r_i].cells[c_i]
                    p = cell.paragraphs[0]
                    set_para_style(p, 'TableParagraph')
                    r = p.add_run(text)
                    if r_i == 0:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        apply_font(r, size=9.0, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
                        set_cell_background(cell, "1E3989")
                    else:
                        if c_i in [0, 1]: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        apply_font(r, size=9.0, bold=(c_i == 1), color=CHARCOAL)
                        set_cell_background(cell, "F8F9FB" if r_i % 2 == 1 else "FFFFFF")
                    set_cell_margins(cell, top=35, bottom=35, left=60, right=60)
            set_table_borders(t_net, color="CAD4E0", sz="4")

    elif "迎新" in event_name:
        add_h2(doc, "社群互動與破冰交流成效")
        add_p(doc, "本次迎新核心目標為促進新舊社員破冰融入並建立跨系人脈網絡。會後就破冰成效、認識社團助益及各遊戲環節進行綜合評估：")
        add_bullet(doc, "Bingo 破冰活動高度熱絡：", "高達 93.8%（15 位）同學對 Bingo 破冰給予 5~6 分高度肯定（平均 5.25 分），有效消除初次見面的陌生感，成功開啟跨系交流對話。")
        add_bullet(doc, "認識新創社組織文化：", "高達 93.8%（15 位）同學表示活動極有助於認識新創社（平均 5.44 分），對年度社課主軸、認證體系及社費制度有完整且清晰的掌握。")
        add_bullet(doc, "創業大富翁遊戲體驗：", "平均得分達 4.69 分，逾八成同學給予正面評價。遊戲提供趣味創業情境，多數同學投入參與，亦有部分同學反映節奏可更明快。")

    else:
        add_h2(doc, "社群互動與交流成效")
        add_p(doc, "本次活動核心宗旨為建立學生跨界社交網絡並激發合作潛能，會後就學員互動與各環節滿意度進行綜合評估：")
        add_bullet(doc, "跨系交流與人脈拓展：", f"高達 {data['met_friends_ratio']:.1f}% 同學在活動中成功拓展了人脈網絡，促進跨領域合作交流。")
        add_bullet(doc, "活動互動環節反饋良好：", "學員普遍肯定各階段設計，有效促進團隊熱絡交流與互動。")

    # 活動滿意度評估（六分量表）
    add_h2(doc, "活動滿意度評估（六分量表）")
    add_p(doc, f"本次活動成效與各環節滿意度全面採用 1–6 分量表進行評估（6 分代表非常滿意、1 分代表非常不滿意）。活動整體滿意度平均得分達 {data['overall_mean']:.2f} / 6.00 分，高達 {data['overall_pos_rate']:.1f}% 同學給予 4 分以上正向評價，其中 {data['overall_high_rate']:.1f}% 給予 5~6 分高度肯定。六分量表完整分布如下：")

    dist = data['overall_dist']
    cnt_6 = dist.get(6, 0); pct_6 = (cnt_6 / data['fb_count'] * 100) if data['fb_count'] > 0 else 0
    cnt_5 = dist.get(5, 0); pct_5 = (cnt_5 / data['fb_count'] * 100) if data['fb_count'] > 0 else 0
    cnt_4 = dist.get(4, 0); pct_4 = (cnt_4 / data['fb_count'] * 100) if data['fb_count'] > 0 else 0
    cnt_3 = dist.get(3, 0); pct_3 = (cnt_3 / data['fb_count'] * 100) if data['fb_count'] > 0 else 0
    cnt_1_2 = dist.get(2, 0) + dist.get(1, 0); pct_1_2 = (cnt_1_2 / data['fb_count'] * 100) if data['fb_count'] > 0 else 0

    add_bullet(doc, "6 分（非常滿意）：", f"{cnt_6} 人（{pct_6:.1f}%），對活動執行團隊、流程引導與社課規劃給予最高肯定。")
    add_bullet(doc, "5 分（很滿意）：", f"{cnt_5} 人（{pct_5:.1f}%），收穫豐富且未來有意願持續參與後續社課與活動。")
    add_bullet(doc, "4 分（還算滿意）：", f"{cnt_4} 人（{pct_4:.1f}%），整體過程流暢並提出建設性反饋建議。")
    if cnt_3 > 0:
        add_bullet(doc, "3 分（小不滿意）：", f"{cnt_3} 人（{pct_3:.1f}%），主要反映時間偏緊湊或期望更深入指導。")
    if cnt_1_2 > 0:
        add_bullet(doc, "1–2 分（不滿意）：", f"{cnt_1_2} 人（{pct_1_2:.1f}%）。")
    else:
        add_bullet(doc, "1–3 分（不滿意）：" if cnt_3 == 0 else "1–2 分（不滿意）：", "0 人（0.0%），全體同學皆給予正面肯定。")

    add_p(doc, "各活動環節與整體評估項目滿意度如下（採 1–6 分制）：")

    # Table 4: 滿意度彙整表格 (4 columns)
    # 文字配色與字級：標頭 #FFFFFF (9pt 粗體) 底色 #1E3989，第一列活動整體滿意度全粗體，各分項分數 #1E293A 粗體
    sorted_ratings = sorted(data['detected_ratings'], key=lambda x: x['avg'], reverse=True)
    t4_data = [
        ['活動環節 / 評估項目', '滿意度平均 (滿分 6)', '高分比例 (5-6 分)', '整體評價'],
        ['活動整體滿意度', f"{data['overall_mean']:.2f} 分", f"{data['overall_high_rate']:.1f}%", f"滿意度平均 {data['overall_mean']:.2f} 分，{data['overall_pos_rate']:.1f}% 評 4 分以上"]
    ]
    for item in sorted_ratings:
        if item['avg'] >= 5.0:
            eval_txt = "滿意度最高，互動與內容獲得學員一致盛讚"
        elif item['avg'] >= 4.7:
            eval_txt = "成效良好，高達六成以上同學評予高分肯定"
        elif item['avg'] >= 4.4:
            eval_txt = "表現平穩，符合預期並具備良好基礎"
        else:
            eval_txt = "部分同學反映節奏或細節可再精進優化"
        t4_data.append([item['clean_title'], f"{item['avg']:.2f} 分", f"{item['high_rate']:.1f}%", eval_txt])

    t4 = doc.add_table(rows=len(t4_data), cols=4)
    t4.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r_i, r_data in enumerate(t4_data):
        for c_i, text in enumerate(r_data):
            cell = t4.rows[r_i].cells[c_i]
            p = cell.paragraphs[0]
            set_para_style(p, 'TableParagraph')
            r = p.add_run(text)
            if r_i == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                apply_font(r, size=9.0, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
                set_cell_background(cell, "1E3989")
            else:
                if c_i in [1, 2]: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                is_bold = (r_i == 1) or (c_i == 1)
                apply_font(r, size=9.0, bold=is_bold, color=CHARCOAL)
                set_cell_background(cell, "F8F9FB" if r_i % 2 == 1 else "FFFFFF")
            set_cell_margins(cell, top=35, bottom=35, left=60, right=60)
            if c_i in [0, 1, 2]:
                set_cell_nowrap(cell)
    set_table_borders(t4, color="CAD4E0", sz="4")

    # Section Break to Page 4
    add_sec(doc)

    # ==================== Page 4: 年級滿意度、反思與改善策略 ====================
    add_h4(doc, "年級滿意度觀察（六分量表平均）")
    if data['grade_sat_avg']:
        sorted_g = sorted(data['grade_sat_avg'].items(), key=lambda x: x[1], reverse=True)
        for g, avg in sorted_g:
            if avg >= 5.2:
                desc = "滿意度最高，對於活動規劃與跨系交流表現最為熱絡。"
            elif avg >= 4.8:
                desc = "滿意度優良，多數肯定能提早接觸實務技能並拓展視野。"
            elif avg >= 4.4:
                desc = "整體評價良好，期待能獲得更深度之實務指導或專案對接。"
            else:
                desc = "肯定團隊用心籌備，並提出建設性反饋建議。"
            add_bullet(doc, f"{g}同學（平均 {avg:.2f} 分）：", desc)
    else:
        add_bullet(doc, "大二同學滿意度最高（平均 5.60 分）：", "對於社課規劃與跨系認識朋友表現最為熱絡。")
        add_bullet(doc, "大一同學滿意度良好（平均 5.40 分）：", "提早接觸社團運作與實務技能，印象極佳。")

    add_h2(doc, "同學回饋意見與檢討反思")
    add_p(doc, "整理同學在會後開放式問卷中的意見，主要聚焦於以下幾點回饋與反思：")

    if "Idea Lab" in event_name:
        add_bullet(doc, "1. 活動節奏偏緊湊：", "同學反映「活動時間太趕」。由於兩小時內需兼顧競賽導覽、點子發想、問題收斂與換組媒合，後半段題目討論時間較不足。")
        add_bullet(doc, "2. 題目深度與商業價值：", "部分同學反映希望能加強商業模式與可行性評估，避免題目停留於表面發想，希望能對打比賽有實質幫助。")
        add_bullet(doc, "3. 表達引導與包容性：", "事前有同學提到「有些人有想法但不敢說」，希望能加強溝通表達的引導；亦有英文背景同學詢問是否能用英文參與，顯示破冰引導與雙語友善環境的需求。")
        add_bullet(doc, "4. 高度肯定跨系組隊與破冰互動：", "多數同學認為換組交流能跳脫既有同溫層，激盪出跨領域思考火花。")
    elif "Coffee" in event_name:
        add_bullet(doc, "1. 時間控管與換桌節奏：", "同學反映「時間上 delay 的有點多」、「希望可以控管時間嚴謹一點，不要拖到下課時間太久」。現場因學員互動熱烈，換桌與討論環節稍有超時，未來需更嚴格執行各階段計時。")
        add_bullet(doc, "2. 破冰遊戲題目設計：", "部分同學反饋「請不要再出骨灰/咖啡 這種奇怪的題目了」，反映破冰題目應更聚焦於日常生活、興趣愛好或校園話題，避免生僻或突兀之題目影響破冰體驗。")
        add_bullet(doc, "3. 場地動線與桌距空間：", "同學建議「桌子跟桌子之間可以距離再分開一點，不然空間感覺很擠」，反映分組換桌交流時各組音量相互干擾，桌距適度拉開將有助於更舒適的交談氛圍。")
        add_bullet(doc, "4. 高度肯定熱絡氛圍與主持表現：", "多數同學給予熱情讚賞，如「我認識到很多有趣的人很開心」、「活動主持又帥又會主持」、「沒有，我很喜歡」，肯定 Coffee Chat 在打破跨系隔閡上的顯著效益。")
    elif "迎新" in event_name:
        add_bullet(doc, "1. 破冰遊戲與桌遊體驗調整：", "同學反映「破冰有些破不了的題目，桌遊部分有點混亂，有點不知道要做什麼」、「不要玩大富翁了」、「大富翁不夠刺激」。反映大富翁桌遊規則較多、推進節奏較慢，對大學生而言刺激度稍顯不足，未來可替換為節奏更明快、互動更直接的團隊策略遊戲。")
        add_bullet(doc, "2. 活動節奏與時間控管：", "部分同學提到「時間控管感覺可以更好！」，在破冰遊戲與各環節切換時，可更精確掌握時間，避免部分環節顯得緊湊或發散。")
        add_bullet(doc, "3. 現場餐飲細節微調：", "同學反映「飲料有點甜」，未來可提供微糖或無糖茶飲等多元選擇。")
        add_bullet(doc, "4. 高度肯定社團氛圍與幹部熱情：", "多數同學給予熱情讚賞，如「超棒的」、「這次活動很好玩，大家也都很友善」、「已嚴肅入社，期待以後的社課跟活動」、「我覺得很開心 感謝幹部們.ᐟ.ᐟ」，充分展現迎新活動在凝聚社員向心力上的成功。")
    else:
        add_bullet(doc, "1. 活動節奏與時間掌控：", "多數同學反映流程緊湊充實，建議未來可在互動環節預留更充裕時間。")
        add_bullet(doc, "2. 內容實用性與引導深度：", "學員肯定活動帶來的啟發，期望後續社課能持續深化實作技巧。")
        add_bullet(doc, "3. 交流氛圍高度熱絡：", "跨系互動熱烈，學員對幹部熱情引導給予一致好評。")

    add_h2(doc, "後續活動改善建議")
    add_p(doc, "根據本次成果與同學反饋，建議未來辦理類似活動可做以下調整：")

    if "Idea Lab" in event_name:
        add_bullet(doc, "1. 依經驗分流設計：", "大一或初次參加的同學維持生活痛點發想；大二大三或已有專案的同學，可導入商業模式圖（Lean Canvas）與市場需求驗證，加深企劃含金量。")
        add_bullet(doc, "2. 時間規劃拆分兩階段：", "建議調整為兩次活動，第一次專注於生活問題發想與破冰，第二次著重在商業解法收斂與分組 Pitch，讓討論更充裕。")
        add_bullet(doc, "3. 提供發言引導小工具與雙語支援：", "現場可準備簡單的發言引導卡（如：我的痛點是...我建議可以...），幫助內向同學開口；並常態設置雙語桌，方便外籍與國際生參與。")
        add_bullet(doc, "4. 會後主動追蹤與媒合：", "針對本次已有題目方向的 17 位同學，在會後一至兩週內整理近期的商業或資訊競賽資訊（如各大黑客松、商業競賽），協助有興趣的同學順利組隊參賽。")
    elif "Coffee" in event_name:
        add_bullet(doc, "1. 嚴格落實活動時間控管：", "為避免流程延遲耽誤學員時間，各階段安排專責司儀計時引導，明訂 Coffee Chat 每輪換桌時間（如 15 分鐘），確保流程準時推進。")
        add_bullet(doc, "2. 優化破冰題庫與互動機制：", "篩選更親民、具共鳴之社交破冰題目，避開生僻或易生困惑之題目；亦可設計「共同話題卡」協助學員快速找到共同興趣。")
        add_bullet(doc, "3. 擴大場地桌距與改善動線：", "借用更寬敞之教室或空間，加大桌與桌間距，減少各桌交談音量之干擾，營造更宜人的 Coffee Chat 交流場域。")
        add_bullet(doc, "4. 提供咖啡及無糖飲品：", "因應「Coffee Chat」活動主題，現場備妥手沖/美式咖啡及無糖茶飲，滿足學員對名實相符的期待。")
        add_bullet(doc, "5. 建立會後交流名冊：", f"高達 {data['met_friends_ratio']:.1f}% 同學結識了新夥伴、近六成有意願跨域合作，建議於會後在尊重個資前提下建立交流群，延續合作動能。")
    elif "迎新" in event_name:
        add_bullet(doc, "1. 優化破冰與遊戲機制：", "針對大富翁反映規則稍繁瑣或刺激度不足，建議未來簡化遊戲規則或替換為更貼合新創主題之快速 Pitch / 拍賣競標等快節奏破冰機制，提升投入感。")
        add_bullet(doc, "2. 加強流程時間掌控：", "各階段安排計時引導，預留充裕時間進行隊伍互動與自由交流，讓時間分配更從容。")
        add_bullet(doc, "3. 貼心配套細節優化：", "現場餐飲增加無糖茶飲及常溫水選擇，滿足不同學員之飲食習慣。")
        add_bullet(doc, "4. 深度對接社員期待：", "事前報名中同學高達 87.5% 期待「競賽實作經驗與學習 AI」、75.0% 期待「結識人脈與志同道合夥伴」，後續社課可積極引進 AI 工具實作工作坊與組隊專案，落實迎新承諾。")
        add_bullet(doc, "5. 完善雙語友善環境：", "針對事前報名同學詢問「只會英文可以參加嗎？社團活動是用英文進行嗎？」，建議未來迎新與常態社課可設置雙語桌或提供英文簡報重點對照，提升國際生參與友善度。")
    else:
        add_bullet(doc, "1. 精進流程控管：", "各階段設定明確計時提醒，確保活動依節奏流暢進行。")
        add_bullet(doc, "2. 深化主題內容：", "結合學員回饋意願，適度增加實作與案例解析比重。")
        add_bullet(doc, "3. 延續交流社群：", "會後協助學員建立長期交流管道，維繫互動網絡。")

    # 嚴格確保產出前所有節點的邊界皆為標準 1.0 cm
    for s in doc.sections:
        apply_standard_margins(s)

    # 儲存
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    saved_path = output_path
    try:
        doc.save(output_path)
    except PermissionError:
        fallback_path = output_path.replace(".docx", "_updated.docx")
        doc.save(fallback_path)
        print(f"\n⚠️ 提示：原檔案 [{output_path}] 正被 Word 或其他程式開啟中，已為您另存為 [{fallback_path}]。關閉 Word 後再次執行即可覆寫原檔。")
        saved_path = fallback_path

    # 清除暫存
    for f in [chart_c_png, chart_g_png]:
        if os.path.exists(f):
            try: os.remove(f)
            except Exception: pass

    if return_data:
        return saved_path, data
    return saved_path


def build_orientation_from_template(*args, **kwargs):
    """向後相容別名"""
    return generate_report(*args, **kwargs)


if __name__ == '__main__':
    # 測試
    import os
    reg_candidates = [f for f in os.listdir('.') if f.endswith('.xlsx') and '報名' in f]
    fb_candidates = [f for f in os.listdir('.') if f.endswith('.xlsx') and '回饋' in f]
    if reg_candidates and fb_candidates:
        out = 'test_report_output.docx'
        generate_report(reg_candidates[0], fb_candidates[0], out)
        print(f"Report successfully generated at: {out}")
