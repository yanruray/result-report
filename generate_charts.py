import os
import math
from PIL import Image, ImageDraw, ImageFont

def get_font(size, bold=False):
    candidate_paths = []
    if bold:
        candidate_paths = [
            "C:/Windows/Fonts/msjhbd.ttc",
            "C:/Windows/Fonts/msjh.ttc",
            "/System/Library/Fonts/PingFang.ttc",
            "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
            "/usr/share/fonts/truetype/noto/NotoSansCJK-Bold.ttc"
        ]
    else:
        candidate_paths = [
            "C:/Windows/Fonts/msjh.ttc",
            "/System/Library/Fonts/PingFang.ttc",
            "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
            "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc"
        ]
    for p in candidate_paths:
        if os.path.exists(p):
            try:
                return ImageFont.truetype(p, size)
            except Exception:
                pass
    try:
        return ImageFont.truetype("arial.ttf", size)
    except Exception:
        return ImageFont.load_default()

def draw_pie_chart(data, title, output_path, width=750, height=480, colors=None):
    # 4x super-sampling for crisp antialiasing
    scale = 4
    W = width * scale
    H = height * scale
    img = Image.new("RGBA", (W, H), (255, 255, 255, 255))
    draw = ImageDraw.Draw(img)

    # Fonts with robust fallback
    font_title = get_font(26 * scale, bold=True)
    font_bold = get_font(22 * scale, bold=True)
    font_regular = get_font(20 * scale, bold=False)
    font_small = get_font(17 * scale, bold=False)

    # Colors matching master template palette
    if not colors:
        colors = [
            (30, 57, 137),    # Navy #1E3989
            (36, 98, 235),    # Blue #2462EB
            (96, 165, 250),   # Sky Blue #60A5FA
            (245, 158, 11),   # Amber #F59E0B
            (16, 185, 129),   # Emerald #10B981
        ]

    # Draw Title
    title_bbox = draw.textbbox((0, 0), title, font=font_title)
    title_w = title_bbox[2] - title_bbox[0]
    draw.text(((W - title_w) // 2, 28 * scale), title, fill=(30, 57, 137), font=font_title)

    # Center of pie (positioned on the left side)
    cx = int(W * 0.34)
    cy = int(H * 0.56)
    r = int(H * 0.35)
    bbox = [cx - r, cy - r, cx + r, cy + r]

    total = sum(item[1] for item in data)
    start_angle = -90.0

    # Draw slices
    for idx, (label, val) in enumerate(data):
        fraction = val / total
        sweep = fraction * 360.0
        end_angle = start_angle + sweep
        color = colors[idx % len(colors)]
        
        draw.pieslice(bbox, start=start_angle, end=end_angle, fill=color, outline=(255, 255, 255), width=4 * scale)
        start_angle = end_angle

    # Donut hole
    inner_r = int(r * 0.54)
    inner_bbox = [cx - inner_r, cy - inner_r, cx + inner_r, cy + inner_r]
    draw.ellipse(inner_bbox, fill=(255, 255, 255), outline=(255, 255, 255))
    
    # Center text in donut
    center_lbl = "總計"
    c_lbl_bbox = draw.textbbox((0, 0), center_lbl, font=font_small)
    c_lbl_w = c_lbl_bbox[2] - c_lbl_bbox[0]
    draw.text((cx - c_lbl_w // 2, cy - 22 * scale), center_lbl, fill=(99, 116, 138), font=font_small)

    center_val = f"{total} 人"
    c_val_bbox = draw.textbbox((0, 0), center_val, font=font_bold)
    c_val_w = c_val_bbox[2] - c_val_bbox[0]
    draw.text((cx - c_val_w // 2, cy + 2 * scale), center_val, fill=(30, 57, 137), font=font_bold)

    # Draw Legend on the right (vertically centered with pie)
    lx = int(W * 0.62)
    line_h = 46 * scale
    total_legend_h = len(data) * line_h
    ly_start = cy - (total_legend_h // 2)

    for idx, (label, val) in enumerate(data):
        pct = (val / total) * 100
        cur_y = ly_start + idx * line_h
        color = colors[idx % len(colors)]

        # Color indicator circle / rounded box
        box_sz = 14 * scale
        draw.ellipse([lx, cur_y + 4 * scale, lx + box_sz, cur_y + 4 * scale + box_sz], fill=color)

        # Text: Label
        lbl_txt = f"{label}"
        draw.text((lx + box_sz + 12 * scale, cur_y), lbl_txt, fill=(30, 41, 58), font=font_regular)

        # Text: Value & Percentage
        val_txt = f"{val} 人 ({pct:.1f}%)"
        draw.text((lx + box_sz + 12 * scale, cur_y + 22 * scale), val_txt, fill=(99, 116, 138), font=font_small)

    # Subtle card border around chart
    draw.rounded_rectangle([2 * scale, 2 * scale, W - 3 * scale, H - 3 * scale], radius=8 * scale, outline=(226, 232, 240), width=2 * scale)

    # Downsample with Lanczos
    final_img = img.resize((width, height), Image.Resampling.LANCZOS)
    final_img.save(output_path, dpi=(300, 300))
    print(f"Successfully saved {output_path}")

if __name__ == "__main__":
    # Generate College Chart
    college_data = [
        ("商管學院", 15),
        ("工學院", 6),
        ("國際事務學院", 1),
    ]
    draw_pie_chart(college_data, "學院結構分布", "chart_college.png", width=650, height=420)

    # Generate Grade Chart
    grade_data = [
        ("大學一年級", 9),
        ("大學二年級", 6),
        ("大學三年級", 6),
        ("大學四年級以上", 1),
    ]
    draw_pie_chart(grade_data, "年級結構分布", "chart_grade.png", width=650, height=420)
