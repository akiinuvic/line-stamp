"""
沖縄振興特定事業推進費 補助金申請 事業概要フロー生成スクリプト
Generates an A3 landscape Excel one-pager flowchart.

Usage:
  python generate_flow.py
  python generate_flow.py --data flow_data.json          # browser JSON export
  python generate_flow.py --output my_output.xlsx
"""

import argparse
import json
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter, range_boundaries

# ---------------------------------------------------------------------------
# Color palette
# ---------------------------------------------------------------------------
C = {
    "title_bg":   "1F3864",  # dark navy
    "title_fg":   "FFFFFF",
    "prob_h":     "C00000",  # red-ish  - 課題
    "prob_b":     "FFE7E7",
    "purp_h":     "833C00",  # brown    - 目的
    "purp_b":     "FFF2CC",
    "act_h":      "375623",  # dark green - 取組
    "act_b":      "E2EFDA",
    "leg_h":      "1F3864",  # navy    - 要件
    "leg_b":      "DEEAF1",
    "cost_h":     "833C00",  # brown   - 費用
    "cost_b":     "FCE4D6",
    "eval_h":     "4472C4",  # blue    - 評価
    "eval_b":     "EEF4FF",
    "vis_h":      "375623",  # green   - 将来像
    "vis_b":      "E8F5E9",
    "law_bg":     "F2F2F2",
    "law_hdr":    "595959",
    "arrow":      "404040",
    "border":     "595959",
    "WHITE":      "FFFFFF",
    "BLACK":      "000000",
    "gray_line":  "BFBFBF",
}

FONT_JP = "Yu Gothic"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def fill(hex_color: str) -> PatternFill:
    return PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")


def border(color: str = "595959", style: str = "thin") -> Border:
    s = Side(style=style, color=color)
    return Border(left=s, right=s, top=s, bottom=s)


def font(bold=False, size=10, color="000000") -> Font:
    return Font(bold=bold, size=size, color=color, name=FONT_JP)


def align(h="center", v="center", wrap=True) -> Alignment:
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


def apply_outer_border(ws, cell_range: str, color: str = "595959", style: str = "thin"):
    """Apply a clean outer border to a merged/unmerged range."""
    min_col, min_row, max_col, max_row = range_boundaries(cell_range)
    thick = Side(style=style, color=color)
    no_s = Side(style=None)
    for row in ws.iter_rows(min_row=min_row, max_row=max_row,
                            min_col=min_col, max_col=max_col):
        for c in row:
            l = thick if c.column == min_col else no_s
            r = thick if c.column == max_col else no_s
            t = thick if c.row == min_row else no_s
            b = thick if c.row == max_row else no_s
            c.border = Border(left=l, right=r, top=t, bottom=b)


def block(ws, cell_range: str, text: str, bg: str, fg: str = "000000",
          size: int = 10, bold: bool = False,
          h: str = "center", v: str = "center", wrap: bool = True,
          border_color: str = "595959", border_style: str = "thin"):
    ws.merge_cells(cell_range)
    first = cell_range.split(":")[0]
    c = ws[first]
    c.value = text
    c.fill = fill(bg)
    c.font = Font(bold=bold, size=size, color=fg, name=FONT_JP)
    c.alignment = Alignment(horizontal=h, vertical=v, wrap_text=wrap)
    apply_outer_border(ws, cell_range, border_color, border_style)


# ---------------------------------------------------------------------------
# Layout constants  (all indices are 1-based Excel rows/cols)
# ---------------------------------------------------------------------------
#
#  Column map (1-based):
#   1       : left margin
#   2-4     : S1 (課題)         width 14 each → 3 cols
#   5       : arrow             width 3
#   6-8     : S2 (目的)
#   9       : arrow
#   10-12   : S3 (取組)
#   13      : arrow
#   14-16   : S4 (要件)
#   17      : arrow
#   18-20   : S5 (費用)
#   21      : arrow
#   22-24   : S6 (評価)
#   25      : arrow
#   26-28   : S7 (将来像)
#   29      : right margin
#
# Row map:
#   1       : title             h=36
#   2       : spacer            h=6
#   3       : section header    h=22
#   4       : section question  h=18
#   5-12    : content           h=15 each (8 rows)
#   13      : spacer            h=8
#   14-15   : legal basis hdr   h=18
#   16-22   : legal basis body  h=15 each
#   23      : bottom margin     h=6
#

SEC_COLS = {
    "S1": ("B", "D"),
    "S2": ("F", "H"),
    "S3": ("J", "L"),
    "S4": ("N", "P"),
    "S5": ("R", "T"),
    "S6": ("V", "X"),
    "S7": ("Z", "AB"),
}

ARROW_COLS = ["E", "I", "M", "Q", "U", "Y"]

# Sections definition
SECTIONS = [
    {
        "key": "S1",
        "label": "① 課題の把握",
        "question": "なぜやるのか\n（現状・問題点）",
        "h_color": "C00000", "b_color": "FFE7E7",
        "fields": [
            ("地域課題（社会的背景）",
             "例）〇〇市では□□の問題が深刻化しており…"),
            ("現状のギャップ",
             "例）現在は△△の状態であり、理想とのギャップが□□千円・□□人分…"),
            ("課題の特殊性\n（沖縄固有の事情）",
             "例）離島・遠隔地ゆえに本土と比較して□□の格差が存在する…"),
        ],
    },
    {
        "key": "S2",
        "label": "② 事業目的・概要",
        "question": "何のためにやるのか\n（目的・必要性）",
        "h_color": "7B3F00", "b_color": "FFF8E7",
        "fields": [
            ("事業名",
             "令和□年度 ○○推進事業"),
            ("事業の必要性",
             "例）〇〇振興基本方針 第□章に掲げる施策の実現に向け…"),
            ("事業目的",
             "例）本事業は□□を通じて○○の向上を図り、…"),
            ("事業期間",
             "令和□年□月 ～ 令和□年□月（□か年）"),
        ],
    },
    {
        "key": "S3",
        "label": "③ 取組内容（手段）",
        "question": "どのようにやるのか\n（具体的な取組・実施体制）",
        "h_color": "375623", "b_color": "E2EFDA",
        "fields": [
            ("取組①",
             "【委託】〇〇調査・企画業務\n単価□千円×□回＝□千円"),
            ("取組②",
             "【旅費】先進地視察\n単価□千円×□人×□回＝□千円"),
            ("取組③",
             "【印刷】広報・啓発資料作成\n□千部×□円＝□千円"),
            ("実施体制",
             "主体：〇〇市□□課\n協力：民間事業者□社"),
        ],
    },
    {
        "key": "S4",
        "label": "④ 要件確認（適法性）",
        "question": "法令・要綱に\n沿っているか",
        "h_color": "1F3864", "b_color": "DEEAF1",
        "fields": [
            ("沖縄振興に資する\n事業であること\n（第４条第１項第１号ア前段）",
             "□ 該当\n根拠："),
            ("沖縄の特殊性に起因\nする事業であること\n（第４条第１項第１号ア後段）",
             "□ 該当\n根拠："),
            ("公共の利益に資する\n事業であること\n（第４条第１項第１号イ）",
             "□ 該当\n根拠："),
            ("先導性・広域性を有\nする事業であること\n（第４条第１項第２号）",
             "□ 先導性  □ 広域性\n根拠："),
        ],
    },
    {
        "key": "S5",
        "label": "⑤ 事業費・補助額",
        "question": "いくら必要か\n（費用・行程）",
        "h_color": "833C00", "b_color": "FCE4D6",
        "fields": [
            ("総事業費（千円）",
             "□,□□□ 千円"),
            ("補助対象経費（千円）",
             "□,□□□ 千円"),
            ("補助額（千円）",
             "□,□□□ 千円"),
            ("補助率",
             "国□/□ ＋ 市町村□/□\n（民間負担 □/□）"),
            ("年間行程（主な実施月）",
             "4月：企画  6-8月：実施\n10月：中間報告  3月：完了"),
        ],
    },
    {
        "key": "S6",
        "label": "⑥ 効果検証（評価指標）",
        "question": "どう測るか\n（定量的アウトカム指標）",
        "h_color": "2F5496", "b_color": "EEF4FF",
        "fields": [
            ("指標①（定量）",
             "指標名：\n基準値：□　目標値：□\n達成予定年度：令和□年度"),
            ("指標②（定量）",
             "指標名：\n基準値：□　目標値：□\n達成予定年度：令和□年度"),
            ("数値目標の設定根拠",
             "例）□□調査（令和□年）によれば全国平均は□□であり、本事業により□年で□□まで引き上げる。"),
        ],
    },
    {
        "key": "S7",
        "label": "⑦ 将来像・アウトカム",
        "question": "どうなるのか\n（達成後の姿・波及効果）",
        "h_color": "205867", "b_color": "E0F4F8",
        "fields": [
            ("短期アウトカム\n（事業終了時）",
             "例）〇〇の体制が整備され、□□件の□□が実現する。"),
            ("中長期アウトカム\n（事業終了後3-5年）",
             "例）□□率が□□%向上し、地域の□□力が強化される。"),
            ("地域・社会への\n波及効果",
             "例）隣接市町村への水平展開、□□産業の振興、雇用□□名増加等。"),
        ],
    },
]

LAW_REFS = [
    ("①",  "補助金等に係る予算の\n執行の適正化に関する法律\n（補助金適正化法）",
     "補助事業者の義務・\n補助金の目的外使用禁止等\nの遵守事項を定める"),
    ("②",  "沖縄振興基本方針\n（閣議決定）",
     "沖縄振興の基本的な\n方向性・施策の優先順位を\n示す上位方針"),
    ("③",  "沖縄振興特別措置法\n（及び施行令）",
     "第４条各号の要件（振興・\n特殊性・公益性・先導性）\nの法的根拠"),
    ("④",  "沖縄振興特定事業推進費\n市町村補助金交付要綱",
     "市町村申請に係る\n補助対象・補助率・\n手続きを規定"),
    ("⑤",  "沖縄振興特定事業推進費\n民間補助金交付要綱",
     "民間事業者申請に係る\n補助対象・補助率・\n手続きを規定"),
]


# ---------------------------------------------------------------------------
# Build workbook
# ---------------------------------------------------------------------------
def build_excel(output_path: str, values: dict = None):
    """
    values: optional dict from browser JSON export
      { "S1": {"s1_f1": "text", ...}, "S2": {...}, ... }
    Field IDs follow the pattern  s{section_num}_f{field_index_1based}.
    """
    def field_value(sec_key: str, field_idx: int, default: str) -> str:
        if not values:
            return default
        fid = f"{sec_key.lower()}_f{field_idx + 1}"
        return values.get(sec_key, {}).get(fid, default)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "事業概要フロー"

    # ---- Page setup (A3 landscape) ----
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = 8          # 8 = A3
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_margins.left = 0.35
    ws.page_margins.right = 0.35
    ws.page_margins.top = 0.35
    ws.page_margins.bottom = 0.35
    ws.page_margins.header = 0.1
    ws.page_margins.footer = 0.1

    # ---- Column widths ----
    ws.column_dimensions["A"].width = 1.5   # left margin
    ws.column_dimensions["AC"].width = 1.5  # right margin
    for arrow_col in ARROW_COLS:
        ws.column_dimensions[arrow_col].width = 3
    # Section columns – 3 cols each, width 13
    for s in SECTIONS:
        sc, ec = SEC_COLS[s["key"]]
        sc_idx = ord(sc) - ord("A") + 1
        for ci in range(sc_idx, sc_idx + 3):
            ws.column_dimensions[get_column_letter(ci)].width = 13

    # ---- Row heights ----
    ws.row_dimensions[1].height = 38    # title
    ws.row_dimensions[2].height = 6     # spacer
    ws.row_dimensions[3].height = 24    # section header
    ws.row_dimensions[4].height = 20    # section question
    for r in range(5, 13):
        ws.row_dimensions[r].height = 15  # content rows (8 rows)
    ws.row_dimensions[12].height = 20
    ws.row_dimensions[13].height = 10   # spacer
    ws.row_dimensions[14].height = 22   # law header
    ws.row_dimensions[15].height = 16   # law sub-header
    for r in range(16, 25):
        ws.row_dimensions[r].height = 15
    ws.row_dimensions[24].height = 20
    ws.row_dimensions[25].height = 8    # bottom margin

    # ================================================================
    # ROW 1: Title
    # ================================================================
    block(ws, "B1:AB1",
          "令和８年度 沖縄振興特定事業推進費 補助金申請  事業概要フロー",
          bg=C["title_bg"], fg=C["title_fg"],
          size=16, bold=True, border_color="1F3864", border_style="medium")

    # ================================================================
    # ROW 3: Section header boxes + arrows
    # ================================================================
    for s in SECTIONS:
        sc, ec = SEC_COLS[s["key"]]
        block(ws, f"{sc}3:{ec}3",
              s["label"],
              bg=s["h_color"], fg="FFFFFF",
              size=11, bold=True,
              border_color=s["h_color"], border_style="medium")

    # Arrow cells row 3
    for ac in ARROW_COLS:
        c = ws[f"{ac}3"]
        c.value = "▶"
        c.font = Font(size=18, color=C["arrow"], name=FONT_JP, bold=True)
        c.alignment = align()
        c.fill = fill("FFFFFF")

    # ================================================================
    # ROW 4: Sub-question row
    # ================================================================
    for s in SECTIONS:
        sc, ec = SEC_COLS[s["key"]]
        block(ws, f"{sc}4:{ec}4",
              s["question"],
              bg=s["b_color"], fg="404040",
              size=9, bold=False,
              border_color=s["h_color"])

    for ac in ARROW_COLS:
        c = ws[f"{ac}4"]
        c.fill = fill("FFFFFF")

    # ================================================================
    # ROWS 5-12: Content fields
    # ================================================================
    # We distribute each section's fields across the 8 available rows
    # (rows 5 to 12). If a section has N fields, each gets floor(8/N) rows.
    CONTENT_START = 5
    CONTENT_END = 12

    def write_fields(sec, sc, ec):
        fields = sec["fields"]
        n = len(fields)
        total_rows = CONTENT_END - CONTENT_START + 1  # 8
        rows_per_field = total_rows // n
        remainder = total_rows - rows_per_field * n

        row = CONTENT_START
        for i, (label, placeholder) in enumerate(fields):
            extra = 1 if i < remainder else 0
            span = rows_per_field + extra
            end_row = row + span - 1
            cell_range = f"{sc}{row}:{ec}{end_row}"

            value = field_value(sec["key"], i, placeholder)
            text = f"【{label}】\n{value}"
            block(ws, cell_range, text,
                  bg=sec["b_color"], fg="303030",
                  size=8, bold=False, h="left", v="top",
                  border_color=sec["h_color"])
            row = end_row + 1

    for s in SECTIONS:
        sc, ec = SEC_COLS[s["key"]]
        write_fields(s, sc, ec)

    # Arrow cells rows 5-12
    for ac in ARROW_COLS:
        ws.merge_cells(f"{ac}5:{ac}12")
        c = ws[f"{ac}5"]
        c.value = "▶"
        c.font = Font(size=22, color=C["arrow"], name=FONT_JP, bold=True)
        c.alignment = align()
        c.fill = fill("FFFFFF")

    # ================================================================
    # ROW 13: Spacer / connector line
    # ================================================================
    block(ws, "B13:AB13", "", bg="E8E8E8",
          border_color=C["gray_line"], border_style="hair")

    # ================================================================
    # ROW 14-24: Legal basis section
    # ================================================================
    # Header
    block(ws, "B14:AB14",
          "根拠法令・要綱（本フローが依拠する法制度）",
          bg="2E4057", fg="FFFFFF",
          size=12, bold=True,
          border_color="2E4057", border_style="medium")

    # 5 law reference columns – split the 27 section columns evenly
    # B-AB = cols 2-28 = 27 cols  →  each law ≈ 5 cols, with 2 leftover
    # Layout: B-F | G-K | L-P | Q-U | V-AB
    law_ranges = ["B15:F24", "G15:K24", "L15:P24", "Q15:U24", "V15:AB24"]
    law_bg_colors = ["FFF8E7", "E8F5E9", "EEF4FF", "FCE4D6", "FFE7E7"]
    law_hdr_colors = ["7B3F00", "375623", "1F3864", "833C00", "C00000"]

    for i, (num, name, desc) in enumerate(LAW_REFS):
        sr = law_ranges[i]
        # Sub-header row 15
        sr_parts = sr.split(":")
        sc15 = sr_parts[0][:sr_parts[0].index("1")]  # strip row number → col letter(s)
        # parse col from range like "B15:F24"
        sc_str = sr.split(":")[0]
        sc_col = ""
        for ch in sc_str:
            if ch.isalpha():
                sc_col += ch
        ec_str = sr.split(":")[1]
        ec_col = ""
        for ch in ec_str:
            if ch.isalpha():
                ec_col += ch

        hdr_range = f"{sc_col}15:{ec_col}15"
        body_range = f"{sc_col}16:{ec_col}24"

        # Number badge + name
        block(ws, hdr_range,
              f"{num}  {name}",
              bg=law_hdr_colors[i], fg="FFFFFF",
              size=9, bold=True,
              border_color=law_hdr_colors[i], border_style="medium")

        block(ws, body_range,
              desc,
              bg=law_bg_colors[i], fg="303030",
              size=9, bold=False, h="left", v="top",
              border_color=law_hdr_colors[i])

    # ================================================================
    # Final tweaks: freeze top row, zoom
    # ================================================================
    ws.freeze_panes = "B2"
    ws.sheet_view.zoomScale = 75

    wb.save(output_path)
    print(f"Saved: {output_path}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="事業概要フロー Excel 生成")
    parser.add_argument("--data",   metavar="JSON", help="ブラウザからエクスポートした flow_data.json")
    parser.add_argument("--output", metavar="XLSX", default="/home/user/line-stamp/okinawa_subsidy_flow.xlsx",
                        help="出力ファイルパス（デフォルト: okinawa_subsidy_flow.xlsx）")
    args = parser.parse_args()

    override = None
    if args.data:
        with open(args.data, encoding="utf-8") as f:
            override = json.load(f)

    build_excel(args.output, override)
