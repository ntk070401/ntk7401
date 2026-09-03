# -*- coding: utf-8 -*-
"""Create a grid table (detail lines + text notes) in the active view
from data copied from Excel (Ctrl+C on a selected range).

Reads Excel's HTML clipboard format to preserve:
  - merged cells (colspan/rowspan)
  - text alignment
  - actual column widths / row heights (from Excel's <col>/<tr> style)

Table geometry is scaled by the current view's Scale so the printed
size on paper matches what you type in inches, regardless of the
view's drawing scale at creation time. (Note: this is a one-time
scale-correct creation; if you change the view scale afterwards,
the table will NOT auto-resize - you would need to recreate it.)

Falls back to plain tab-separated text (uniform cell size, no merge,
no alignment) if Excel HTML clipboard data is not available.
"""
__title__ = "Table\nExcel to Revit"

import re
import clr

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('System')

from System import IO
from System.Net import WebUtility
from System.Drawing import Point
from System.Windows.Forms import (
    Clipboard, Form, Label as WinLabel, TextBox as WinTextBox,
    ComboBox as WinComboBox, Button as WinButton, DialogResult,
    FormStartPosition, ComboBoxStyle, FormBorderStyle
)

from Autodesk.Revit.DB import (
    XYZ, Line, Transaction, TextNote, TextNoteOptions, TextNoteType,
    FilteredElementCollector, HorizontalTextAlignment, VerticalTextAlignment,
    ViewType, BuiltInCategory, GraphicsStyleType
)

from pyrevit import forms, revit, script

doc = revit.doc
uidoc = revit.uidoc
view = doc.ActiveView
output = script.get_output()

INCH_TO_FT = 1.0 / 12.0
PT_TO_INCH = 1.0 / 72.0


def inch_to_ft(value_inch):
    """Convert an inch value to Revit internal feet."""
    return value_inch * INCH_TO_FT


def pt_to_ft(value_pt):
    """Convert a point (1/72 inch) value to Revit internal feet."""
    return inch_to_ft(value_pt * PT_TO_INCH)


# ---------------------------------------------------------------------------
# 0. Validate active view
# ---------------------------------------------------------------------------
allowed_view_types = [
    ViewType.FloorPlan, ViewType.CeilingPlan, ViewType.Elevation,
    ViewType.Section, ViewType.DraftingView, ViewType.EngineeringPlan
]
if view.ViewType not in allowed_view_types:
    forms.alert(
        'Current view type is not supported for Detail Lines / Text Notes.\n'
        'Please switch to a Plan, Section, Elevation or Drafting view.',
        exitscript=True
    )

view_scale = view.Scale if view.Scale else 1
output.print_md('**View scale:** 1:{}'.format(view_scale))

# ---------------------------------------------------------------------------
# 1. Read table data from clipboard
#    Preferred: "HTML Format" (Excel writes real table markup with
#    colspan/rowspan/style/col-width/row-height so we can rebuild merges,
#    alignment, and true column/row sizes).
#    Fallback: plain text, tab-separated (uniform size, no merge/align).
# ---------------------------------------------------------------------------

def get_clipboard_html():
    """Return (fragment, raw_html) from clipboard 'HTML Format', or (None, None).
    fragment = content between <!--StartFragment--> / <!--EndFragment-->
    raw_html = the full HTML document (needed to read the <style> block,
    which sits before StartFragment and holds Excel's alignment rules)."""
    if not Clipboard.ContainsData('HTML Format'):
        return None, None
    raw_obj = Clipboard.GetData('HTML Format')
    if raw_obj is None:
        return None, None
    if isinstance(raw_obj, IO.Stream):
        reader = IO.StreamReader(raw_obj)
        raw_html = reader.ReadToEnd()
    else:
        raw_html = str(raw_obj)

    start_marker = '<!--StartFragment-->'
    end_marker = '<!--EndFragment-->'
    start_idx = raw_html.find(start_marker)
    end_idx = raw_html.find(end_marker)
    if start_idx != -1 and end_idx != -1:
        fragment = raw_html[start_idx + len(start_marker):end_idx]
    else:
        fragment = raw_html
    return fragment, raw_html

def parse_css_classes(raw_html):
    """Parse the <style> block for '.classname { ... }' rules and extract
    text-align / vertical-align declarations, keyed by class name.
    Excel typically stores cell alignment here (e.g. .xl65, .xl66),
    not as inline style on the <td> itself."""
    style_match = re.search(r'<style[^>]*>(.*?)</style>', raw_html, re.S | re.I)
    if not style_match:
        return {}
    style_block = style_match.group(1)

    classes = {}
    for rule_m in re.finditer(r'\.([\w\-]+)\s*\{([^}]*)\}', style_block, re.S | re.I):
        class_name = rule_m.group(1)
        body = rule_m.group(2)

        align_m = re.search(r'text-align\s*:\s*(\w+)', body, re.I)
        valign_m = re.search(r'vertical-align\s*:\s*(\w+)', body, re.I)

        entry = {}
        if align_m and align_m.group(1).lower() != 'general':
            entry['align'] = align_m.group(1).lower()
        if valign_m:
            entry['valign'] = valign_m.group(1).lower()

        if entry:
            classes[class_name] = entry
    return classes

def strip_html_tags(inner_html):
    """Remove HTML tags and decode entities to get plain cell text.

    Excel's HTML clipboard source often contains literal line breaks and
    extra whitespace purely for pretty-printing the markup (not meant as
    real line breaks in the cell). Only explicit <br> tags represent an
    intentional line break; everything else is collapsed into a single
    space, matching how a browser would render the same HTML.
    """
    # placeholder so <br> survives whitespace-collapsing below
    text = re.sub(r'<br\s*/?>', '\x00LB\x00', inner_html, flags=re.I)
    text = re.sub(r'<[^>]+>', '', text)
    text = WebUtility.HtmlDecode(text)

    # collapse all whitespace runs (spaces, tabs, and literal newlines
    # that were only present for HTML source formatting) into one space
    text = re.sub(r'\s+', ' ', text)

    # restore the real, intentional line breaks
    text = text.replace('\x00LB\x00', '\n')

    # trim each line individually, keep intentional line breaks
    lines = [line.strip() for line in text.split('\n')]
    return '\n'.join(lines).strip()


def parse_col_widths_pt(html_fragment, n_cols):
    """Parse <col style='width:XXpt'> tags (honoring 'span') into a list
    of per-column widths in points. Returns None if no <col> tags found."""
    col_tags = re.findall(r'<col([^>]*)>', html_fragment, re.I)
    if not col_tags:
        return None

    widths = []
    for attrs in col_tags:
        span_m = re.search(r'span\s*=\s*"?(\d+)"?', attrs, re.I)
        span = int(span_m.group(1)) if span_m else 1
        width_m = re.search(r'width\s*:\s*([\d.]+)pt', attrs, re.I)
        width_pt = float(width_m.group(1)) if width_m else None
        widths.extend([width_pt] * span)

    known = [w for w in widths if w is not None]
    if not known:
        return None
    avg = sum(known) / len(known)
    widths = [w if w is not None else avg for w in widths]

    if len(widths) < n_cols:
        widths.extend([avg] * (n_cols - len(widths)))
    elif len(widths) > n_cols:
        widths = widths[:n_cols]
    return widths


def parse_html_table(html_fragment, css_classes):
    """Parse <tr>/<td> rows directly from the fragment into placed cells,
    plus per-row heights and per-column widths (in points) when available.
    Alignment is read from inline style first, falling back to the
    cell's CSS class (css_classes), since Excel usually stores alignment
    at the class level rather than inline.
    """
    row_iter = re.findall(r'<tr([^>]*)>(.*?)</tr>', html_fragment, re.S | re.I)
    if not row_iter:
        return None

    placed_cells = []
    occupied = {}
    max_cols = 0
    row_heights_pt = []

    for r, (tr_attrs, row_html) in enumerate(row_iter):
        height_m = re.search(r'height\s*:\s*([\d.]+)pt', tr_attrs, re.I)
        row_heights_pt.append(float(height_m.group(1)) if height_m else None)

        cell_matches = re.findall(
            r'<(td|th)([^>]*)>(.*?)</\1>', row_html, re.S | re.I
        )
        c = 0
        for tag, attrs, inner in cell_matches:
            while occupied.get((r, c)):
                c += 1

            colspan_m = re.search(r'colspan\s*=\s*"?(\d+)"?', attrs, re.I)
            rowspan_m = re.search(r'rowspan\s*=\s*"?(\d+)"?', attrs, re.I)
            colspan = int(colspan_m.group(1)) if colspan_m else 1
            rowspan = int(rowspan_m.group(1)) if rowspan_m else 1

            class_m = re.search(r'class\s*=\s*["\']?([\w\-]+(?:\s+[\w\-]+)*)["\']?', attrs, re.I)
            cell_classes = class_m.group(1).split() if class_m else []

            align_m = re.search(r'text-align\s*:\s*(\w+)', attrs, re.I)
            align = align_m.group(1).lower() if align_m else None
            if align == 'general':
                align = None

            valign_m = re.search(r'vertical-align\s*:\s*(\w+)', attrs, re.I)
            if not valign_m:
                valign_m = re.search(r'valign\s*=\s*"?(\w+)"?', attrs, re.I)
            valign = valign_m.group(1).lower() if valign_m else None

            # fall back to CSS class rules (Excel usually puts alignment here)
            for cls in cell_classes:
                class_info = css_classes.get(cls)
                if not class_info:
                    continue
                if align is None and 'align' in class_info:
                    align = class_info['align']
                if valign is None and 'valign' in class_info:
                    valign = class_info['valign']

            text = strip_html_tags(inner)

            for dr in range(rowspan):
                for dc in range(colspan):
                    occupied[(r + dr, c + dc)] = True

            placed_cells.append({
                'row': r, 'col': c,
                'rowspan': rowspan, 'colspan': colspan,
                'text': text, 'align': align, 'valign': valign
            })
            c += colspan
        max_cols = max(max_cols, c)

    n_rows = len(row_iter)
    n_cols = max_cols

    known_rh = [h for h in row_heights_pt if h is not None]
    if known_rh:
        avg_rh = sum(known_rh) / len(known_rh)
        row_heights_pt = [h if h is not None else avg_rh for h in row_heights_pt]
    else:
        row_heights_pt = None

    col_widths_pt = parse_col_widths_pt(html_fragment, n_cols)

    return placed_cells, n_rows, n_cols, col_widths_pt, row_heights_pt

def parse_plain_text_table():
    """Fallback: parse tab-separated plain text clipboard content.
    No merge/alignment/size information; every cell is a single 1x1 cell."""
    if not Clipboard.ContainsText():
        forms.alert(
            'Clipboard has no usable data.\n'
            'Copy a range from Excel first (Ctrl+C).',
            exitscript=True
        )
    raw_text = Clipboard.GetText()
    raw_rows = raw_text.replace('\r\n', '\n').strip('\n').split('\n')
    rows_split = [row.split('\t') for row in raw_rows]
    n_rows = len(rows_split)
    n_cols = max(len(r) for r in rows_split)

    placed_cells = []
    for r, row in enumerate(rows_split):
        for c in range(n_cols):
            text = row[c] if c < len(row) else ''
            placed_cells.append({
                'row': r, 'col': c,
                'rowspan': 1, 'colspan': 1,
                'text': text, 'align': None, 'valign': None
            })
    return placed_cells, n_rows, n_cols

html_fragment, raw_html = get_clipboard_html()
css_classes = parse_css_classes(raw_html) if raw_html else {}
# --- DEBUG: kiểm tra khối <style> và các class Excel xuất ra ---
if raw_html:
    style_match = re.search(r'<style[^>]*>(.*?)</style>', raw_html, re.S | re.I)
    if style_match:
        style_preview = style_match.group(1)[:1500]
        style_preview = style_preview.replace('<', '&lt;').replace('>', '&gt;')
        output.print_md('**DEBUG <style> block (1500 ký tự đầu):**\n```\n{}\n```'.format(style_preview))
    else:
        output.print_md('**DEBUG:** Không tìm thấy khối <style> trong raw_html.')

    output.print_md('**DEBUG:** Số class parse được: {}'.format(len(css_classes)))
    for cls_name, cls_info in list(css_classes.items())[:10]:
        output.print_md('  - .{} -> {}'.format(cls_name, cls_info))

    # In luôn attrs của vài td đầu tiên (kể cả class name) để đối chiếu
    sample_tds = re.findall(r'<td([^>]*)>', html_fragment, re.I)[:6]
    output.print_md('**DEBUG: attrs của 6 <td> đầu tiên trong fragment:**')
    for attrs in sample_tds:
        output.print_md('  `{}`'.format(attrs))

parsed = parse_html_table(html_fragment, css_classes) if html_fragment else None
using_html = parsed is not None

if using_html:
    placed_cells, n_rows, n_cols, col_widths_pt, row_heights_pt = parsed
else:
    placed_cells, n_rows, n_cols = parse_plain_text_table()
    col_widths_pt = None
    row_heights_pt = None

align_count = sum(1 for c in placed_cells if c.get('align'))
valign_count = sum(1 for c in placed_cells if c.get('valign'))
output.print_md(
    '**Detected table:** {} rows x {} columns | merged cells: {} | '
    'align detected: {} ô | valign detected: {} ô'.format(
        n_rows, n_cols,
        sum(1 for c in placed_cells if c['rowspan'] > 1 or c['colspan'] > 1),
        align_count, valign_count
    )
)

# ---------------------------------------------------------------------------
# 2. Combined input dialog
#    - Fallback cell size (only used when Excel width/height not detected)
#    - Overall size multiplier (applies in both cases)
#    - Line style + text note type
# ---------------------------------------------------------------------------
lines_cat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines)
line_style_dict = {}
for subcat in lines_cat.SubCategories:
    gstyle = subcat.GetGraphicsStyle(GraphicsStyleType.Projection)
    if gstyle:
        line_style_dict[subcat.Name] = gstyle

text_types = FilteredElementCollector(doc).OfClass(TextNoteType).ToElements()
text_type_dict = {}
for tnt in text_types:
    name_param = tnt.LookupParameter('Type Name')
    name = name_param.AsString() if name_param else tnt.Name
    text_type_dict[name] = tnt


class TableSetupForm(Form):
    """WinForms dialog collecting fallback cell size, size multiplier,
    line style, and text note type."""

    def __init__(self):
        self.Text = 'Table Setup'
        self.Width = 380
        self.Height = 400
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog

        y = 15

        WinLabel(
            Text='Fallback column width (inch) - used only if Excel column widths not found:',
            Location=Point(15, y), AutoSize=False, Width=340, Height=30, Parent=self
        )
        y += 32
        self.tb_width = WinTextBox(Text='1.0', Location=Point(15, y), Width=320, Parent=self)
        y += 32

        WinLabel(
            Text='Fallback row height (inch) - used only if Excel row heights not found:',
            Location=Point(15, y), AutoSize=False, Width=340, Height=30, Parent=self
        )
        y += 32
        self.tb_height = WinTextBox(Text='0.3', Location=Point(15, y), Width=320, Parent=self)
        y += 32

        WinLabel(
            Text='Size multiplier (x) - scales the whole table up/down:',
            Location=Point(15, y), AutoSize=False, Width=340, Height=20, Parent=self
        )
        y += 22
        self.tb_multiplier = WinTextBox(Text='1.0', Location=Point(15, y), Width=320, Parent=self)
        y += 35

        WinLabel(
            Text='Cell padding (%) - text inset from cell borders:',
            Location=Point(15, y), AutoSize=False, Width=340, Height=20, Parent=self
        )
        y += 22
        self.tb_padding = WinTextBox(Text='10', Location=Point(15, y), Width=320, Parent=self)
        y += 35

        WinLabel(Text='Line Style (grid):', Location=Point(15, y), AutoSize=True, Parent=self)
        y += 20
        self.cb_line_style = WinComboBox(
            Location=Point(15, y), Width=320,
            DropDownStyle=ComboBoxStyle.DropDownList, Parent=self
        )
        for name in sorted(line_style_dict.keys()):
            self.cb_line_style.Items.Add(name)
        if self.cb_line_style.Items.Count > 0:
            self.cb_line_style.SelectedIndex = 0
        y += 35

        WinLabel(Text='Text Note Type:', Location=Point(15, y), AutoSize=True, Parent=self)
        y += 20
        self.cb_text_type = WinComboBox(
            Location=Point(15, y), Width=320,
            DropDownStyle=ComboBoxStyle.DropDownList, Parent=self
        )
        for name in sorted(text_type_dict.keys()):
            self.cb_text_type.Items.Add(name)
        if self.cb_text_type.Items.Count > 0:
            self.cb_text_type.SelectedIndex = 0
        y += 40

        btn_ok = WinButton(
            Text='OK', Location=Point(150, y), Width=75,
            DialogResult=DialogResult.OK, Parent=self
        )
        btn_cancel = WinButton(
            Text='Cancel', Location=Point(230, y), Width=75,
            DialogResult=DialogResult.Cancel, Parent=self
        )

        self.AcceptButton = btn_ok
        self.CancelButton = btn_cancel

setup_form = TableSetupForm()
result = setup_form.ShowDialog()

if result != DialogResult.OK:
    script.exit()

fallback_w_inch = setup_form.tb_width.Text
fallback_h_inch = setup_form.tb_height.Text
multiplier_text = setup_form.tb_multiplier.Text
padding_text = setup_form.tb_padding.Text
chosen_line_style_name = setup_form.cb_line_style.SelectedItem
chosen_text_type_name = setup_form.cb_text_type.SelectedItem

if (not fallback_w_inch or not fallback_h_inch or not multiplier_text or not padding_text
        or not chosen_line_style_name or not chosen_text_type_name):
    forms.alert('Missing input. Please fill in all fields.', exitscript=True)

fallback_w_inch = float(fallback_w_inch)
fallback_h_inch = float(fallback_h_inch)
size_multiplier = float(multiplier_text)
padding_percent = float(padding_text)
pad_fraction = padding_percent / 100.0
chosen_line_style = line_style_dict[chosen_line_style_name]
chosen_text_type = text_type_dict[chosen_text_type_name]

min_width_limit = TextNote.GetMinimumAllowedWidth(doc, chosen_text_type.Id)
max_width_limit = TextNote.GetMaximumAllowedWidth(doc, chosen_text_type.Id)
output.print_md(
    '**DEBUG:** Text width limit : {:.4f} ft - {:.4f} ft'.format(
        min_width_limit, max_width_limit
    )
)

# ---------------------------------------------------------------------------
# 3. Pick insertion point (top-left corner of the table)
# ---------------------------------------------------------------------------   
try:
    origin = uidoc.Selection.PickPoint('Pick the top-left corner of the table')
except Exception:
    script.exit()

# ---------------------------------------------------------------------------
# 4. Build geometry: per-column/row sizes (from Excel if available),
#    scaled by size_multiplier and by the current view's Scale so the
#    printed result matches the requested inch sizes on paper.
# ---------------------------------------------------------------------------
overall_factor = size_multiplier * view_scale

if col_widths_pt:
    col_widths_ft = [pt_to_ft(w) * overall_factor for w in col_widths_pt]
else:
    col_widths_ft = [inch_to_ft(fallback_w_inch) * overall_factor] * n_cols

if row_heights_pt:
    row_heights_ft = [pt_to_ft(h) * overall_factor for h in row_heights_pt]
else:
    row_heights_ft = [inch_to_ft(fallback_h_inch) * overall_factor] * n_rows

# add vertical breathing room so text doesn't touch top/bottom grid lines
row_heights_ft = [h * (1.0 + pad_fraction) for h in row_heights_ft]

col_x = [origin.X]
for w in col_widths_ft:
    col_x.append(col_x[-1] + w)

row_y = [origin.Y]
for h in row_heights_ft:
    row_y.append(row_y[-1] - h)  # going downward

# owner matrix: which master cell (row, col) each grid unit belongs to
owner = [[None] * n_cols for _ in range(n_rows)]
for cell in placed_cells:
    for dr in range(cell['rowspan']):
        for dc in range(cell['colspan']):
            owner[cell['row'] + dr][cell['col'] + dc] = (cell['row'], cell['col'])

align_map = {
    'left': HorizontalTextAlignment.Left,
    'center': HorizontalTextAlignment.Center,
    'right': HorizontalTextAlignment.Right,
}
valign_map = {
    'top': VerticalTextAlignment.Top,
    'middle': VerticalTextAlignment.Middle,
    'bottom': VerticalTextAlignment.Bottom,
}

t = Transaction(doc, 'Create Table from Clipboard')
t.Start()

try:
    # horizontal grid lines - skip segments internal to a merged cell
    for i in range(n_rows + 1):
        for j in range(n_cols):
            if i == 0 or i == n_rows:
                draw = True
            else:
                draw = owner[i - 1][j] != owner[i][j]
            if draw:
                line = Line.CreateBound(
                    XYZ(col_x[j], row_y[i], origin.Z),
                    XYZ(col_x[j + 1], row_y[i], origin.Z)
                )
                dl = doc.Create.NewDetailCurve(view, line)
                dl.LineStyle = chosen_line_style

    # vertical grid lines - skip segments internal to a merged cell
    for j in range(n_cols + 1):
        for i in range(n_rows):
            if j == 0 or j == n_cols:
                draw = True
            else:
                draw = owner[i][j - 1] != owner[i][j]
            if draw:
                line = Line.CreateBound(
                    XYZ(col_x[j], row_y[i], origin.Z),
                    XYZ(col_x[j], row_y[i + 1], origin.Z)
                )
                dl = doc.Create.NewDetailCurve(view, line)
                dl.LineStyle = chosen_line_style

    # text notes, respecting merged bounding box + alignment + padding
    for cell in placed_cells:
        content = cell['text']
        if content.strip() == '':
            continue

        r, c = cell['row'], cell['col']
        rs, cs = cell['rowspan'], cell['colspan']
        x0 = col_x[c]
        x1 = col_x[c + cs]
        y0, y1 = row_y[r], row_y[r + rs]  # y0 = top edge, y1 = bottom edge

        cell_width_ft = x1 - x0
        cell_height_ft = y0 - y1
        h_pad = cell_width_ft * pad_fraction
        v_pad = cell_height_ft * pad_fraction

        h_align = align_map.get(cell['align'], HorizontalTextAlignment.Left)
        if h_align == HorizontalTextAlignment.Left:
            anchor_x = x0 + h_pad / 2.0
        elif h_align == HorizontalTextAlignment.Right:
            anchor_x = x1 - h_pad / 2.0
        else:
            anchor_x = (x0 + x1) / 2.0

        v_align = valign_map.get(cell['valign'], VerticalTextAlignment.Middle)
        if v_align == VerticalTextAlignment.Top:
            anchor_y = y0 - v_pad / 2.0
        elif v_align == VerticalTextAlignment.Bottom:
            anchor_y = y1 + v_pad / 2.0
        else:
            anchor_y = (y0 + y1) / 2.0

        text_options = TextNoteOptions(chosen_text_type.Id)
        text_options.HorizontalAlignment = h_align
        text_options.VerticalAlignment = v_align

        # IMPORTANT: TextNote.Width is an annotation-scale property - Revit
        # automatically multiplies it by view.Scale internally (same
        # mechanism as TextNoteType's Text Size) to compute the actual
        # rendered box size. Since cell_width_ft already includes view_scale
        # (needed to match the Detail Line grid), we must divide it back out
        # here so Revit's own internal multiplication lands on the correct
        # true (paper-space) width instead of double-scaling it.
        desired_width_cell = max((cell_width_ft - h_pad) / view_scale, 0.001)
        safe_width = desired_width_cell
        if safe_width < min_width_limit:
            output.print_md(
                '**CẢNH BÁO:** Ô ({}, {}) nội dung "{}" - cột quá hẹp so với '
                'Text Type đang chọn (cần {:.4f} ft, chỉ có {:.4f} ft). '
                'Note sẽ rộng hơn ô thật. Hãy tăng "Size multiplier" hoặc '
                'chọn Text Note Type có font nhỏ hơn.'.format(
                    r, c, content, min_width_limit, safe_width
                )
            )
            safe_width = min_width_limit
        elif safe_width > max_width_limit:
            safe_width = max_width_limit

        TextNote.Create(
            doc, view.Id, XYZ(anchor_x, anchor_y, origin.Z),
            safe_width, content, text_options
        )

    t.Commit()
    output.print_md('**Done.** Table created: {} rows x {} columns.'.format(n_rows, n_cols))

except Exception as ex:
    t.RollBack()
    forms.alert('Error creating table:\n{}'.format(ex))