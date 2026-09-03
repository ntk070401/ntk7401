# -*- coding: utf-8 -*-
"""Create a grid table (detail lines + text notes) in the active view
from data copied from Excel (Ctrl+C on a selected range).

Reads Excel's HTML clipboard format to preserve:
  - merged cells (colspan/rowspan)
  - text alignment (horizontal + vertical), read from inline style OR
    from the cell's CSS class (Excel usually stores alignment there)
  - actual column widths / row heights (from Excel's <col>/<tr> style)

Table geometry is scaled by the current view's Scale so the printed
size on paper matches what you type in inches, regardless of the
view's drawing scale at creation time. (Note: this is a one-time
scale-correct creation; if you change the view scale afterwards,
the table will NOT auto-resize - you would need to recreate it.)

TextNote width is constrained to each cell's usable width. Long text
may wrap inside the cell when it cannot fit on one line.

Falls back to plain tab-separated text (uniform cell size, no merge,
no alignment) if Excel HTML clipboard data is not available.

Input values are validated before geometry creation.
"""
__title__ = "Table\nExcel to Revit"

import re
import clr

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('System')

from System import IO
from System.Net import WebUtility
from System.Drawing import Point, Size, Font, FontStyle, GraphicsUnit
from System.Windows.Forms import (
    Clipboard, Form, Label as WinLabel, TextBox as WinTextBox,
    ComboBox as WinComboBox, Button as WinButton, DialogResult,
    FormStartPosition, ComboBoxStyle, FormBorderStyle, AutoScaleMode,
    TableLayoutPanel, FlowLayoutPanel, DockStyle, FlowDirection,
    SizeType, ColumnStyle, RowStyle, Padding, AutoSizeMode, AnchorStyles
)

from Autodesk.Revit.DB import (
    XYZ, Line, Transaction, TextNote, TextNoteOptions, TextNoteType,
    FilteredElementCollector, HorizontalTextAlignment, VerticalTextAlignment,
    ViewType, BuiltInCategory, GraphicsStyleType, BuiltInParameter
)

from pyrevit import forms, revit, script

doc = revit.doc
uidoc = revit.uidoc
view = doc.ActiveView
output = script.get_output()

INCH_TO_FT = 1.0 / 12.0
PT_TO_INCH = 1.0 / 72.0


def parse_user_float(raw_value, field_name, minimum=None, maximum=None, inclusive_min=True,
                    inclusive_max=True):
    """Parse and validate a numeric value entered in the setup dialog."""
    value_text = str(raw_value).strip().replace(',', '.')
    try:
        value = float(value_text)
    except (TypeError, ValueError):
        raise ValueError('{} must be a valid number.'.format(field_name))

    if minimum is not None:
        invalid_min = value < minimum or (not inclusive_min and value == minimum)
        if invalid_min:
            comparator = 'greater than' if not inclusive_min else 'at least'
            raise ValueError('{} must be {} {}.'.format(field_name, comparator, minimum))
    if maximum is not None:
        invalid_max = value > maximum or (not inclusive_max and value == maximum)
        if invalid_max:
            comparator = 'less than' if not inclusive_max else 'at most'
            raise ValueError('{} must be {} {}.'.format(field_name, comparator, maximum))
    return value


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
    ViewType.Section, ViewType.DraftingView, ViewType.EngineeringPlan,
    ViewType.Legend
]
if view.ViewType not in allowed_view_types:
    forms.alert(
        'Current view type is not supported for Detail Lines / Text Notes.\n'
        'Please switch to a Plan, Section, Elevation, Legend or Drafting view.',
        exitscript=True
    )

view_scale = view.Scale if view.Scale else 1
output.print_md('**View scale:** 1:{}'.format(view_scale))

# ---------------------------------------------------------------------------
# 1. Read table data from clipboard
#    Preferred: "HTML Format" (Excel writes real table markup with
#    colspan/rowspan/style/col-width/row-height/CSS classes so we can
#    rebuild merges, alignment, and true column/row sizes).
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


def parse_col_widths_pt(html_fragment, n_cols):
    """Parse <col style='width:XXpt'> tags (honoring 'span') into a list
    of per-column widths in points. Returns None if no <col> tags found."""
    col_tags = re.findall(r'<col([^>]*)>', html_fragment, re.I)
    if not col_tags:
        return None

    widths = []
    for attrs in col_tags:
        span_m = re.search(r"span\s*=\s*['\"]?(\d+)['\"]?", attrs, re.I)
        span = int(span_m.group(1)) if span_m else 1
        width_m = re.search(r"width\s*:\s*([\d.]+)\s*pt", attrs, re.I)
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

    Excel's clipboard StartFragment/EndFragment markers often sit *inside*
    the <table> tag, so we search for <tr> rows directly rather than
    requiring an enclosing <table> wrapper.

    Alignment is read from inline style first, falling back to the
    cell's CSS class (css_classes), since Excel usually stores alignment
    at the class level rather than inline.

    Returns (placed_cells, n_rows, n_cols, col_widths_pt, row_heights_pt).
    col_widths_pt / row_heights_pt are None if not found in the HTML.
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

            colspan_m = re.search(r"colspan\s*=\s*['\"]?(\d+)['\"]?", attrs, re.I)
            rowspan_m = re.search(r"rowspan\s*=\s*['\"]?(\d+)['\"]?", attrs, re.I)
            colspan = int(colspan_m.group(1)) if colspan_m else 1
            rowspan = int(rowspan_m.group(1)) if rowspan_m else 1

            # class attribute may or may not be quoted: class=xl68 OR class="xl68"
            class_m = re.search(r'class\s*=\s*["\']?([\w\-]+(?:\s+[\w\-]+)*)["\']?', attrs, re.I)
            cell_classes = class_m.group(1).split() if class_m else []

            align_m = re.search(r'text-align\s*:\s*(\w+)', attrs, re.I)
            align = align_m.group(1).lower() if align_m else None
            if align == 'general':
                align = None

            valign_m = re.search(r'vertical-align\s*:\s*(\w+)', attrs, re.I)
            if not valign_m:
                valign_m = re.search(r"valign\s*=\s*['\"]?(\w+)['\"]?", attrs, re.I)
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
parsed = parse_html_table(html_fragment, css_classes) if html_fragment else None
using_html = parsed is not None

if using_html:
    placed_cells, n_rows, n_cols, col_widths_pt, row_heights_pt = parsed
else:
    placed_cells, n_rows, n_cols = parse_plain_text_table()
    col_widths_pt = None
    row_heights_pt = None

merged_count = sum(1 for c in placed_cells if c['rowspan'] > 1 or c['colspan'] > 1)
align_count = sum(1 for c in placed_cells if c.get('align'))
valign_count = sum(1 for c in placed_cells if c.get('valign'))
output.print_md(
    '**Detected table:** {} rows x {} columns | merged cells: {} | '
    'Excel column widths: {} | Excel row heights: {} | '
    'align detected: {} o | valign detected: {} o'.format(
        n_rows, n_cols, merged_count,
        'Yes' if col_widths_pt else 'No',
        'Yes' if row_heights_pt else 'No',
        align_count, valign_count
    )
)

# ---------------------------------------------------------------------------
# 2. Combined input dialog
#    - Fallback cell size (only used when Excel width/height not detected)
#    - Overall size multiplier (applies in both cases)
#    - Cell padding
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

if not line_style_dict:
    forms.alert(
        'No projection line styles are available in this document.\n'
        'Create or load a line style, then run the tool again.',
        exitscript=True
    )

if not text_type_dict:
    forms.alert(
        'No Text Note Types are available in this document.\n'
        'Create or load a Text Note Type, then run the tool again.',
        exitscript=True
    )

class TableSetupForm(Form):
    """DPI-safe setup dialog.

    The layout uses TableLayoutPanel instead of fixed y-coordinates. WinForms
    performs one DPI scaling pass for the form; no manual coordinate or font
    multiplication is used, which prevents double-scaling on 125%-250% displays.
    """

    def __init__(self):
        # Let WinForms scale the entire form consistently for the active display.
        # Do not multiply font sizes, control positions, or dimensions manually.
        self.AutoScaleMode = AutoScaleMode.Dpi
        self.Font = Font('Segoe UI', 9.0, FontStyle.Regular, GraphicsUnit.Point)

        self.Text = 'Table Setup'
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.Sizable
        self.AutoSize = False
        self.ClientSize = Size(720, 460)
        self.MinimumSize = Size(560, 330)
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.ShowInTaskbar = False
        self.Padding = Padding(12, 12, 12, 12)
        # Keep AutoSize disabled so the user can resize the form manually.
        # The content panel is docked to Fill below.

        layout = TableLayoutPanel()
        layout.AutoSize = False
        layout.AutoSizeMode = AutoSizeMode.GrowAndShrink
        layout.Dock = DockStyle.Fill
        layout.ColumnCount = 2
        layout.RowCount = 0
        layout.Padding = Padding(0, 0, 0, 0)
        layout.ColumnStyles.Add(ColumnStyle(SizeType.AutoSize))
        layout.ColumnStyles.Add(ColumnStyle(SizeType.Percent, 100.0))
        self.Controls.Add(layout)
        self.layout = layout

        self._add_labeled_row(
            'Fallback column width (in):',
            self._make_textbox('1.0')
        )
        self.tb_width = self._last_control

        self._add_labeled_row(
            'Fallback row height (in):',
            self._make_textbox('0.3')
        )
        self.tb_height = self._last_control

        self._add_labeled_row(
            'Size multiplier (x):',
            self._make_textbox('1.0')
        )
        self.tb_multiplier = self._last_control

        self._add_labeled_row(
            'Cell padding (%):',
            self._make_textbox('10')
        )
        self.tb_padding = self._last_control

        line_style_combo = self._make_combo_box()
        for name in sorted(line_style_dict.keys()):
            line_style_combo.Items.Add(name)
        if line_style_combo.Items.Count > 0:
            line_style_combo.SelectedIndex = 0
        self._add_labeled_row('Line style (grid):', line_style_combo)
        self.cb_line_style = line_style_combo

        text_type_combo = self._make_combo_box()
        for name in sorted(text_type_dict.keys()):
            text_type_combo.Items.Add(name)
        if text_type_combo.Items.Count > 0:
            text_type_combo.SelectedIndex = 0
        self._add_labeled_row('Text Note Type:', text_type_combo)
        self.cb_text_type = text_type_combo

        # Buttons occupy a final row and stay aligned to the right.
        button_panel = FlowLayoutPanel()
        button_panel.AutoSize = True
        button_panel.AutoSizeMode = AutoSizeMode.GrowAndShrink
        button_panel.Dock = DockStyle.Fill
        button_panel.FlowDirection = FlowDirection.RightToLeft
        button_panel.WrapContents = False
        button_panel.Margin = Padding(0, 10, 0, 0)

        btn_cancel = WinButton(Text='Cancel')
        btn_cancel.AutoSize = True
        btn_cancel.AutoSizeMode = AutoSizeMode.GrowAndShrink
        btn_cancel.MinimumSize = Size(80, 28)
        btn_cancel.DialogResult = DialogResult.Cancel
        btn_cancel.Margin = Padding(6, 0, 0, 0)

        btn_ok = WinButton(Text='OK')
        btn_ok.AutoSize = True
        btn_ok.AutoSizeMode = AutoSizeMode.GrowAndShrink
        btn_ok.MinimumSize = Size(80, 28)
        btn_ok.DialogResult = DialogResult.OK
        btn_ok.Margin = Padding(6, 0, 0, 0)

        button_panel.Controls.Add(btn_cancel)
        button_panel.Controls.Add(btn_ok)

        row = layout.RowCount
        layout.RowCount += 1
        layout.RowStyles.Add(RowStyle(SizeType.AutoSize))
        layout.Controls.Add(button_panel, 0, row)
        layout.SetColumnSpan(button_panel, 2)

        self.AcceptButton = btn_ok
        self.CancelButton = btn_cancel

    def _make_textbox(self, default_text):
        control = WinTextBox(Text=default_text)
        control.Dock = DockStyle.Fill
        control.MinimumSize = Size(140, 0)
        control.Margin = Padding(0, 4, 0, 4)
        return control

    def _make_combo_box(self):
        control = WinComboBox()
        control.Dock = DockStyle.Fill
        control.MinimumSize = Size(320, 0)
        control.DropDownWidth = 500
        control.DropDownStyle = ComboBoxStyle.DropDownList
        control.Margin = Padding(0, 4, 0, 4)
        return control

    def _add_labeled_row(self, label_text, control):
        row = self.layout.RowCount
        self.layout.RowCount += 1
        self.layout.RowStyles.Add(RowStyle(SizeType.AutoSize))

        label = WinLabel(Text=label_text)
        label.AutoSize = True
        label.Anchor = AnchorStyles.Left
        label.Margin = Padding(0, 6, 12, 6)

        self.layout.Controls.Add(label, 0, row)
        self.layout.Controls.Add(control, 1, row)
        self._last_control = control


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

try:
    fallback_w_inch = parse_user_float(
        fallback_w_inch, 'Fallback column width', minimum=0.0, inclusive_min=False
    )
    fallback_h_inch = parse_user_float(
        fallback_h_inch, 'Fallback row height', minimum=0.0, inclusive_min=False
    )
    size_multiplier = parse_user_float(
        multiplier_text, 'Size multiplier', minimum=0.0, inclusive_min=False
    )
    padding_percent = parse_user_float(
        padding_text, 'Cell padding', minimum=0.0, maximum=99.0
    )
except ValueError as input_error:
    forms.alert('Invalid setup value:\n{}'.format(input_error), exitscript=True)

pad_fraction = padding_percent / 100.0
chosen_line_style = line_style_dict[chosen_line_style_name]
chosen_text_type = text_type_dict[chosen_text_type_name]

# TextNote.Create requires a valid width. Keep the note width within
# Revit's allowed limits while matching the usable width of the cell.
min_width_limit = TextNote.GetMinimumAllowedWidth(doc, chosen_text_type.Id)
max_width_limit = TextNote.GetMaximumAllowedWidth(doc, chosen_text_type.Id)

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
    # horizontal grid lines - merge consecutive drawable segments into
    # a single continuous Detail Line instead of one line per unit cell
    for i in range(n_rows + 1):
        j = 0
        while j < n_cols:
            if i == 0 or i == n_rows:
                draw = True
            else:
                draw = owner[i - 1][j] != owner[i][j]

            if not draw:
                j += 1
                continue

            run_start = j
            while j < n_cols:
                if i == 0 or i == n_rows:
                    still_draw = True
                else:
                    still_draw = owner[i - 1][j] != owner[i][j]
                if not still_draw:
                    break
                j += 1
            run_end = j  # exclusive

            line = Line.CreateBound(
                XYZ(col_x[run_start], row_y[i], origin.Z),
                XYZ(col_x[run_end], row_y[i], origin.Z)
            )
            dl = doc.Create.NewDetailCurve(view, line)
            dl.LineStyle = chosen_line_style

    # vertical grid lines - merge consecutive drawable segments into
    # a single continuous Detail Line instead of one line per unit cell
    for j in range(n_cols + 1):
        i = 0
        while i < n_rows:
            if j == 0 or j == n_cols:
                draw = True
            else:
                draw = owner[i][j - 1] != owner[i][j]

            if not draw:
                i += 1
                continue

            run_start = i
            while i < n_rows:
                if j == 0 or j == n_cols:
                    still_draw = True
                else:
                    still_draw = owner[i][j - 1] != owner[i][j]
                if not still_draw:
                    break
                i += 1
            run_end = i  # exclusive

            line = Line.CreateBound(
                XYZ(col_x[j], row_y[run_start], origin.Z),
                XYZ(col_x[j], row_y[run_end], origin.Z)
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

        # TextNote width is stored in annotation/paper units. The grid
        # already includes view.Scale, so divide by the view scale here.
        desired_width = max((cell_width_ft - h_pad) / view_scale, 0.001)
        safe_width = max(min_width_limit, min(desired_width, max_width_limit))

        TextNote.Create(
            doc, view.Id, XYZ(anchor_x, anchor_y, origin.Z),
            safe_width, content, text_options
        )

    t.Commit()
    output.print_md('**Done.** Table created: {} rows x {} columns.'.format(n_rows, n_cols))

except Exception as ex:
    t.RollBack()
    forms.alert('Table creation failed:\n{}'.format(ex))
