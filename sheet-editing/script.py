import clr

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from System.Windows.Forms import (
    Application, Form, DataGridView, DataGridViewTextBoxColumn,
    DataGridViewCellStyle, DataGridViewContentAlignment, DataGridViewTriState,
    DataGridViewColumnHeadersHeightSizeMode, DataGridViewSelectionMode,
    Button, Label, Panel, Padding, DockStyle, DataGridViewAutoSizeColumnsMode,
    MessageBox, MessageBoxButtons, MessageBoxIcon, DialogResult,
    FormStartPosition, FlatStyle, AnchorStyles, BorderStyle, AutoScaleMode
)
from System.Drawing import Size, Point, Font, FontStyle, Color, ContentAlignment

import re
from pyrevit import revit, DB, forms, script

doc = revit.doc
output = script.get_output()

DEFAULT_PARAM_NAME = "TT_SET"

PARAM_NAME = forms.ask_for_string(
    default=DEFAULT_PARAM_NAME,
    prompt="Enter the parameter name to filter Sheets by (leave empty for ALL sheets):",
    title="Filter Parameter"
)

# ask_for_string returns None only when the user cancels the dialog.
# An empty string means the user deliberately cleared the field -> no filter.
if PARAM_NAME is None:
    script.exit()

PARAM_NAME = PARAM_NAME.strip()
USE_ALL_SHEETS = (PARAM_NAME == "")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def get_param_string(element, param_name):
    """Return the string value of a parameter, or '' if not found/empty."""
    p = element.LookupParameter(param_name)
    if p is None or not p.HasValue:
        return ""
    # AsString works for TEXT parameters; fall back to AsValueString for others
    val = p.AsString()
    if val is None:
        val = p.AsValueString()
    return val or ""


def natural_sort_key(text):
    """Split text into digit/non-digit chunks so numbers sort numerically
    (e.g. 'DM-9.00' comes before 'DM-10.00' instead of after, as plain
    string comparison would give)."""
    return [
        int(chunk) if chunk.isdigit() else chunk.lower()
        for chunk in re.split(r"(\d+)", text)
    ]


def collect_sheets():
    return list(
        DB.FilteredElementCollector(doc)
        .OfClass(DB.ViewSheet)
        .WhereElementIsNotElementType()
        .ToElements()
    )


# ---------------------------------------------------------------------------
# Step 1: collect sheets, optionally filter by parameter value(s)
# ---------------------------------------------------------------------------
all_sheets = collect_sheets()

if not all_sheets:
    forms.alert("No sheets were found in this model.", exitscript=True)

if USE_ALL_SHEETS:
    # No parameter name given -> use every sheet in the model, no value filter.
    filtered_sheets = all_sheets
    selected_values = None
    filter_desc = "All sheets (no filter)"
else:
    param_values = sorted(
        set(get_param_string(s, PARAM_NAME) for s in all_sheets if get_param_string(s, PARAM_NAME))
    )

    if not param_values:
        forms.alert(
            "No sheet was found with a value for parameter '{}'.".format(PARAM_NAME),
            exitscript=True
        )

    selected_values = forms.SelectFromList.show(
        param_values,
        title="Select one or more {} values to filter sheets".format(PARAM_NAME),
        button_name="Select",
        multiselect=True
    )

    if not selected_values:
        script.exit()

    filtered_sheets = [
        s for s in all_sheets if get_param_string(s, PARAM_NAME) in selected_values
    ]

    if not filtered_sheets:
        forms.alert("No sheet matches the selected value(s).", exitscript=True)

    filter_desc = "{} = {}".format(PARAM_NAME, ", ".join(selected_values))

# sort by Sheet Number (natural sort: numeric-aware, not plain string sort)
filtered_sheets.sort(key=lambda s: natural_sort_key(s.SheetNumber))


# ---------------------------------------------------------------------------
# Step 2: build the editable grid (WinForms)
# ---------------------------------------------------------------------------
class SheetEditForm(Form):

    # Color palette
    COLOR_HEADER_BG = Color.FromArgb(45, 62, 80)      # dark slate blue
    COLOR_HEADER_FG = Color.White
    COLOR_ALT_ROW = Color.FromArgb(245, 247, 250)     # light gray-blue
    COLOR_SELECTION = Color.FromArgb(52, 152, 219)    # bright blue
    COLOR_APPLY_BTN = Color.FromArgb(39, 174, 96)     # green
    COLOR_CANCEL_BTN = Color.FromArgb(149, 165, 166)  # gray
    COLOR_TITLE_BG = Color.FromArgb(45, 62, 80)

    def __init__(self, sheets):
        Form.__init__(self)
        self.sheets = sheets
        self.result_rows = None  # filled on Apply

        self.Text = "Edit Sheets"
        self.AutoScaleMode = AutoScaleMode.Dpi
        self.Width = 2600
        self.Height = 1800
        self.MinimumSize = Size(700, 450)
        self.StartPosition = FormStartPosition.CenterScreen
        self.Font = Font("Segoe UI", 9)
        self.BackColor = Color.White

        # ---- Title bar (colored header strip with context info) ----
        title_panel = Label()
        title_panel.Dock = DockStyle.Top
        title_panel.Height = 56
        title_panel.BackColor = self.COLOR_TITLE_BG
        title_panel.ForeColor = Color.White
        title_panel.Font = Font("Segoe UI", 12, FontStyle.Bold)
        title_panel.Text = "  Edit Sheets"
        title_panel.TextAlign = ContentAlignment.MiddleLeft
        self.Controls.Add(title_panel)

        subtitle = Label()
        subtitle.Dock = DockStyle.Top
        subtitle.Height = 28
        subtitle.BackColor = Color.FromArgb(236, 240, 241)
        subtitle.ForeColor = Color.FromArgb(80, 80, 80)
        subtitle.Font = Font("Segoe UI", 9, FontStyle.Italic)
        subtitle.Text = "  {}   |   {} sheet(s) found".format(
            filter_desc, len(sheets)
        )
        subtitle.TextAlign = ContentAlignment.MiddleLeft
        self.Controls.Add(subtitle)

        # ---- Grid ----
        self.grid = DataGridView()
        self.grid.Dock = DockStyle.Fill
        self.grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        self.grid.AllowUserToAddRows = False
        self.grid.AllowUserToDeleteRows = False
        self.grid.AllowUserToResizeRows = False
        self.grid.RowHeadersVisible = False
        self.grid.BorderStyle = getattr(BorderStyle, "None")
        self.grid.BackgroundColor = Color.White
        self.grid.GridColor = Color.FromArgb(224, 224, 224)
        self.grid.SelectionMode = DataGridViewSelectionMode.CellSelect
        self.grid.MultiSelect = False
        self.grid.RowTemplate.Height = 34
        self.grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        self.grid.ColumnHeadersHeight = 38
        self.grid.EnableHeadersVisualStyles = False

        header_style = DataGridViewCellStyle()
        header_style.BackColor = self.COLOR_HEADER_BG
        header_style.ForeColor = self.COLOR_HEADER_FG
        header_style.Font = Font("Segoe UI", 10, FontStyle.Bold)
        header_style.Alignment = DataGridViewContentAlignment.MiddleLeft
        header_style.WrapMode = getattr(DataGridViewTriState, "False")
        header_style.Padding = Padding(6, 0, 0, 0)
        self.grid.ColumnHeadersDefaultCellStyle = header_style

        default_style = DataGridViewCellStyle()
        default_style.Font = Font("Segoe UI", 10)
        default_style.SelectionBackColor = self.COLOR_SELECTION
        default_style.SelectionForeColor = Color.White
        default_style.WrapMode = getattr(DataGridViewTriState, "False")
        default_style.Padding = Padding(4, 2, 4, 2)
        self.grid.DefaultCellStyle = default_style

        # row height must be set AFTER the font/padding are applied, and
        # applied to existing rows too (not just the template for new rows)
        self.grid.RowTemplate.Height = 34

        self.grid.AlternatingRowsDefaultCellStyle.BackColor = self.COLOR_ALT_ROW

        col_id = DataGridViewTextBoxColumn()
        col_id.HeaderText = "ElementId"
        col_id.Name = "ElementId"
        col_id.ReadOnly = True
        col_id.Visible = False  # hidden helper column

        col_number = DataGridViewTextBoxColumn()
        col_number.HeaderText = "Sheet Number"
        col_number.Name = "SheetNumber"
        col_number.FillWeight = 30

        col_name = DataGridViewTextBoxColumn()
        col_name.HeaderText = "Sheet Name"
        col_name.Name = "SheetName"
        col_name.FillWeight = 70

        self.grid.Columns.Add(col_id)
        self.grid.Columns.Add(col_number)
        self.grid.Columns.Add(col_name)

        for sh in sheets:
            row_index = self.grid.Rows.Add(sh.Id.IntegerValue, sh.SheetNumber, sh.Name)
            self.grid.Rows[row_index].Height = 40

        self.Controls.Add(self.grid)
        self.grid.BringToFront()

        # ---- Bottom button bar ----
        button_panel = Panel()
        button_panel.Dock = DockStyle.Bottom
        button_panel.Height = 56
        button_panel.BackColor = Color.FromArgb(250, 250, 250)
        self.Controls.Add(button_panel)

        # NOTE: button_panel has not been through layout yet at this point,
        # so its .Width is unreliable. The panel is Dock=Bottom and will
        # stretch to the form's client width, so we use self.ClientSize.Width
        # (already known, since self.Width was set earlier) to place buttons.
        panel_width = self.ClientSize.Width

        btn_cancel = Button()
        btn_cancel.Text = "Cancel"
        btn_cancel.Width = 130
        btn_cancel.Height = 34
        btn_cancel.Anchor = AnchorStyles.Top | AnchorStyles.Right
        btn_cancel.Location = Point(panel_width - 130, 11)
        btn_cancel.FlatStyle = FlatStyle.Flat
        btn_cancel.FlatAppearance.BorderColor = self.COLOR_CANCEL_BTN
        btn_cancel.ForeColor = Color.FromArgb(80, 80, 80)
        btn_cancel.BackColor = Color.White
        btn_cancel.Font = Font("Segoe UI", 9)
        btn_cancel.Click += self.on_cancel

        btn_apply = Button()
        btn_apply.Text = "Apply"
        btn_apply.Width = 130
        btn_apply.Height = 34
        btn_apply.Anchor = AnchorStyles.Top | AnchorStyles.Right
        btn_apply.Location = Point(panel_width - 250, 11)
        btn_apply.FlatStyle = FlatStyle.Flat
        btn_apply.FlatAppearance.BorderSize = 0
        btn_apply.ForeColor = Color.White
        btn_apply.BackColor = self.COLOR_APPLY_BTN
        btn_apply.Font = Font("Segoe UI", 9, FontStyle.Bold)
        btn_apply.Click += self.on_apply

        button_panel.Controls.Add(btn_apply)
        button_panel.Controls.Add(btn_cancel)

    def on_apply(self, sender, args):
        rows = []
        for row in self.grid.Rows:
            if row.IsNewRow:
                continue
            eid = int(row.Cells["ElementId"].Value)
            number = (row.Cells["SheetNumber"].Value or "").strip()
            name = (row.Cells["SheetName"].Value or "").strip()
            if not number or not name:
                MessageBox.Show(
                    "Sheet Number and Sheet Name cannot be empty.",
                    "Validation Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                )
                return
            rows.append((eid, number, name))

        # check duplicate sheet numbers within the edited set itself
        numbers = [r[1] for r in rows]
        if len(numbers) != len(set(numbers)):
            MessageBox.Show(
                "Some Sheet Numbers are duplicated in the table. Please review and try again.",
                "Validation Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )
            return

        self.result_rows = rows
        self.DialogResult = DialogResult.OK
        self.Close()

    def on_cancel(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()


edit_form = SheetEditForm(filtered_sheets)
dialog_result = edit_form.ShowDialog()

if dialog_result != DialogResult.OK or edit_form.result_rows is None:
    script.exit()

edited_rows = edit_form.result_rows


# ---------------------------------------------------------------------------
# Step 3: apply changes inside a Transaction
# ---------------------------------------------------------------------------
# Build lookup: ElementId -> ViewSheet, and original values
sheet_by_id = {s.Id.IntegerValue: s for s in filtered_sheets}

changes = []  # (sheet, new_number, new_name, old_number)
for eid, new_number, new_name in edited_rows:
    sheet = sheet_by_id.get(eid)
    if sheet is None:
        continue
    old_number = sheet.SheetNumber
    old_name = sheet.Name
    if new_number != old_number or new_name != old_name:
        changes.append((sheet, new_number, new_name, old_number, old_name))

if not changes:
    forms.alert("No changes to apply.", exitscript=True)

t = DB.Transaction(doc, "Edit Sheets")
t.Start()

try:
    # Pass 1: any sheet whose SheetNumber is changing gets a temporary unique
    # number first, to avoid "duplicate sheet number" errors when numbers
    # are being swapped/rotated between sheets.
    temp_map = {}
    for sheet, new_number, new_name, old_number, old_name in changes:
        if new_number != old_number:
            temp_number = "TMP_{}".format(sheet.Id.IntegerValue)
            sheet.SheetNumber = temp_number
            temp_map[sheet.Id.IntegerValue] = temp_number

    # Pass 2: set final Sheet Number and Sheet Name
    for sheet, new_number, new_name, old_number, old_name in changes:
        if new_name != old_name:
            sheet.Name = new_name
        if new_number != old_number:
            sheet.SheetNumber = new_number

    t.Commit()
except Exception as ex:
    t.RollBack()
    forms.alert("An error occurred, changes were rolled back:\n{}".format(ex), exitscript=True)

output.print_md("### Updated {} sheet(s):".format(len(changes)))
for sheet, new_number, new_name, old_number, old_name in changes:
    output.print_md("- `{}` -> `{}`  |  `{}` -> `{}`".format(
        old_number, new_number, old_name, new_name
    ))
