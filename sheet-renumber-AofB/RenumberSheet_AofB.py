# -*- coding: utf-8 -*-
__title__ = "Sheet Numbering\nA of B"
__doc__ = "Auto-number sheets (Sheet_Pos Index) and total count (Sheet_Amount), filtered by a user-defined parameter, grouped by TT_SHEET_TYPE, then sorted by Sheet Number"

import re
from pyrevit import revit, DB, forms, script

doc = revit.doc
output = script.get_output()

POS_INDEX_PARAM = "Sheet_Pos Index"
AMOUNT_PARAM = "Sheet_Amount"
SHEET_TYPE_PARAM = "TT_SHEET_TYPE"


def natural_sort_key(text):
    """
    Extract the numeric part from strings like 'DM-100', 'DM-1200'
    and return it as an integer for correct numeric sorting.
    Falls back to sorting last if no digits are found.
    """
    match = re.search(r'\d+', text)
    if match:
        return (int(match.group()), text)
    return (float('inf'), text)


# ---------- 0. Ask user which parameter to filter by ----------
filter_param_name = forms.ask_for_string(
    default="TT_SET",
    prompt="Enter the parameter name to filter/group sheets by:",
    title="Sheet Numbering Setup"
)

if not filter_param_name:
    script.exit()

# ---------- 1. Collect all ViewSheets ----------
all_sheets = DB.FilteredElementCollector(doc)\
    .OfClass(DB.ViewSheet)\
    .WhereElementIsNotElementType()\
    .ToElements()

if not all_sheets:
    forms.alert("No sheets found in the model.", exitscript=True)

# ---------- 2. Build distinct values for the chosen filter parameter ----------
series_map = {}  # {filter_value: [sheet1, sheet2, ...]}
sheets_missing_param = 0

for sheet in all_sheets:
    filter_param = sheet.LookupParameter(filter_param_name)

    # Skip sheets that don't have this parameter at all
    if filter_param is None:
        sheets_missing_param += 1
        continue

    filter_value = filter_param.AsString()
    if not filter_value:
        continue

    if filter_value not in series_map:
        series_map[filter_value] = []
    series_map[filter_value].append(sheet)

if not series_map:
    forms.alert(
        "No sheet has a value in '{}' parameter.\n"
        "Check the parameter name is correct (case-sensitive).".format(filter_param_name),
        exitscript=True
    )

if sheets_missing_param > 0:
    print("Note: {} sheet(s) do not have parameter '{}' and were ignored.".format(
        sheets_missing_param, filter_param_name
    ))

# ---------- 3. Let user pick which values to number ----------
available_values = sorted(series_map.keys())

selected_values = forms.SelectFromList.show(
    available_values,
    title="Select '{}' Value(s) to Number".format(filter_param_name),
    multiselect=True
)

if not selected_values:
    script.exit()

# ---------- 4. Number sheets per selected value ----------
# Order within each group: group by TT_SHEET_TYPE first (natural sort,
# e.g. "DM-100" < "DM-200" < "DM-1200"), then sort by Sheet Number within each sub-group.
result_log = []
updated_count = 0
skipped_count = 0

with revit.Transaction("Auto Number Sheets by {}".format(filter_param_name)):
    for filter_value in selected_values:
        sheets_in_group = series_map[filter_value]

        # ---- 4a. Sub-group by TT_SHEET_TYPE ----
        sheet_type_map = {}  # {sheet_type_value: [sheet1, sheet2, ...]}
        sheets_missing_type = []

        for sheet in sheets_in_group:
            type_param = sheet.LookupParameter(SHEET_TYPE_PARAM)
            type_value = type_param.AsString() if type_param else None

            if not type_value:
                # Sheets without a TT_SHEET_TYPE value go into a fallback bucket
                # so they are still numbered, just sorted last.
                sheets_missing_type.append(sheet)
                continue

            if type_value not in sheet_type_map:
                sheet_type_map[type_value] = []
            sheet_type_map[type_value].append(sheet)

        # ---- 4b. Sort TT_SHEET_TYPE groups using natural sort ----
        # (numeric-aware, so "DM-1200" correctly sorts after "DM-200")
        sorted_type_keys = sorted(sheet_type_map.keys(), key=natural_sort_key)

        # ---- 4c. Flatten: sort by Sheet Number within each TT_SHEET_TYPE group ----
        ordered_sheets = []
        for type_key in sorted_type_keys:
            group_sheets = sheet_type_map[type_key]
            group_sorted = sorted(group_sheets, key=lambda s: s.SheetNumber)
            ordered_sheets.extend(group_sorted)

        # Append sheets missing TT_SHEET_TYPE at the end, still sorted by Sheet Number
        if sheets_missing_type:
            ordered_sheets.extend(
                sorted(sheets_missing_type, key=lambda s: s.SheetNumber)
            )

        total_count = len(ordered_sheets)

        # ---- 4d. Assign Pos Index / Amount over the full ordered list ----
        for index, sheet in enumerate(ordered_sheets, start=1):
            pos_param = sheet.LookupParameter(POS_INDEX_PARAM)
            amount_param = sheet.LookupParameter(AMOUNT_PARAM)

            # Skip if either parameter is missing or read-only
            if pos_param is None or amount_param is None:
                skipped_count += 1
                continue
            if pos_param.IsReadOnly or amount_param.IsReadOnly:
                skipped_count += 1
                continue

            # Both are Text parameters, so values must be converted to string
            pos_param.Set(str(index))
            amount_param.Set(str(total_count))

            result_log.append([
                filter_value,
                sheet.SheetNumber,
                index,
                total_count
            ])
            updated_count += 1

# ---------- 5. Output results ----------
output.print_md("## Sheet Numbering Results")
output.print_md("- **Filter parameter:** {}".format(filter_param_name))
output.print_md("- **Values processed:** {}".format(", ".join(selected_values)))
output.print_md("- **Updated:** {}".format(updated_count))
output.print_md("- **Skipped (missing/read-only param):** {}".format(skipped_count))

output.print_table(
    table_data=result_log,
    columns=[filter_param_name, "Sheet Number", "Pos Index", "Amount"]
)

forms.alert(
    "Done!\nUpdated: {}\nSkipped: {}".format(updated_count, skipped_count),
    title="Sheet Numbering"
)