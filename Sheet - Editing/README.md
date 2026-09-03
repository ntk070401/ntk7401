# -*- coding: utf-8 -*-
"""Filter Sheets by a text parameter (default 'TT_SET'), then bulk-edit
Sheet Number / Sheet Name.

Workflow:
    1. Collect all ViewSheet in the document.
    2. Ask for the parameter name to filter by (default "TT_SET").
       - If left empty, ALL sheets in the model are used (no filter).
    3. Read distinct values of that parameter and let the user pick one
       or more values to filter by (multi-select).
    4. Show an editable grid (DataGridView) with Sheet Number / Sheet Name.
    5. On "Apply", write changes back inside a Transaction.
"""