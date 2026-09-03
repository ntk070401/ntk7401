""Create a grid table (detail lines + text notes) in the active view
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
no alignment) if Excel HTML clipboard data is not available. -->
""