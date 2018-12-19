# ExcelCellMarking
We have a big Excel sheet (or rather a set of those) with a list of tasks to be done during the day. For convinience's sake we need to mark who has done what and when (for tracking). In order to simplify it a bit this macro was added to the workbook.

What `GeneralMarking()` does is pretty simple: depending on what 1 letter you input in a cell, following a press of "Enter" it will add appropriate username and timestamp and color the cell with suitable color. Red for failed items, green for completed, orange for "in progress". It will also select the next non-hidden cell in the column. It support English and Russian keyboards both upper and lower cases:
- `C` = completed (green)
- `E` = error (red)
- `L` = late (special case, gray)
- `P` = clear cell
- `S` = in progress (orange)
- `Z` = omitted (special case, gray)

`markascomp()` is a partner function which marks a selected range of cells in a similar manner, which we use to mark sets of tasks, which are usually done in a quck succession (we use a seperate button, that triggers the function for each set).
