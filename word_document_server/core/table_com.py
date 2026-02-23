"""Table manipulation helpers for Word COM automation.

Pure synchronous functions that operate on Word Table COM objects.
Used by live_tools.word_live_modify_table.
All row/col indices are 1-based (Word COM standard).
"""


def get_info(table):
    """Return table structure: dimensions and cell contents.

    Args:
        table: Word Table COM object.

    Returns:
        dict with rows, cols, and data (2D list of cell texts).
    """
    rows = table.Rows.Count
    cols = table.Columns.Count
    data = []
    for r in range(1, rows + 1):
        row_data = []
        for c in range(1, cols + 1):
            try:
                text = table.Cell(r, c).Range.Text
                # Word cell text ends with \r\x07 (paragraph mark + cell mark)
                text = text.rstrip("\r\x07")
                row_data.append(text)
            except Exception:
                row_data.append(None)  # merged/missing cell
        data.append(row_data)
    return {"rows": rows, "cols": cols, "data": data}


def set_cell(table, row, col, text):
    """Set cell text at (row, col)."""
    _validate_cell(table, row, col)
    table.Cell(row, col).Range.Text = text
    return {"row": row, "col": col, "text": text}


def add_column(table, before_col=None, header=None, cells=None):
    """Add a column before before_col (1-based) or at end if None.

    Args:
        table: Word Table COM object.
        before_col: Insert before this column index. None = append at end.
        header: Optional text for the first row of the new column.
        cells: Optional list of cell values (starting from row 2 if header set, row 1 otherwise).

    Returns:
        dict with new_col index and total_cols.
    """
    total = table.Columns.Count
    if before_col is not None:
        if before_col < 1 or before_col > total + 1:
            raise ValueError(f"before_col {before_col} out of range (1-{total + 1})")
        if before_col <= total:
            table.Columns.Add(table.Columns(before_col))
            new_col = before_col
        else:
            table.Columns.Add()
            new_col = table.Columns.Count
    else:
        table.Columns.Add()
        new_col = table.Columns.Count

    if header is not None:
        table.Cell(1, new_col).Range.Text = str(header)
    if cells:
        start_row = 2 if header is not None else 1
        for i, val in enumerate(cells):
            r = start_row + i
            if r > table.Rows.Count:
                break
            table.Cell(r, new_col).Range.Text = str(val)

    return {"new_col": new_col, "total_cols": table.Columns.Count}


def delete_column(table, col):
    """Delete column at index col (1-based)."""
    if col < 1 or col > table.Columns.Count:
        raise ValueError(f"Col {col} out of range (1-{table.Columns.Count})")
    table.Columns(col).Delete()
    return {"deleted_col": col, "remaining_cols": table.Columns.Count}


def add_row(table, before_row=None, cells=None):
    """Add a row before before_row (1-based) or at end if None.

    Args:
        table: Word Table COM object.
        before_row: Insert before this row index. None = append at end.
        cells: Optional list of cell values for the new row.

    Returns:
        dict with new_row index and total_rows.
    """
    total = table.Rows.Count
    if before_row is not None:
        if before_row < 1 or before_row > total + 1:
            raise ValueError(f"before_row {before_row} out of range (1-{total + 1})")
        if before_row <= total:
            table.Rows.Add(table.Rows(before_row))
            new_row = before_row
        else:
            table.Rows.Add()
            new_row = table.Rows.Count
    else:
        table.Rows.Add()
        new_row = table.Rows.Count

    if cells:
        for i, val in enumerate(cells):
            c = i + 1
            if c > table.Columns.Count:
                break
            table.Cell(new_row, c).Range.Text = str(val)

    return {"new_row": new_row, "total_rows": table.Rows.Count}


def delete_row(table, row):
    """Delete row at index row (1-based)."""
    if row < 1 or row > table.Rows.Count:
        raise ValueError(f"Row {row} out of range (1-{table.Rows.Count})")
    table.Rows(row).Delete()
    return {"deleted_row": row, "remaining_rows": table.Rows.Count}


def merge_cells(table, start_row, start_col, end_row, end_col):
    """Merge cells in a rectangular area. All indices 1-based."""
    _validate_cell(table, start_row, start_col)
    _validate_cell(table, end_row, end_col)
    table.Cell(start_row, start_col).Merge(table.Cell(end_row, end_col))
    return {"merged": f"({start_row},{start_col})-({end_row},{end_col})"}


def autofit(table, mode="content"):
    """AutoFit table. mode: 'content', 'window', or 'fixed'."""
    modes = {"content": 1, "window": 2, "fixed": 0}
    val = modes.get(mode.lower())
    if val is None:
        raise ValueError(f"Unknown autofit mode '{mode}'. Use: content, window, fixed")
    table.AutoFitBehavior(val)
    return {"autofit": mode}


def _validate_cell(table, row, col):
    """Validate row/col are within table bounds."""
    if row < 1 or row > table.Rows.Count:
        raise ValueError(f"Row {row} out of range (1-{table.Rows.Count})")
    if col < 1 or col > table.Columns.Count:
        raise ValueError(f"Col {col} out of range (1-{table.Columns.Count})")
