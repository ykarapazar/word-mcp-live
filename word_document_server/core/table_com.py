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


def set_cell(table, row, col, text, accept_revisions=False):
    """Set cell text at (row, col). accept_revisions=True clears tracked changes first."""
    _validate_cell(table, row, col)
    cell = table.Cell(row, col)
    if accept_revisions:
        revs = cell.Range.Revisions
        if revs.Count > 0:
            revs.AcceptAll()
    cell.Range.Text = text
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


def set_row(table, row, cells, accept_revisions=False):
    """Set all cell values in an existing row.

    Args:
        table: Word Table COM object.
        row: 1-based row index.
        cells: List of strings (1-based col order). None values skip that cell.
        accept_revisions: Accept tracked changes in cells before writing.

    Returns:
        dict with row, cells_updated list, and total_cols.
    """
    if row < 1 or row > table.Rows.Count:
        raise ValueError(f"Row {row} out of range (1-{table.Rows.Count})")
    updated = []
    for i, val in enumerate(cells):
        c = i + 1
        if c > table.Columns.Count:
            break
        if val is None:
            continue
        cell = table.Cell(row, c)
        if accept_revisions:
            revs = cell.Range.Revisions
            if revs.Count > 0:
                revs.AcceptAll()
        cell.Range.Text = str(val)
        updated.append(c)
    return {"row": row, "cells_updated": updated, "total_cols": table.Columns.Count}


def set_range(table, data, start_row=1, start_col=1, accept_revisions=False):
    """Set a rectangular block of cells from a 2D list.

    Args:
        table: Word Table COM object.
        data: 2D list (list of row-lists). None values skip that cell.
        start_row: 1-based row to start writing at (default 1).
        start_col: 1-based column to start writing at (default 1).
        accept_revisions: Accept tracked changes in cells before writing.

    Returns:
        dict with cells_updated count, start position, and data dimensions.
    """
    updated = []
    for ri, row_data in enumerate(data):
        r = start_row + ri
        if r > table.Rows.Count:
            break
        for ci, val in enumerate(row_data):
            c = start_col + ci
            if c > table.Columns.Count:
                break
            if val is None:
                continue
            cell = table.Cell(r, c)
            if accept_revisions:
                revs = cell.Range.Revisions
                if revs.Count > 0:
                    revs.AcceptAll()
            cell.Range.Text = str(val)
            updated.append([r, c])
    return {
        "cells_updated": len(updated),
        "start": [start_row, start_col],
        "data_rows": len(data),
        "data_cols": max((len(r) for r in data), default=0),
    }


def delete_table(table, scrub_orphans: bool = True):
    """Delete the entire table object from the document.

    Args:
        table: Word.Table COM object to delete.
        scrub_orphans: After deletion, scan a small window around the
            former table location and remove any orphan cell-separator
            bytes (\\x07) that don't belong to a remaining table.
            Word.Table.Delete() occasionally leaves these behind, and
            they corrupt subsequent Find/Replace and add_table calls.
    """
    rows = table.Rows.Count
    cols = table.Columns.Count
    doc = table.Range.Document
    start = table.Range.Start
    table.Delete()

    scrubbed = 0
    if scrub_orphans:
        # Scan a 20-char window around the former table for stranded \x07
        # bytes that are NOT inside any remaining table's range.
        try:
            content_end = doc.Content.End - 1
            scan_start = max(0, start - 1)
            scan_end = min(start + 20, content_end)
            if scan_end > scan_start:
                in_table_ranges = []
                for t in doc.Tables:
                    try:
                        in_table_ranges.append((t.Range.Start, t.Range.End))
                    except Exception:
                        continue
                window = doc.Range(scan_start, scan_end)
                text = window.Text or ""
                # Walk right-to-left so deletions don't shift positions of
                # later matches we still need to find.
                for i in range(len(text) - 1, -1, -1):
                    if text[i] != "\x07":
                        continue
                    pos = scan_start + i
                    if any(s <= pos <= e for (s, e) in in_table_ranges):
                        continue  # belongs to a real surviving table
                    try:
                        doc.Range(pos, pos + 1).Delete()
                        scrubbed += 1
                    except Exception:
                        pass
        except Exception:
            # Best-effort cleanup; never fail the delete itself.
            pass
    return {
        "deleted": True,
        "had_rows": rows,
        "had_cols": cols,
        "scrubbed_orphans": scrubbed,
    }


def _validate_cell(table, row, col):
    """Validate row/col are within table bounds."""
    if row < 1 or row > table.Rows.Count:
        raise ValueError(f"Row {row} out of range (1-{table.Rows.Count})")
    if col < 1 or col > table.Columns.Count:
        raise ValueError(f"Col {col} out of range (1-{table.Columns.Count})")
