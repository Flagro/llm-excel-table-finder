"""I/O utilities for Excel table finder."""

import csv
from pathlib import Path
from typing import Optional

import click

from src.agent import TablesOutput, TablesWithHeadersOutput
from src.excel_tools import (
    CellRange,
    ExcelReaderBase,
    OpenpyxlReader,
    PyxlsbReader,
    XlrdReader,
)


def get_excel_reader(file_path: str) -> ExcelReaderBase:
    """
    Get the appropriate Excel reader based on file extension.

    Args:
        file_path: Path to the Excel file

    Returns:
        ExcelReaderBase instance (OpenpyxlReader for .xlsx, XlrdReader for .xls)

    Raises:
        ValueError: If file extension is not supported
        ImportError: If required package is not installed
    """
    file_ext = Path(file_path).suffix.lower()

    if file_ext in [".xlsx", ".xlsm"]:
        return OpenpyxlReader(file_path)
    elif file_ext == ".xls":
        return XlrdReader(file_path)
    elif file_ext == ".xlsb":
        return PyxlsbReader(file_path)
    else:
        raise ValueError(
            f"Unsupported file extension: {file_ext}. Supported: .xlsx, .xlsm, .xls, .xlsb"
        )


def export_to_csv(
    tables: TablesOutput | TablesWithHeadersOutput,
    excel_reader: ExcelReaderBase,
    output_path: Optional[str] = None,
):
    """
    Export found tables to CSV files.

    Args:
        tables: TablesOutput or TablesWithHeadersOutput object
        excel_reader: ExcelReaderBase instance
        output_path: Optional base path for output files
    """
    for idx, table in enumerate(tables.tables):
        # Determine output filename
        if output_path:
            base_path = Path(output_path)
            if len(tables.tables) > 1:
                # Multiple tables: append index
                filename = f"{base_path.stem}_table_{idx+1}{base_path.suffix or '.csv'}"
                out_file = base_path.parent / filename
            else:
                # Single table: use provided name
                out_file = base_path if base_path.suffix else base_path.with_suffix(".csv")
        else:
            # Default naming
            out_file = Path(f"table_{table.sheet_name}_{idx+1}.csv")

        # Get the table data range
        if hasattr(table, "headers"):
            # Table with headers
            full_range = f"{table.header_range.split(':')[0]}:{table.data_range.split(':')[1]}"
        else:
            # Simple table range
            full_range = table.range

        # Get cells from the range
        cells = excel_reader.get_cells_in_range(table.sheet_name, full_range)

        # Organize cells into rows
        range_obj = CellRange.from_excel_notation(full_range)

        rows = []
        for row_idx in range(range_obj.start_row, range_obj.end_row + 1):
            row = []
            for col_idx in range(range_obj.start_col, range_obj.end_col + 1):
                # Find cell at this position
                cell_addr = CellRange.to_column_letter(col_idx) + str(row_idx + 1)
                cell = next((c for c in cells if c.address == cell_addr), None)
                row.append(str(cell.value) if cell and cell.value is not None else "")
            rows.append(row)

        # Write to CSV
        with open(out_file, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerows(rows)

        click.echo(f"Exported table to: {out_file}")
