"""Excel reader implementation for .xlsb files using pyxlsb."""

from typing import List, Dict, Any
from src.excel_tools.base import ExcelReaderBase, CellData, CellRange, Direction


class PyxlsbReader(ExcelReaderBase):
    """Excel reader implementation for .xlsb files using pyxlsb."""

    def __init__(self, file_path: str):
        """Initialize the pyxlsb reader."""
        super().__init__(file_path)
        try:
            from pyxlsb import open_workbook
        except ImportError:
            raise ImportError(
                "pyxlsb is required for .xlsb files. Install it with: pip install pyxlsb"
            )

        self.workbook = open_workbook(file_path)
        self._sheet_dimensions = {}  # Cache for sheet dimensions

    def get_sheet_names(self) -> List[str]:
        """Get list of all sheet names in the workbook."""
        return self.workbook.sheets

    def get_sheet_bounds(self, sheet_name: str) -> str:
        """Get the used range of a sheet in Excel notation."""
        dims = self._get_sheet_dimensions(sheet_name)

        if dims is None or dims["max_row"] == 0 or dims["max_col"] == 0:
            return "A1:A1"

        # Convert to Excel notation
        start_cell = CellRange.to_column_letter(0) + "1"
        end_cell = CellRange.to_column_letter(dims["max_col"]) + str(dims["max_row"] + 1)

        return f"{start_cell}:{end_cell}"

    def get_cells_in_range(self, sheet_name: str, range_notation: str) -> List[CellData]:
        """Get cells with values and formatting in the specified range."""
        cell_range = CellRange.from_excel_notation(range_notation)
        cells = []

        with self.workbook.get_sheet(sheet_name) as sheet:
            # pyxlsb iterates through rows
            for row_idx, row in enumerate(sheet.rows()):
                # Skip rows before the range
                if row_idx < cell_range.start_row:
                    continue
                # Stop after the range
                if row_idx > cell_range.end_row:
                    break

                # Process cells in this row within the column range
                for col_idx in range(cell_range.start_col, cell_range.end_col + 1):
                    value = None
                    if col_idx < len(row):
                        cell_value = row[col_idx].v if row[col_idx] is not None else None
                        value = cell_value

                    address = CellRange.to_column_letter(col_idx) + str(row_idx + 1)

                    # pyxlsb doesn't provide formatting info easily, so we'll use minimal formatting
                    formatting = self._get_minimal_formatting()

                    cells.append(CellData(address=address, value=value, formatting=formatting))

        return cells

    def iterate_until_empty(
        self, sheet_name: str, start_cell: str, direction: Direction
    ) -> List[CellData]:
        """Iterate from a cell in a direction until an empty cell is found."""
        col, row = CellRange._parse_cell_address(start_cell)
        cells = []

        # Define direction deltas
        deltas = {
            Direction.UP: (0, -1),
            Direction.DOWN: (0, 1),
            Direction.LEFT: (-1, 0),
            Direction.RIGHT: (1, 0),
        }

        delta_col, delta_row = deltas[direction]

        # Get sheet dimensions to avoid going out of bounds
        dims = self._get_sheet_dimensions(sheet_name)

        if dims is None:
            return cells

        # For pyxlsb, we need to read the data efficiently
        # We'll cache the sheet data for this operation
        sheet_data = self._get_sheet_data(sheet_name)

        while True:
            # Check bounds
            if row < 0 or col < 0:
                break
            if row > dims["max_row"] or col > dims["max_col"]:
                break

            # Get cell value from cached data
            value = None
            if row < len(sheet_data) and col < len(sheet_data[row]):
                value = sheet_data[row][col]

            # Check if cell is empty
            if value is None or str(value).strip() == "":
                break

            # Add cell to results
            address = CellRange.to_column_letter(col) + str(row + 1)
            formatting = self._get_minimal_formatting()
            cells.append(CellData(address=address, value=value, formatting=formatting))

            # Move to next cell
            col += delta_col
            row += delta_row

        return cells

    def get_sheet_preview(
        self, sheet_name: str, max_rows: int = 10, max_cols: int = 10
    ) -> List[CellData]:
        """Get a preview of the sheet (first N rows and columns)."""
        cells = []

        with self.workbook.get_sheet(sheet_name) as sheet:
            for row_idx, row in enumerate(sheet.rows()):
                if row_idx >= max_rows:
                    break

                for col_idx in range(min(max_cols, len(row))):
                    value = row[col_idx].v if row[col_idx] is not None else None
                    address = CellRange.to_column_letter(col_idx) + str(row_idx + 1)
                    formatting = self._get_minimal_formatting()

                    cells.append(CellData(address=address, value=value, formatting=formatting))

        return cells

    def get_last_non_empty_cell_in_column(self, sheet_name: str, column: str) -> CellData | None:
        """Find the last non-empty cell in a column."""
        # Convert column letter to index
        col_idx = 0
        for char in column.upper():
            col_idx = col_idx * 26 + (ord(char) - ord("A") + 1)
        col_idx -= 1  # Convert to 0-indexed

        # Get all sheet data
        sheet_data = self._get_sheet_data(sheet_name)

        # Iterate from bottom to top
        for row_idx in range(len(sheet_data) - 1, -1, -1):
            if col_idx < len(sheet_data[row_idx]):
                value = sheet_data[row_idx][col_idx]
                if value is not None and str(value).strip() != "":
                    address = column + str(row_idx + 1)
                    formatting = self._get_minimal_formatting()
                    return CellData(address=address, value=value, formatting=formatting)

        return None

    def get_last_non_empty_cell_in_row(self, sheet_name: str, row: int) -> CellData | None:
        """Find the last non-empty cell in a row."""
        # Get all sheet data
        sheet_data = self._get_sheet_data(sheet_name)

        row_idx = row - 1  # Convert to 0-indexed

        if row_idx < 0 or row_idx >= len(sheet_data):
            return None

        # Iterate from right to left
        for col_idx in range(len(sheet_data[row_idx]) - 1, -1, -1):
            value = sheet_data[row_idx][col_idx]
            if value is not None and str(value).strip() != "":
                address = CellRange.to_column_letter(col_idx) + str(row)
                formatting = self._get_minimal_formatting()
                return CellData(address=address, value=value, formatting=formatting)

        return None

    def close(self):
        """Close the workbook and free resources."""
        if hasattr(self, "workbook"):
            self.workbook.close()

    def _get_sheet_dimensions(self, sheet_name: str) -> Dict[str, int] | None:
        """
        Get the dimensions of a sheet (cached).

        Returns:
            Dict with 'max_row' and 'max_col' (0-indexed), or None if sheet is empty
        """
        if sheet_name in self._sheet_dimensions:
            return self._sheet_dimensions[sheet_name]

        max_row = -1
        max_col = -1

        with self.workbook.get_sheet(sheet_name) as sheet:
            for row_idx, row in enumerate(sheet.rows()):
                max_row = row_idx
                # Find the rightmost non-empty cell in this row
                for col_idx in range(len(row) - 1, -1, -1):
                    if row[col_idx] is not None and row[col_idx].v is not None:
                        max_col = max(max_col, col_idx)
                        break

        if max_row == -1 or max_col == -1:
            self._sheet_dimensions[sheet_name] = None
            return None

        dims = {"max_row": max_row, "max_col": max_col}
        self._sheet_dimensions[sheet_name] = dims
        return dims

    def _get_sheet_data(self, sheet_name: str) -> List[List[Any]]:
        """
        Get all data from a sheet as a 2D list (cached).
        This is used to avoid re-reading the sheet multiple times.
        """
        # Note: This loads entire sheet into memory
        # For very large sheets, this might not be ideal
        sheet_data = []

        with self.workbook.get_sheet(sheet_name) as sheet:
            for row in sheet.rows():
                row_data = []
                for cell in row:
                    value = cell.v if cell is not None else None
                    row_data.append(value)
                sheet_data.append(row_data)

        return sheet_data

    @staticmethod
    def _get_minimal_formatting() -> Dict[str, Any]:
        """
        Return minimal formatting dict.

        Note: pyxlsb doesn't easily expose cell formatting information,
        so we return an empty dict or minimal info.
        """
        return {
            "bold": None,
            "italic": None,
            "underline": None,
            "font_size": None,
            "font_color": None,
            "fill_color": None,
            "has_border": None,
            "number_format": None,
            "horizontal_alignment": None,
            "vertical_alignment": None,
        }
