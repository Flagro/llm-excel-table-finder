"""Excel reader implementation for .xls files using xlrd."""

from typing import List, Dict, Any

from src.excel_tools.base import ExcelReaderBase, CellData, CellRange, Direction


class XlrdReader(ExcelReaderBase):
    """Excel reader implementation for .xls files using xlrd."""

    def __init__(self, file_path: str):
        """Initialize the xlrd reader."""
        super().__init__(file_path)
        try:
            import xlrd
        except ImportError:
            raise ImportError("xlrd is required for .xls files. Install it with: pip install xlrd")

        self.workbook = xlrd.open_workbook(file_path, formatting_info=True)

    def get_sheet_names(self) -> List[str]:
        """Get list of all sheet names in the workbook."""
        return self.workbook.sheet_names()

    def get_sheet_bounds(self, sheet_name: str) -> str:
        """Get the used range of a sheet in Excel notation."""
        sheet = self.workbook.sheet_by_name(sheet_name)

        if sheet.nrows == 0 or sheet.ncols == 0:
            return "A1:A1"

        # xlrd provides nrows and ncols (0-indexed counts)
        max_row = sheet.nrows
        max_col = sheet.ncols

        # Convert to Excel notation
        start_cell = "A1"
        end_cell = CellRange.to_column_letter(max_col - 1) + str(max_row)

        return f"{start_cell}:{end_cell}"

    def get_cells_in_range(self, sheet_name: str, range_notation: str) -> List[CellData]:
        """Get cells with values and formatting in the specified range."""
        sheet = self.workbook.sheet_by_name(sheet_name)
        cell_range = CellRange.from_excel_notation(range_notation)

        cells = []
        for row in range(cell_range.start_row, min(cell_range.end_row + 1, sheet.nrows)):
            for col in range(cell_range.start_col, min(cell_range.end_col + 1, sheet.ncols)):
                cell = sheet.cell(row, col)
                address = CellRange.to_column_letter(col) + str(row + 1)

                # Extract formatting information
                formatting = self._get_cell_formatting(sheet, row, col)

                # Get cell value
                value = self._get_cell_value(cell)

                cells.append(CellData(address=address, value=value, formatting=formatting))

        return cells

    def iterate_until_empty(
        self, sheet_name: str, start_cell: str, direction: Direction
    ) -> List[CellData]:
        """Iterate from a cell in a direction until an empty cell is found."""
        sheet = self.workbook.sheet_by_name(sheet_name)
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

        while True:
            # Check bounds
            if row < 0 or col < 0 or row >= sheet.nrows or col >= sheet.ncols:
                break

            # Get the cell at current position
            cell = sheet.cell(row, col)
            value = self._get_cell_value(cell)

            # Check if cell is empty
            if value is None or str(value).strip() == "":
                break

            # Add cell to results
            address = CellRange.to_column_letter(col) + str(row + 1)
            formatting = self._get_cell_formatting(sheet, row, col)
            cells.append(CellData(address=address, value=value, formatting=formatting))

            # Move to next cell
            col += delta_col
            row += delta_row

        return cells

    def close(self):
        """Close the workbook and free resources."""
        # xlrd doesn't require explicit closing, but we set to None for consistency
        self.workbook = None

    @staticmethod
    def _get_cell_value(cell):
        """Get the actual value from an xlrd cell."""
        import xlrd

        # Handle different cell types
        if cell.ctype == xlrd.XL_CELL_EMPTY:
            return None
        elif cell.ctype == xlrd.XL_CELL_TEXT:
            return cell.value
        elif cell.ctype == xlrd.XL_CELL_NUMBER:
            return cell.value
        elif cell.ctype == xlrd.XL_CELL_DATE:
            # Convert Excel date to a tuple
            date_tuple = xlrd.xldate_as_tuple(cell.value, 0)
            return f"{date_tuple[0]}-{date_tuple[1]:02d}-{date_tuple[2]:02d}"
        elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
            return bool(cell.value)
        else:
            return cell.value

    def _get_cell_formatting(self, sheet, row: int, col: int) -> Dict[str, Any]:
        """Extract formatting information from a cell."""
        import xlrd

        formatting = {}

        try:
            # Get the cell's XF record
            cell = sheet.cell(row, col)
            xf_index = cell.xf_index

            if xf_index is not None:
                xf = self.workbook.format_map.get(self.workbook.xf_list[xf_index].format_key)

                # Get font information
                font = self.workbook.font_list[self.workbook.xf_list[xf_index].font_index]
                formatting["bold"] = bool(font.bold)
                formatting["italic"] = bool(font.italic)
                formatting["underline"] = bool(font.underline_type)
                formatting["font_size"] = font.height / 20  # Convert to points

                # Get number format
                if xf:
                    formatting["number_format"] = xf.format_str

                # Get alignment
                alignment = self.workbook.xf_list[xf_index].alignment
                formatting["horizontal_alignment"] = alignment.hor_align
                formatting["vertical_alignment"] = alignment.vert_align

                # Background color/pattern
                background = self.workbook.xf_list[xf_index].background
                formatting["background_color_index"] = background.pattern_colour_index

        except (AttributeError, IndexError, KeyError):
            # If we can't get formatting info, just return empty dict
            pass

        return formatting
