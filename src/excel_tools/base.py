"""Abstract base class and common data structures for Excel operations."""

from abc import ABC, abstractmethod
from typing import List, Dict, Any, Tuple
from dataclasses import dataclass
from enum import Enum


class Direction(str, Enum):
    """Direction for cell iteration."""

    UP = "up"
    DOWN = "down"
    LEFT = "left"
    RIGHT = "right"


@dataclass
class CellData:
    """Data structure for cell information."""

    address: str
    value: Any
    formatting: Dict[str, Any]


@dataclass
class CellRange:
    """Data structure for cell range."""

    start_col: int  # 0-indexed
    start_row: int  # 0-indexed
    end_col: int  # 0-indexed
    end_row: int  # 0-indexed

    def to_excel_notation(self) -> str:
        """Convert to Excel A1 notation (e.g., A3:C10)."""
        start = self.to_column_letter(self.start_col) + str(self.start_row + 1)
        end = self.to_column_letter(self.end_col) + str(self.end_row + 1)
        return f"{start}:{end}"

    @staticmethod
    def to_column_letter(col: int) -> str:
        """Convert column index to Excel letter (0 -> A, 1 -> B, etc.)."""
        result = ""
        while col >= 0:
            result = chr(col % 26 + ord("A")) + result
            col = col // 26 - 1
        return result

    @classmethod
    def from_excel_notation(cls, notation: str) -> "CellRange":
        """Parse Excel A1 notation (e.g., A3:C10) into CellRange."""
        if ":" not in notation:
            # Single cell
            col, row = cls._parse_cell_address(notation)
            return cls(col, row, col, row)

        start_cell, end_cell = notation.split(":")
        start_col, start_row = cls._parse_cell_address(start_cell)
        end_col, end_row = cls._parse_cell_address(end_cell)

        return cls(start_col, start_row, end_col, end_row)

    @staticmethod
    def _parse_cell_address(address: str) -> Tuple[int, int]:
        """Parse cell address like 'A3' into (col_idx, row_idx)."""
        col_letters = ""
        row_digits = ""

        for char in address:
            if char.isalpha():
                col_letters += char.upper()
            elif char.isdigit():
                row_digits += char

        # Convert column letters to index
        col_idx = 0
        for char in col_letters:
            col_idx = col_idx * 26 + (ord(char) - ord("A") + 1)
        col_idx -= 1  # Convert to 0-indexed

        row_idx = int(row_digits) - 1  # Convert to 0-indexed

        return col_idx, row_idx


class ExcelReaderBase(ABC):
    """Abstract base class for Excel file operations."""

    def __init__(self, file_path: str):
        """Initialize the Excel reader with a file path."""
        self.file_path = file_path

    @abstractmethod
    def get_sheet_names(self) -> List[str]:
        """Get list of all sheet names in the workbook."""
        pass

    @abstractmethod
    def get_sheet_bounds(self, sheet_name: str) -> str:
        """
        Get the used range of a sheet in Excel notation.

        Args:
            sheet_name: Name of the sheet

        Returns:
            Range in Excel notation (e.g., "A1:Z100")
        """
        pass

    @abstractmethod
    def get_cells_in_range(self, sheet_name: str, range_notation: str) -> List[CellData]:
        """
        Get cells with values and formatting in the specified range.

        Args:
            sheet_name: Name of the sheet
            range_notation: Range in Excel notation (e.g., "A3:C10")

        Returns:
            List of CellData objects
        """
        pass

    @abstractmethod
    def iterate_until_empty(
        self, sheet_name: str, start_cell: str, direction: Direction
    ) -> List[CellData]:
        """
        Iterate from a cell in a direction until an empty cell is found.

        Args:
            sheet_name: Name of the sheet
            start_cell: Starting cell in Excel notation (e.g., "A3")
            direction: Direction to iterate (up, down, left, right)

        Returns:
            List of CellData objects encountered (excluding the empty cell)
        """
        pass

    @abstractmethod
    def close(self):
        """Close the workbook and free resources."""
        pass

    def __enter__(self):
        """Context manager entry."""
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit."""
        self.close()
