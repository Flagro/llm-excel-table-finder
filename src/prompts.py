"""Prompts for the Excel table finder agent."""

from typing import List


# Table finding prompts
TABLE_FINDING_PROMPT_WITH_HEADERS = """You are an expert at analyzing Excel spreadsheets to find tables.

Your task is to find all tables in the following sheets: {sheet_names}

For each table you find, you must identify:
1. The complete range including headers (in Excel notation like A3:C10)
2. The header row values (list of column names)
3. The data range (excluding the header row)

A table typically has:
- A header row with column names (often bold or with different formatting)
- Multiple rows of data below the headers
- Consistent structure across rows
- Empty cells or different content around its boundaries

RECOMMENDED STRATEGY:
1. Start with get_sheet_preview for each sheet to quickly see the first 10x10 cells
2. Use get_sheet_bounds to understand the full data area
3. Use get_cells_in_range for detailed inspection of specific regions
4. Use iterate_until_empty to find exact table boundaries

After analyzing the sheets, provide your findings as structured output with:
- Sheet name
- List of header names
- Header range (e.g., A1:C1)
- Data range (e.g., A2:C10)
- Optional description

Be thorough and find all tables, even small ones."""


TABLE_FINDING_PROMPT_WITHOUT_HEADERS = """You are an expert at analyzing Excel spreadsheets to find tables.

Your task is to find all tables in the following sheets: {sheet_names}

A table is a rectangular region of cells that contains structured data, typically with:
- A header row (often with bold formatting or different styling)
- Multiple rows of data
- Consistent columns
- Empty cells or different content around its boundaries

RECOMMENDED STRATEGY:
1. Start with get_sheet_preview for each sheet to quickly see the first 10x10 cells
2. Use get_sheet_bounds to understand the full data area
3. Use get_cells_in_range for detailed inspection of specific regions
4. Use iterate_until_empty to find exact table boundaries

After analyzing the sheets, provide your findings as structured output with:
- Sheet name
- Range in Excel notation (e.g., A3:C10)
- Optional description of what the table contains

Be thorough and find all tables, even small ones."""


# Structured output extraction prompts
STRUCTURED_OUTPUT_PROMPT_WITH_HEADERS = """Based on your analysis, extract all found tables with their headers.
Previous conversation: {last_message_content}

Provide the structured output with all tables you found."""


STRUCTURED_OUTPUT_PROMPT_WITHOUT_HEADERS = """Based on your analysis, extract all found table ranges.
Previous conversation: {last_message_content}

Provide the structured output with all table ranges you found."""


def get_table_finding_prompt(sheet_names: List[str], include_headers: bool) -> str:
    """
    Get the prompt for finding tables in Excel sheets.

    Args:
        sheet_names: List of sheet names to analyze
        include_headers: Whether to include header extraction instructions

    Returns:
        The formatted prompt string
    """
    sheet_names_str = ", ".join(sheet_names)
    if include_headers:
        return TABLE_FINDING_PROMPT_WITH_HEADERS.format(sheet_names=sheet_names_str)
    else:
        return TABLE_FINDING_PROMPT_WITHOUT_HEADERS.format(sheet_names=sheet_names_str)


def get_structured_output_prompt(last_message_content: str, include_headers: bool) -> str:
    """
    Get the prompt for extracting structured output from the analysis.

    Args:
        last_message_content: Content of the last message from the agent
        include_headers: Whether to extract headers or just ranges

    Returns:
        The formatted prompt string
    """
    if include_headers:
        return STRUCTURED_OUTPUT_PROMPT_WITH_HEADERS.format(
            last_message_content=last_message_content
        )
    else:
        return STRUCTURED_OUTPUT_PROMPT_WITHOUT_HEADERS.format(
            last_message_content=last_message_content
        )
