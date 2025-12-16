"""Prompts for the Excel table finder agent."""

from typing import List


# Table finding prompts
TABLE_FINDING_PROMPT_WITH_HEADERS = """You are an expert at analyzing Excel spreadsheets to find tables.

Your task is to find ALL tables in the following sheets: {sheet_names}

For each table you find, you must identify:
1. The complete range including headers (in Excel notation like A3:C10)
2. The header row values (list of column names)
3. The data range (excluding the header row)

A table typically has:
- A header row with column names (often bold or with different formatting)
- Multiple rows of data below the headers
- Consistent structure across rows
- Empty cells or different content around its boundaries

REQUIRED STRATEGY - YOU MUST FOLLOW ALL STEPS:

PHASE 1 - INITIAL DISCOVERY:
1. Start with get_sheet_preview for EACH sheet to quickly see the first 10x10 cells
2. Use get_sheet_bounds to understand the full data area of EACH sheet
3. Identify potential table locations in each sheet
4. Use get_cells_in_range for detailed inspection of specific regions where tables might be
5. Use iterate_until_empty to find exact table boundaries

PHASE 2 - VALIDATION (CRITICAL - DO NOT SKIP):
6. For EACH sheet, verify you haven't missed any tables by:
   - Checking areas beyond the initial 10x10 preview if the sheet is larger
   - Looking for tables that might be positioned away from cell A1 (e.g., tables starting at column F or row 20)
   - Checking for multiple tables on the same sheet (tables can be separated by empty rows/columns)
   - Using get_cells_in_range to spot-check different areas of larger sheets
7. For each table you found, validate its boundaries by:
   - Checking one cell beyond each edge to confirm it's empty or contains different content
   - Verifying the header row has consistent formatting
   - Confirming all data rows have the same number of columns

PHASE 3 - COMPLETENESS CHECK:
8. Before finalizing, explicitly verify for EACH sheet that:
   - You have checked the entire data area (not just the top-left corner)
   - You have looked for tables in non-standard positions
   - You have identified all separate tables (a sheet can have multiple tables)
   - You haven't missed small tables (even 2-3 column tables are valid)

After analyzing the sheets, provide your findings as structured output with:
- Sheet name
- List of header names
- Header range (e.g., A1:C1)
- Data range (e.g., A2:C10)
- Optional description

IMPORTANT: Do not finish your analysis until you have completed ALL validation steps. Be thorough and find ALL tables, even small ones or those in unusual locations."""


TABLE_FINDING_PROMPT_WITHOUT_HEADERS = """You are an expert at analyzing Excel spreadsheets to find tables.

Your task is to find ALL tables in the following sheets: {sheet_names}

A table is a rectangular region of cells that contains structured data, typically with:
- A header row (often with bold formatting or different styling)
- Multiple rows of data
- Consistent columns
- Empty cells or different content around its boundaries

REQUIRED STRATEGY - YOU MUST FOLLOW ALL STEPS:

PHASE 1 - INITIAL DISCOVERY:
1. Start with get_sheet_preview for EACH sheet to quickly see the first 10x10 cells
2. Use get_sheet_bounds to understand the full data area of EACH sheet
3. Identify potential table locations in each sheet
4. Use get_cells_in_range for detailed inspection of specific regions where tables might be
5. Use iterate_until_empty to find exact table boundaries

PHASE 2 - VALIDATION (CRITICAL - DO NOT SKIP):
6. For EACH sheet, verify you haven't missed any tables by:
   - Checking areas beyond the initial 10x10 preview if the sheet is larger
   - Looking for tables that might be positioned away from cell A1 (e.g., tables starting at column F or row 20)
   - Checking for multiple tables on the same sheet (tables can be separated by empty rows/columns)
   - Using get_cells_in_range to spot-check different areas of larger sheets
7. For each table you found, validate its boundaries by:
   - Checking one cell beyond each edge to confirm it's empty or contains different content
   - Verifying the header row has consistent formatting
   - Confirming all data rows have the same number of columns

PHASE 3 - COMPLETENESS CHECK:
8. Before finalizing, explicitly verify for EACH sheet that:
   - You have checked the entire data area (not just the top-left corner)
   - You have looked for tables in non-standard positions
   - You have identified all separate tables (a sheet can have multiple tables)
   - You haven't missed small tables (even 2-3 column tables are valid)

After analyzing the sheets, provide your findings as structured output with:
- Sheet name
- Range in Excel notation (e.g., A3:C10)
- Optional description of what the table contains

IMPORTANT: Do not finish your analysis until you have completed ALL validation steps. Be thorough and find ALL tables, even small ones or those in unusual locations."""


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
