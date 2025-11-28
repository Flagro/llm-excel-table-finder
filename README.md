# LLM Excel Table Finder

A LangGraph ReAct agent that intelligently finds and extracts tables from Excel files using AI.

## Features

- 🤖 **AI-Powered Table Detection**: Uses LangGraph ReAct agent with OpenAI to intelligently find tables in Excel sheets
- 📊 **Flexible Analysis**: Analyze all sheets or specific sheets
- 📋 **Multiple Output Formats**: Get table ranges as JSON or export directly to CSV
- 🔍 **Header Extraction**: Optionally extract column headers and separate data ranges
- 🏗️ **Extensible Architecture**: Abstract base class design supports both `.xlsx` and `.xls` formats
- 🛠️ **Smart Tools**: Agent uses specialized tools to explore spreadsheets:
  - Get sheet boundaries
  - Request cells with values and formatting
  - Iterate in directions until empty cells

## Installation

This project uses `uv` for package management. First, install `uv` if you haven't:

```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

Then install the package:

```bash
# Clone the repository
git clone <your-repo-url>
cd llm-excel-table-finder

# Install with uv
uv pip install -e .

# Or install with pip
pip install -e .
```

## Prerequisites

You need to set your OpenAI API key:

```bash
export OPENAI_API_KEY='your-api-key-here'
```

## Supported File Formats

The package includes built-in support for:

- **`.xlsx` files**: Using `openpyxl` (automatically installed)
- **`.xls` files**: Using `xlrd` (automatically installed)

Both implementations provide:
- Cell value extraction
- Formatting information (bold, italic, colors, borders, etc.)
- Sheet navigation and bounds detection
- Directional iteration until empty cells

## Usage

### Command Line Interface

```bash
# Find all tables and print JSON
excel-table-finder myfile.xlsx

# Find tables in specific sheets
excel-table-finder myfile.xlsx -s Sheet1 -s Sheet2

# Export tables to CSV
excel-table-finder myfile.xlsx --csv -o output.csv

# Get tables with headers and data ranges
excel-table-finder myfile.xlsx --include-headers

# Use a different OpenAI model
excel-table-finder myfile.xlsx --model gpt-4
```

### Python API

```python
from src.agent import ExcelTableFinderAgent
from src.excel_tools import OpenpyxlReader, XlrdReader

# Create reader for .xlsx files
reader = OpenpyxlReader("myfile.xlsx")

# Or for .xls files
# reader = XlrdReader("myfile.xls")

# Create agent
agent = ExcelTableFinderAgent(
    excel_reader=reader,
    sheet_names=["Sheet1", "Sheet2"],  # Optional, None = all sheets
    model_name="gpt-4o-mini",
    include_headers=True  # Get headers and data ranges
)

# Find tables
result = agent.find_tables()

# Access results
for table in result.tables:
    print(f"Sheet: {table.sheet_name}")
    print(f"Range: {table.range}")
    if hasattr(table, 'headers'):
        print(f"Headers: {table.headers}")
        print(f"Data Range: {table.data_range}")
```

## Architecture

### Agent Flow

The agent uses a LangGraph ReAct pattern:

1. **Initialization**: Agent receives sheets to analyze
2. **Exploration**: Uses tools to explore spreadsheet structure
3. **Analysis**: Identifies table boundaries, headers, and data
4. **Extraction**: Returns structured output with table information

### Tools Available to Agent

1. **get_sheet_bounds**: Get the used range of a sheet
2. **get_cells_in_range**: Get cell values and formatting for a range
3. **iterate_until_empty**: Navigate from a cell until hitting empty cells

### Output Formats

#### Simple Table Ranges

```json
{
  "tables": [
    {
      "sheet_name": "Sheet1",
      "range": "A1:D10",
      "description": "Sales data table"
    }
  ]
}
```

#### Tables with Headers

```json
{
  "tables": [
    {
      "sheet_name": "Sheet1",
      "headers": ["Name", "Age", "City", "Score"],
      "header_range": "A1:D1",
      "data_range": "A2:D10",
      "description": "User information table"
    }
  ]
}
```

## Development

### Running Tests

```bash
# Install dev dependencies
uv pip install -e ".[dev]"

# Run tests
pytest
```

### Code Formatting

```bash
# Format code
black src/

# Lint code
ruff check src/
```

## License

MIT License - see LICENSE file for details

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

