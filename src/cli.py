"""CLI interface for the Excel table finder agent."""

import json
import sys
from typing import Optional

import click

from src.agent import ExcelTableFinderAgent
from src.io import export_to_csv, get_excel_reader


@click.command()
@click.argument("file_path", type=click.Path(exists=True))
@click.option(
    "--sheet",
    "-s",
    multiple=True,
    help="Sheet name(s) to analyze. Can be specified multiple times. Use quotes for names with spaces. If not provided, all sheets are analyzed.",
)
@click.option(
    "--csv", is_flag=True, help="Export found tables to CSV files instead of returning JSON."
)
@click.option(
    "--output",
    "-o",
    type=click.Path(),
    help="Output file path for CSV export (used with --csv flag). For multiple tables, index will be appended.",
)
@click.option(
    "--include-headers",
    is_flag=True,
    help="Extract headers and separate data ranges in the output.",
)
@click.option(
    "--model",
    default="gpt-4o-mini",
    help="OpenAI model to use for the agent (default: gpt-4o-mini).",
)
@click.option(
    "--api-key",
    envvar="OPENAI_API_KEY",
    help="OpenAI API key (defaults to OPENAI_API_KEY environment variable).",
)
@click.option(
    "--max-recursion",
    type=int,
    default=25,
    help="Maximum number of tool call iterations for the agent (default: 25).",
)
def main(
    file_path: str,
    sheet: tuple,
    csv: bool,
    output: Optional[str],
    include_headers: bool,
    model: str,
    api_key: Optional[str],
    max_recursion: int,
):
    """
    Excel Table Finder - Find tables in Excel files using AI.

    FILE_PATH: Path to the Excel file (.xlsx, .xlsm, .xls, or .xlsb)

    Examples:

        # Find all tables and print JSON
        excel-table-finder myfile.xlsx

        # Find tables in specific sheets
        excel-table-finder myfile.xlsx -s Sheet1 -s Sheet2

        # Find tables in sheets with spaces in names (use quotes)
        excel-table-finder myfile.xlsx -s "Sales Data" -s "Monthly Report"

        # Export tables to CSV
        excel-table-finder myfile.xlsx --csv -o output.csv

        # Get tables with headers
        excel-table-finder myfile.xlsx --include-headers

        # Limit the number of tool calls
        excel-table-finder myfile.xlsx --max-recursion 15
    """
    try:
        # Validate OpenAI API key
        if not api_key:
            click.echo("Error: OpenAI API key is not provided.", err=True)
            click.echo("\nPlease provide your OpenAI API key either:", err=True)
            click.echo(
                "  1. Via environment variable: export OPENAI_API_KEY='your-api-key-here'", err=True
            )
            click.echo("  2. Via command line option: --api-key your-api-key-here", err=True)
            sys.exit(1)

        # Validate options
        if output and not csv:
            click.echo("Warning: --output is only used with --csv flag", err=True)

        # Get the appropriate Excel reader
        click.echo(f"Opening file: {file_path}")
        excel_reader = get_excel_reader(file_path)

        # Convert sheet tuple to list
        sheet_names = list(sheet) if sheet else None

        # Create and run the agent
        click.echo("Analyzing spreadsheet to find tables...")
        agent = ExcelTableFinderAgent(
            excel_reader=excel_reader,
            sheet_names=sheet_names,
            model_name=model,
            include_headers=include_headers or csv,  # Always include headers for CSV
            api_key=api_key,
            max_recursion=max_recursion,
        )

        result = agent.find_tables()

        # Handle output
        if csv:
            click.echo(f"Found {len(result.tables)} table(s)")
            try:
                export_to_csv(result, excel_reader, output)
            except Exception as e:
                click.echo(f"Error exporting to CSV: {e}", err=True)
                sys.exit(1)
        else:
            # Output JSON
            click.echo(json.dumps(result.model_dump(), indent=2))

        # Close the reader
        excel_reader.close()

    except ImportError as e:
        click.echo(f"Error: {e}", err=True)
        click.echo("\nPlease install the required package:", err=True)
        if "openpyxl" in str(e):
            click.echo("  pip install openpyxl", err=True)
        elif "xlrd" in str(e):
            click.echo("  pip install xlrd", err=True)
        elif "pyxlsb" in str(e):
            click.echo("  pip install pyxlsb", err=True)
        sys.exit(1)
    except Exception as e:
        click.echo(f"Error: {e}", err=True)
        sys.exit(1)


if __name__ == "__main__":
    main()
