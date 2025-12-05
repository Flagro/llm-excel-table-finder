"""LangGraph ReAct agent for finding tables in Excel files."""

from typing import List, Optional, Literal
from pydantic import BaseModel, Field

from langgraph.prebuilt import create_react_agent
from langchain_core.messages import HumanMessage
from langchain_core.tools import tool
from langchain_openai import ChatOpenAI

from src.excel_tools import ExcelReaderBase, Direction
from src.prompts import (
    get_table_finding_prompt,
    get_structured_output_prompt,
)


# Structured output models
class TableRange(BaseModel):
    """Model for a single table range."""

    sheet_name: str = Field(description="Name of the sheet containing the table")
    range: str = Field(description="Excel range notation for the table (e.g., A3:C10)")
    description: Optional[str] = Field(
        default=None, description="Optional description of the table"
    )


class TableWithHeaders(BaseModel):
    """Model for a table with headers and data range."""

    sheet_name: str = Field(description="Name of the sheet containing the table")
    headers: List[str] = Field(description="List of column header names")
    data_range: str = Field(
        description="Excel range notation for the data rows (excluding headers)"
    )
    header_range: str = Field(description="Excel range notation for the header row")
    description: Optional[str] = Field(
        default=None, description="Optional description of the table"
    )


class TablesOutput(BaseModel):
    """Output model for found tables."""

    tables: List[TableRange] = Field(description="List of found tables with their ranges")


class TablesWithHeadersOutput(BaseModel):
    """Output model for found tables with headers."""

    tables: List[TableWithHeaders] = Field(
        description="List of found tables with headers and data ranges"
    )


class ExcelTableFinderAgent:
    """LangGraph ReAct agent for finding tables in Excel files."""

    def __init__(
        self,
        excel_reader: ExcelReaderBase,
        sheet_names: Optional[List[str]] = None,
        model_name: str = "gpt-4o-mini",
        include_headers: bool = False,
    ):
        """
        Initialize the Excel table finder agent.

        Args:
            excel_reader: Instance of ExcelReaderBase for Excel operations
            sheet_names: List of sheet names to analyze (None = all sheets)
            model_name: Name of the LLM model to use
            include_headers: Whether to extract headers and separate data ranges
        """
        self.excel_reader = excel_reader
        self.include_headers = include_headers

        # Get sheet names to analyze
        if sheet_names is None or len(sheet_names) == 0:
            self.sheet_names = excel_reader.get_sheet_names()
        else:
            self.sheet_names = sheet_names

        # Initialize LLM
        self.llm = ChatOpenAI(model=model_name, temperature=0)

        # Create tools
        self.tools = self._create_tools()

        # Create the ReAct agent
        self.agent = create_react_agent(self.llm, self.tools)

    def _create_tools(self):
        """Create the tools for the agent."""

        @tool
        def get_sheet_bounds(sheet_name: str) -> str:
            """
            Get the boundaries of a sheet in Excel notation.

            Args:
                sheet_name: Name of the sheet to get bounds for

            Returns:
                Range in Excel notation (e.g., "A1:Z100")
            """
            return self.excel_reader.get_sheet_bounds(sheet_name)

        @tool
        def get_cells_in_range(sheet_name: str, range_notation: str) -> str:
            """
            Get cells with values and formatting in the requested area.

            Args:
                sheet_name: Name of the sheet
                range_notation: Range in Excel notation (e.g., "A3:C10")

            Returns:
                JSON string with cell data including addresses, values, and formatting
            """
            cells = self.excel_reader.get_cells_in_range(sheet_name, range_notation)
            result = []
            for cell in cells:
                result.append(
                    {
                        "address": cell.address,
                        "value": str(cell.value) if cell.value is not None else "",
                        "formatting": cell.formatting,
                    }
                )
            return str(result)

        @tool
        def iterate_until_empty(
            sheet_name: str, start_cell: str, direction: Literal["up", "down", "left", "right"]
        ) -> str:
            """
            Iterate from a cell in a direction until an empty cell is found.

            Args:
                sheet_name: Name of the sheet
                start_cell: Starting cell in Excel notation (e.g., "A3")
                direction: Direction to iterate - must be one of: "up", "down", "left", "right"

            Returns:
                JSON string with list of encountered cells (excluding the empty cell)
            """
            cells = self.excel_reader.iterate_until_empty(
                sheet_name, start_cell, Direction(direction)
            )
            result = []
            for cell in cells:
                result.append(
                    {
                        "address": cell.address,
                        "value": str(cell.value) if cell.value is not None else "",
                        "formatting": cell.formatting,
                    }
                )
            return str(result)

        return [get_sheet_bounds, get_cells_in_range, iterate_until_empty]

    def find_tables(self) -> TablesOutput | TablesWithHeadersOutput:
        """
        Find tables in the Excel file.

        Returns:
            TablesOutput or TablesWithHeadersOutput depending on include_headers flag
        """
        # Create the initial prompt
        prompt = get_table_finding_prompt(self.sheet_names, self.include_headers)

        # Run the agent
        final_state = self.agent.invoke({"messages": [HumanMessage(content=prompt)]})

        # Extract the final response and structure it
        last_message = final_state["messages"][-1]

        # Use structured output to parse the response
        if self.include_headers:
            structured_llm = self.llm.with_structured_output(TablesWithHeadersOutput)
        else:
            structured_llm = self.llm.with_structured_output(TablesOutput)

        result = structured_llm.invoke(
            [
                HumanMessage(
                    content=get_structured_output_prompt(last_message.content, self.include_headers)
                )
            ]
        )

        return result
