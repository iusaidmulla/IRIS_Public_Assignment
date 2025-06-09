# Excel Processor API

A FastAPI application that reads and processes data from Excel sheets, providing endpoints to interact with table data.

## Features

- Identify and extract tables from Excel sheets
- List all available tables
- Get row names for specific tables
- Calculate sums of numeric values in table rows
- Robust error handling and input validation

## Installation

1. Clone the repository:
   ```bash
   git clone <repository-url>
   cd <repository-directory>

2. Create and activate a virtual environment (recommended):
   python -m venv venv
   source venv/bin/activate  # On Windows use `venv\Scripts\activate`

3. Install dependencies:
   pip install -r requirements.txt

## Usage
1. Place your Excel file in the data directory (default expects capbudg.xls)
2. Run the FastAPI application:
   uvicorn main:app --reload
3. Access the API documentation at:
   http://127.0.0.1:9090/list_tables/

   {
    "tables": [
        "INITIAL INVESTMENT",
        "Initial Investment=",
        "Opportunity cost (if any)=",
        "Lifetime of the investment",
        "Salvage Value at end of project=",
        "Deprec. method(1:St.line;2:DDB)=",
        "Tax Credit (if any )=",
        "Other invest.(non-depreciable)=",
        "Initial Investment in Work. Cap=",
        "Working Capital as % of Rev=",
        "Salvageable fraction at end=",
        "Revenues",
        "Fixed Expenses",
        "Investment",
        "- Tax Credit",
        "Net Investment",
        "+ Working Cap",
        "+ Opp. Cost",
        "+ Other invest.",
        "Initial Investment",
        "Equipment",
        "Working Capital",
        "Lifetime Index",
        "-Var. Expenses",
        "- Fixed Expenses",
        "EBITDA",
        "- Depreciation",
        "EBIT",
        "-Tax",
        "EBIT(1-t)",
        "+ Depreciation",
        "- ∂ Work. Cap",
        "NATCF",
        "Discount Factor",
        "Discounted CF",
        "Book Value (beginning)",
        "Depreciation",
        "BV(ending)"
    ]
}
