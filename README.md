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
3. Access the API documentation at: http://127.0.0.1:9090/list_tables/

![image](https://github.com/user-attachments/assets/9b13818d-c22e-4e28-a59b-f6ccd49113ea)

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

GET http://127.0.0.1:9090/get_table_details?table_name=Initial Investment

Parameters:

table_name: Name of the table to get details for
![image](https://github.com/user-attachments/assets/378abe85-87fe-4203-8275-6794bd2d828d)
Response Example:
{
    "table_name": "Initial Investment",
    "row_names": [
        "Initial Investment",
        "SALVAGE VALUE"
    ]
}

GET http://127.0.0.1:9090/row_sum?table_name=Initial Investment&row_name=Initial Investment

Parameters:
table_name: Name of the table containing the row

row_name: Name of the row to sum values for
![image](https://github.com/user-attachments/assets/0be75e87-9b59-4cbe-93e5-d7be15bdfc70)
Response Example:

{
    "table_name": "Initial Investment",
    "row_name": "Initial Investment",
    "sum": 62484.0
}


## Potential Improvements

### 1. **Support for Modern Excel Formats**  
   - Currently uses `xlrd` (which only supports `.xls` files).  
   - Migrate to `openpyxl` or `pandas` for `.xlsx` support and better performance.  

### 2. **Enhanced Table Detection**  
   - Detect merged cells and structured tables with headers.  
   - Support for tables that span multiple sheets.  

### 3. **Advanced Data Operations**  
   - Add endpoints for:  
     - Column-wise sums (`/column_sum`).  
     - Filtering rows based on conditions (`/filter_rows?table_name=X&condition=value>100`).  
     - Statistical operations (average, median, min/max).  

### 4. **File Upload & Dynamic Processing**  
   - Allow users to upload Excel files via a `POST /upload` endpoint.  
   - Process files on-the-fly instead of relying on a fixed `capbudg.xls`.  

### 5. **Pagination for Large Tables**  
   - If a table has hundreds of rows, add `?limit=10&offset=0` to `/get_table_details`.  

### 6. **Caching Mechanism**  
   - Cache parsed Excel data in memory (e.g., using `functools.lru_cache`) to avoid re-reading the file on every request.  

### 7. **Frontend UI**  
   - Build a simple React/Vue dashboard to:  
     - Upload files.  
     - Visualize tables.  
     - Interact with the API via a GUI.  

### 8. **Unit & Integration Tests**  
   - Add pytest cases for:  
     - Malformed Excel files.  
     - Edge cases (empty sheets, non-numeric values).  
     - API error responses.  

## Missed Edge Cases  

### 1. **Empty or Corrupt Excel Files**  
   - The app crashes if:  
     - `capbudg.xls` is missing.  
     - The file is corrupted (e.g., invalid `.xls` format).  
   - **Fix**: Validate file existence and structure on startup.  

### 2. **Tables with No Numeric Data**  
   - `/row_sum` fails if a row contains only text (e.g., `["Name", "John", "Doe"]`).  
   - **Fix**: Return `0` or a meaningful error (`"No numeric values in row"`).  

### 3. **Malformed Table Names**  
   - If a table name has trailing spaces (`"Revenue  "`), `/get_table_details` may fail.  
   - **Fix**: Trim whitespace in `get_table_names()`.  

### 4. **Very Large Files**  
   - The current implementation loads the entire file into memory.  
   - **Risk**: Crashes with memory-heavy Excel files.  
   - **Fix**: Stream data or use disk-based processing.  

### 5. **Concurrent Requests**  
   - If multiple users hit `/row_sum` simultaneously, thread safety issues may arise.  
   - **Fix**: Use thread-safe caching or a database (e.g., SQLite).  

### 6. **Non-Standard Numeric Formats**  
   - Fails on:  
     - Currency (`"$1,000"`).  
     - Scientific notation (`"1.23E+5"`).  
     - Fractions (`"1 3/4"`).  
   - **Fix**: Enhance `_extract_numeric()` with regex for diverse formats.  

### 7. **Hidden Sheets or Rows**  
   - The app ignores hidden rows/sheets, which might contain valid data.  
   - **Fix**: Add a `?include_hidden=true` parameter.  

### 8. **Special Characters in Row Names**  
   - Row names with symbols (`#`, `@`, emojis) may break parsing.  
   - **Fix**: Sanitize strings or use URL encoding.  

### Why These Matter  
These improvements and edge-case fixes would make the API:  
✅ **More robust** (handles real-world Excel quirks).  
✅ **Scalable** (works with large files/users).  
✅ **User-friendly** (clear errors, flexible inputs).  



