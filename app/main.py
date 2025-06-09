from fastapi import FastAPI, Query, HTTPException
from typing import List, Dict
from .excel_reader import ExcelProcessor
import os

app = FastAPI(
    title="Excel Processor API",
    description="API for processing Excel sheet data",
    version="1.0.0",
)

# Initialize Excel processor
FILE_PATH = os.path.join(os.path.dirname(__file__), "data", "capbudg.xls")
excel_processor = ExcelProcessor(FILE_PATH)

@app.get("/list_tables", response_model=Dict[str, List[str]])
def list_tables():
    """List all tables in the Excel sheet."""
    try:
        tables = excel_processor.get_table_names()
        return {"tables": tables}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/get_table_details")
def get_table_details(table_name: str = Query(..., description="Name of the table to get details for")):
    """Get row names for a specific table."""
    try:
        row_names = excel_processor.get_table_row_names(table_name)
        return {
            "table_name": table_name,
            "row_names": row_names
        }
    except ValueError as e:
        raise HTTPException(status_code=404, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/row_sum")
def row_sum(
    table_name: str = Query(..., description="Name of the table containing the row"),
    row_name: str = Query(..., description="Name of the row to sum values for")
):
    """Calculate sum of numeric values in a specific row."""
    try:
        total = excel_processor.get_row_sum(table_name, row_name)
        return {
            "table_name": table_name,
            "row_name": row_name,
            "sum": total
        }
    except ValueError as e:
        raise HTTPException(status_code=404, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/")
def read_root():
    return {"message": "Excel Processor API is running. Check /docs for API documentation."}