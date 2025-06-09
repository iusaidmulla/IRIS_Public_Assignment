import xlrd
from typing import Dict, List, Optional, Union
import re

class ExcelProcessor:
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.wb = xlrd.open_workbook(file_path)
        self.sheet = self.wb.sheet_by_index(0)  # Use the first sheet
        self.tables = self._identify_tables()

    def _identify_tables(self) -> Dict[str, Dict]:
        """Improved table identification that handles the specific structure of your Excel file."""
        tables = {}
        current_table = None
        table_data = []
        
        # First pass: Identify all potential table headers
        table_headers = []
        for row_idx in range(self.sheet.nrows):
            row = self.sheet.row_values(row_idx)
            if row and isinstance(row[0], str) and row[0].strip() and any(cell != '' for cell in row[1:]):
                table_headers.append((row_idx, row[0].strip()))
        
        # Second pass: Build tables based on headers
        for i, (row_idx, header) in enumerate(table_headers):
            start_row = row_idx
            end_row = self.sheet.nrows if i == len(table_headers)-1 else table_headers[i+1][0]
            
            table_rows = []
            for r in range(start_row, end_row):
                row = self.sheet.row_values(r)
                if any(cell != '' for cell in row):  # Skip completely empty rows
                    table_rows.append(row)
            
            if table_rows:
                tables[header] = {
                    'data': table_rows,
                    'start_row': start_row,
                    'end_row': end_row - 1
                }
        
        return tables

    def _is_numeric(self, value: str) -> bool:
        """Check if a string can be converted to a number."""
        try:
            float(value)
            return True
        except ValueError:
            return False

    def get_table_names(self) -> List[str]:
        """Get list of all table names."""
        return list(self.tables.keys())

    def get_table_row_names(self, table_name: str) -> List[str]:
        """Get row names (first column values) for a specific table."""
        if table_name not in self.tables:
            raise ValueError(f"Table '{table_name}' not found in Excel sheet")
        
        table_data = self.tables[table_name]['data']
        row_names = []
        
        for row in table_data:
            if row and row[0] and isinstance(row[0], str):
                row_name = row[0].strip()
                # Skip empty row names and section headers
                if row_name and not row_name.startswith("Equity Analysis") and not row_name.startswith("INPUT SHEET"):
                    row_names.append(row_name)
        
        return row_names

    def get_row_sum(self, table_name: str, row_name: str) -> float:
        """Calculate sum of numeric values in a specific row."""
        if table_name not in self.tables:
            raise ValueError(f"Table '{table_name}' not found in Excel sheet")
        
        table_data = self.tables[table_name]['data']
        row_values = None
        
        # Find the row matching the row_name
        for row in table_data:
            # if row and row[0] and isinstance(row[0], str) and row[0].strip() == row_name:
            # Instead of exact match
            if row and row[0] and isinstance(row[0], str) and row[0].strip() == row_name.strip():
                row_values = row[1:]  # Exclude the first column (row name)
                break
        
        if row_values is None:
            raise ValueError(f"Row '{row_name}' not found in table '{table_name}'")
        
        # Calculate sum of numeric values
        total = 0.0
        for value in row_values:
            if isinstance(value, (int, float)):
                total += value
            elif isinstance(value, str):
                # Try to extract numeric value from strings like "10%"
                numeric_value = self._extract_numeric(value)
                if numeric_value is not None:
                    total += numeric_value
        
        return total

    def _extract_numeric(self, value: str) -> Optional[float]:
        """Extract numeric value from string (handles percentages, etc.)."""
        # Remove non-numeric characters (except . and -)
        numeric_str = re.sub(r"[^\d.-]", "", value)
        if numeric_str:
            try:
                return float(numeric_str)
            except ValueError:
                return None
        return None