from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import JSONResponse
import pandas as pd
from typing import Dict, List
import os
from pydantic import BaseModel

app = FastAPI(
    title="Capital Budgeting Excel Processor API",
    description="API for processing capital budgeting Excel sheets",
    version="1.0.0"
)

# Global variable to store the Excel data
excel_data = None

class TableResponse(BaseModel):
    tables: List[str]

class TableDetailsResponse(BaseModel):
    table_name: str
    row_names: List[str]

class RowSumResponse(BaseModel):
    table_name: str
    row_name: str
    sum: float

def load_excel_data(file_path: str = "capbudg.xls"):
    """Load and preprocess the Excel data"""
    global excel_data
    
    try:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Excel file not found at {file_path}")
            
        # Read the Excel file
        df = pd.read_excel(file_path, sheet_name="CapBudgWS", header=None)
        excel_data = df.where(pd.notnull(df), None)
        return True
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error loading Excel file: {str(e)}")

@app.on_event("startup")
async def startup_event():
    """Load the Excel data when the application starts"""
    try:
        load_excel_data()
    except Exception as e:
        raise RuntimeError(f"Failed to load Excel data on startup: {str(e)}")

@app.get("/list_tables", response_model=TableResponse)
async def list_tables():
    """
    List all tables in the Excel sheet.
    
    Returns:
        TableResponse: A dictionary with a list of table names
    """
    if excel_data is None:
        raise HTTPException(status_code=500, detail="Excel data not loaded")
    
    # Identify tables by their headers in column A
    tables = []
    for idx, row in excel_data.iterrows():
        if row[0] and isinstance(row[0], str) and row[0].strip() in [
            "INITIAL INVESTMENT",
            "CASHFLOW DETAILS",
            "DISCOUNT RATE",
            "WORKING CAPITAL",
            "GROWTH RATES",
            "SALVAGE VALUE",
            "OPERATING CASHFLOWS",
            "BOOK VALUE & DEPRECIATION"
        ]:
            tables.append(row[0].strip())
    
    return {"tables": tables}

@app.get("/get_table_details", response_model=TableDetailsResponse)
async def get_table_details(table_name: str = Query(..., description="Name of the table to inspect")):
    """
    Get the row names for a specific table.
    
    Args:
        table_name (str): Name of the table to inspect
        
    Returns:
        TableDetailsResponse: Contains table name and list of row names
    """
    if excel_data is None:
        raise HTTPException(status_code=500, detail="Excel data not loaded")
    
    # Find the table header row
    header_row = None
    for idx, row in excel_data.iterrows():
        if row[0] and isinstance(row[0], str) and row[0].strip() == table_name:
            header_row = idx
            break
    
    if header_row is None:
        raise HTTPException(status_code=404, detail=f"Table '{table_name}' not found")
    
    # Collect row names until we hit an empty row or next table
    row_names = []
    current_row = header_row + 1
    
    while current_row < len(excel_data):
        row_name = excel_data.iloc[current_row, 0]
        
        # Stop if we hit an empty row or next table
        if row_name is None or not str(row_name).strip():
            break
        if isinstance(row_name, str) and row_name.strip() in [
            "INITIAL INVESTMENT",
            "CASHFLOW DETAILS",
            "DISCOUNT RATE",
            "WORKING CAPITAL",
            "GROWTH RATES",
            "SALVAGE VALUE",
            "OPERATING CASHFLOWS",
            "BOOK VALUE & DEPRECIATION"
        ]:
            break
            
        row_names.append(str(row_name).strip())
        current_row += 1
    
    return {
        "table_name": table_name,
        "row_names": row_names
    }

@app.get("/row_sum", response_model=RowSumResponse)
async def calculate_row_sum(
    table_name: str = Query(..., description="Name of the table"),
    row_name: str = Query(..., description="Name of the row to sum")
):
    """
    Calculate the sum of numerical values in a specific row.
    
    Args:
        table_name (str): Name of the table
        row_name (str): Name of the row to sum
        
    Returns:
        RowSumResponse: Contains table name, row name, and sum of values
    """
    if excel_data is None:
        raise HTTPException(status_code=500, detail="Excel data not loaded")
    
    # First get the row names to verify the requested row exists
    table_details = await get_table_details(table_name)
    if row_name not in table_details["row_names"]:
        raise HTTPException(status_code=404, detail=f"Row '{row_name}' not found in table '{table_name}'")
    
    # Find the specific row
    target_row = None
    for idx, row in excel_data.iterrows():
        current_row_name = str(row[0]).strip() if row[0] else None
        if current_row_name == row_name:
            target_row = idx
            break
    
    if target_row is None:
        raise HTTPException(status_code=404, detail=f"Row '{row_name}' not found in Excel data")
    
    # Determine which columns to check based on table
    start_col = 0
    end_col = len(excel_data.columns)
    
    if table_name == "INITIAL INVESTMENT":
        value_col = 2  # Column C
        try:
            value = excel_data.iloc[target_row, value_col]
            if pd.isna(value):
                sum_value = 0
            elif isinstance(value, str) and '%' in value:
                sum_value = float(value.strip('%')) / 100
            else:
                sum_value = float(value)
        except (ValueError, TypeError):
            sum_value = 0
    else:
        # For other tables, sum all numeric values in the row
        sum_value = 0
        for col in range(start_col, end_col):
            try:
                value = excel_data.iloc[target_row, col]
                if pd.notna(value):
                    if isinstance(value, str) and '%' in value:
                        sum_value += float(value.strip('%')) / 100
                    else:
                        sum_value += float(value)
            except (ValueError, TypeError):
                continue
    
    return {
        "table_name": table_name,
        "row_name": row_name,
        "sum": sum_value
    }

@app.get("/")
async def root():
    return {"message": "Capital Budgeting Excel Processor API - Use /docs for API documentation"}