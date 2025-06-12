# Capital Budgeting Excel Processor API

This project provides a FastAPI-based web API for processing and analyzing capital budgeting data from an Excel file. It is designed to help users extract, inspect, and compute values from structured financial spreadsheets, specifically tailored for capital budgeting workflows.

## Features
- **List all tables** in the Excel sheet
- **Get row details** for a specific table
- **Calculate the sum** of numerical values in a specific row
- Robust error handling for missing or malformed data
- Interactive API documentation via Swagger UI (`/docs`)

## Requirements
- Python 3.7+
- [FastAPI](https://fastapi.tiangolo.com/)
- [Uvicorn](https://www.uvicorn.org/)
- [Pandas](https://pandas.pydata.org/)
- [pydantic](https://pydantic-docs.helpmanual.io/)

Install dependencies:
```bash
pip install fastapi uvicorn pandas python-multipart
```

## Setup
1. Place your capital budgeting Excel file as `capbudg.xls` in the same directory as `main.py`.
   - The code expects a sheet named `CapBudgWS`.
2. Save the provided code as `main.py` if not already present.

## Running the API
Start the FastAPI server with:
```bash
uvicorn main:app --reload --port 9090
```

## API Endpoints

### 1. List all tables
- **GET** `/list_tables`
- **Response:** `{ "tables": ["INITIAL INVESTMENT", ...] }`

### 2. Get table details
- **GET** `/get_table_details?table_name=INITIAL%20INVESTMENT`
- **Response:** `{ "table_name": "INITIAL INVESTMENT", "row_names": [ ... ] }`

### 3. Calculate row sum
- **GET** `/row_sum?table_name=INITIAL%20INVESTMENT&row_name=Initial%20Investment=`
- **Response:** `{ "table_name": "...", "row_name": "...", "sum": ... }`

### 4. Root endpoint
- **GET** `/`
- **Response:** `{ "message": "Capital Budgeting Excel Processor API - Use /docs for API documentation" }`

## Testing the API
You can use [Postman](https://www.postman.com/) or any HTTP client to test the endpoints. Example requests:
- List tables: `GET http://localhost:9090/list_tables`
- Get table details: `GET http://localhost:9090/get_table_details?table_name=INITIAL%20INVESTMENT`
- Row sum: `GET http://localhost:9090/row_sum?table_name=INITIAL%20INVESTMENT&row_name=Initial%20Investment=`

## Notes & Limitations
- The Excel file must have a specific structure and sheet name (`CapBudgWS`).
- Not thread-safe for concurrent requests.
- No authentication or file upload support by default.
- Handles only `.xls` files and may not support `.xlsx` or `.csv` out of the box.

## Potential Improvements
- Support for more Excel formats (xlsx, csv)
- API key authentication
- Caching for performance
- Bulk processing endpoints
- File upload via API
- Improved handling for empty or malformed files
- Thread safety for concurrent access

## License
MIT or as specified by the project owner. 