
from fastapi import FastAPI
from openpyxl import load_workbook
from pathlib import Path

app = FastAPI(
    title="Zoo Visitors API",
    description="Zoo visitor statistics for fiscal year 2569",
    version="1.0.0"
)

EXCEL_FILE = Path(__file__).parent / "data" / "zoo_visitors_sep_2569.xlsx"


@app.get("/")
def home():
    return {
        "message": "Welcome to Zoo Visitors API",
        "docs": "/docs"
    }


@app.get("/api/zoo-visitors")
def get_zoo_visitors():
    workbook = load_workbook(
        EXCEL_FILE,
        read_only=True,
        data_only=True
    )

    try:
        sheet = workbook["statistics"]
        rows = sheet.iter_rows(values_only=True)

        headers = next(rows)
        headers = [
            str(header).strip()
            for header in headers
            if header is not None
        ]

        visitors = []

        for row in rows:
            values = row[:len(headers)]

            if not any(value is not None for value in values):
                continue

            record = dict(zip(headers, values))
            visitors.append(record)

    finally:
        workbook.close()

    return {
        "source": "Data.go.th - Zoological Park Organization of Thailand",
        "dataset": "zoo-visitors2569",
        "resource_id": "0ecc96cf-d914-42a1-b184-dbbd013c3c9c",
        "fiscal_year": 2569,
        "sheet": "statistics",
        "total_records": len(visitors),
        "data": visitors
    }
