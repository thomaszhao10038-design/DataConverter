import pandas as pd
from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import StreamingResponse
from io import BytesIO
from openpyxl.writer.excel import save_virtual_workbook

app = FastAPI()


def round_to_10min(ts):
    if pd.isna(ts):
        return ts

    ts = pd.to_datetime(ts)
    minutes = ts.minute
    remainder = minutes % 10

    if remainder < 5:
        rounded = minutes - remainder
    else:
        rounded = minutes + (10 - remainder)

    if rounded == 60:
        ts = ts.replace(minute=0) + pd.Timedelta(hours=1)
    else:
        ts = ts.replace(minute=rounded)

    return ts.replace(second=0, microsecond=0)


@app.post("/convert")
async def convert_file(
    file: UploadFile = File(...),
    timestamp_column: str = Form("Timestamp")
):
    try:
        # Read uploaded Excel file
        df = pd.read_excel(file.file)

        if timestamp_column not in df.columns:
            return {"error": f"Column '{timestamp_column}' not found."}

        # Apply rounding
        df["Rounded_Time"] = df[timestamp_column].apply(round_to_10min)

        # Convert to output Excel
        output_stream = BytesIO()
        excel_data = save_virtual_workbook(df.to_excel(index=False))
        output_stream.write(excel_data)
        output_stream.seek(0)

        return StreamingResponse(
            output_stream,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=converted.xlsx"}
        )

    except Exception as e:
        return {"error": str(e)}


@app.get("/")
def home():
    return {"status": "Data Converter API is running!"}
