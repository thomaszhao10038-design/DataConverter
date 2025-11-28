import pandas as pd
from flask import Flask, request, send_file, jsonify
from io import BytesIO
from openpyxl.writer.excel import save_virtual_workbook

app = Flask(__name__)

def round_to_10min(ts):
    """
    Round a pandas timestamp object to nearest 10 minutes.
    """
    if pd.isna(ts):
        return ts
    # Convert to pandas Timestamp
    ts = pd.to_datetime(ts)
    # Seconds since the hour
    minutes = ts.minute
    remainder = minutes % 10

    if remainder < 5:
        rounded_min = minutes - remainder
    else:
        rounded_min = minutes + (10 - remainder)

    # Adjust hour if >= 60
    if rounded_min == 60:
        ts = ts.replace(minute=0) + pd.Timedelta(hours=1)
    else:
        ts = ts.replace(minute=rounded_min)

    return ts.replace(second=0, microsecond=0)

@app.route("/convert", methods=["POST"])
def convert_file():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]
    ts_col = request.form.get("timestamp_column", "Timestamp")

    try:
        df = pd.read_excel(file)

        if ts_col not in df.columns:
            return jsonify({"error": f"Column '{ts_col}' not found"}), 400

        # Apply rounding
        df["Rounded_Time"] = df[ts_col].apply(round_to_10min)

        # Save to memory
        output = BytesIO()
        excel_data = save_virtual_workbook(df.to_excel(index=False))
        output.write(excel_data)
        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name="converted.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/")
def index():
    return "Data Converter API is running."


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000)
