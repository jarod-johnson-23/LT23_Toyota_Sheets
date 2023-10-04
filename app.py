from flask import Flask, request, send_file
import pandas as pd
import io

app = Flask(__name__)


@app.route("/merge-excel", methods=["POST"])
def merge_excel():
    # Check if 'files' is part of the request
    if "files" not in request.files:
        return "No files part in the request."

    files = request.files.getlist("files")

    # List to store dataframes
    all_data = []

    for file in files:
        # Convert each file content into a DataFrame and append to all_data
        df = pd.read_excel(file)
        all_data.append(df)

    # Concatenate all dataframes
    merged_data = pd.concat(all_data, axis=0, ignore_index=True)

    # Save merged data to Excel format
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine="openpyxl")
    merged_data.to_excel(writer, sheet_name="Merged", index=False)
    writer.save()

    # Return merged excel
    output.seek(0)
    return send_file(
        output,
        as_attachment=True,
        download_name="merged.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(debug=True)
