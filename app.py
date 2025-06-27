from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import json
import os
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from io import BytesIO

app = Flask(__name__)
CORS(app)

# ------------------ Admin Login ------------------
@app.route("/login", methods=["POST"])
def login():
    creds = request.json
    username = creds.get("username")
    password = creds.get("password")

    if username == "vinay0703" and password == "vinay1234":
        return jsonify({"status": "success"})
    else:
        return jsonify({"status": "error", "message": "Invalid credentials"}), 401

# ------------------ Submit Entry ------------------
@app.route("/submit", methods=["POST"])
def submit():
    data = request.json
    month = request.args.get("month")
    filename = f"data_{month}.json"

    if os.path.exists(filename):
        with open(filename, "r") as f:
            existing = json.load(f)
    else:
        existing = []

    existing.append(data)

    with open(filename, "w") as f:
        json.dump(existing, f, indent=2)

    return jsonify({"status": "success"})

# ------------------ Get Data ------------------
@app.route("/data", methods=["GET"])
def get_data():
    month = request.args.get("month")
    filename = f"data_{month}.json"

    if os.path.exists(filename):
        with open(filename, "r") as f:
            return jsonify(json.load(f))
    else:
        return jsonify([])

# ------------------ Edit Entry ------------------
@app.route("/edit", methods=["POST"])
def edit_entry():
    month = request.args.get("month")
    index = int(request.args.get("index"))
    new_data = request.json
    filename = f"data_{month}.json"

    if not os.path.exists(filename):
        return jsonify({"error": "File not found"}), 404

    with open(filename, "r") as f:
        data = json.load(f)

    if 0 <= index < len(data):
        data[index] = new_data
        with open(filename, "w") as f:
            json.dump(data, f, indent=2)
        return jsonify({"status": "updated"})
    else:
        return jsonify({"error": "Invalid index"}), 400

# ------------------ Delete Entry ------------------
@app.route("/delete", methods=["POST"])
def delete_entry():
    month = request.args.get("month")
    index = int(request.args.get("index"))
    filename = f"data_{month}.json"

    if not os.path.exists(filename):
        return jsonify({"error": "File not found"}), 404

    with open(filename, "r") as f:
        data = json.load(f)

    if 0 <= index < len(data):
        data.pop(index)
        with open(filename, "w") as f:
            json.dump(data, f, indent=2)
        return jsonify({"status": "deleted"})
    else:
        return jsonify({"error": "Invalid index"}), 400

# ------------------ Excel Download ------------------
@app.route("/download/<month>", methods=["GET"])
def download(month):
    filename = f"data_{month}.json"
    if not os.path.exists(filename):
        return jsonify([])

    with open(filename, 'r') as f:
        data = json.load(f)

    wb = Workbook()
    ws = wb.active
    ws.title = f"Data_{month}"

    if not data:
        ws.append(["No data available"])
    else:
        headers = list(data[0].keys())
        ws.append(headers)

        # Style headers
        for col in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center')

        # Append data
        for row in data:
            ws.append([row.get(col, "") for col in headers])

        # Auto-adjust column widths
        for col in ws.columns:
            max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
            adjusted_width = max_length + 2
            ws.column_dimensions[col[0].column_letter].width = adjusted_width

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name=f"data_{month}.xlsx",
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

# ------------------ Run Server (Render-compatible) ------------------
if __name__ == "__main__":
    print("✅ Flask server is running...")
    app.run(host="0.0.0.0", port=10000)
