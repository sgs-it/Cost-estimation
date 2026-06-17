"""
quote_api.py — Flask REST API for real-time quote estimation + Excel export
Run:  python quote_api.py
API:  POST /estimate        body: { job_type, config }
      POST /export/excel    body: { job_type, config, meta }  → returns .xlsx file
      GET  /ping
"""
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import traceback, io, re, datetime
from price_logic import get_estimate
from excel_exporter import generate_excel

app = Flask(__name__)
CORS(app, expose_headers=["Content-Disposition"])  # allow browser requests and reading the attachment filename

@app.route("/ping")
def ping():
    return jsonify({"status": "ok"})

@app.route("/estimate", methods=["POST"])
def estimate():
    try:
        data = request.get_json(force=True)
        job_type = data.get("job_type")
        config   = data.get("config", {})
        result   = get_estimate(job_type, config)
        return jsonify({"ok": True, "result": result})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e), "trace": traceback.format_exc()}), 400

@app.route("/export/excel", methods=["POST"])
def export_excel():
    try:
        data     = request.get_json(force=True)
        job_type = data.get("job_type")
        config   = data.get("config", {})
        meta     = data.get("meta", {})
        result   = get_estimate(job_type, config)

        xlsx_bytes = generate_excel(job_type, result, config, meta)

        # Build filename matching reference naming convention
        today  = datetime.date.today()
        prefix_map = {
            "amc_outdoor":    "AMC - OM",
            "amc_indoor":     "AMC - IOM",
            "project_outdoor":"PTO",
            "project_indoor": "PTI",
        }
        prefix  = prefix_map.get(job_type, "QUOTE")
        site    = re.sub(r'[^\w\s-]', '', meta.get("site", "Site")).strip().replace(" ", "_")
        client  = re.sub(r'[^\w\s-]', '', meta.get("client","Client")).strip().replace(" ","_")
        fname   = f"{prefix} - {today.strftime('%Y-%m-%d')} - {client} - {site}.xlsx"

        return send_file(
            io.BytesIO(xlsx_bytes),
            as_attachment=True,
            download_name=fname,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        return jsonify({"ok": False, "error": str(e), "trace": traceback.format_exc()}), 400

if __name__ == "__main__":
    print("Quote Estimation API running -> http://localhost:5050")
    app.run(host="0.0.0.0", port=5050, debug=False)
