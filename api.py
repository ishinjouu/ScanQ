from flask import Flask, jsonify, request, Response
from collections import OrderedDict
import json
from app import load_processor_module
from utils import sanitize_for_json

app = Flask(__name__)
cached_data = []

# PING
@app.route("/api/ping", methods=["GET"])
def ping():
    return jsonify({"message": "API aktif dan bisa diakses!"})


# SUBMIT FILE (utama)
@app.route("/api/submit", methods=["POST"])
def submit_data():
    if 'file' not in request.files:
        return jsonify({"error": "❌ File tidak ditemukan di form-data"}), 400

    file = request.files['file']

    try:
        extract_table_from_pdf, bersihkan_dataframe, transform_to_final_format = load_processor_module(file.filename)

        df_raw = extract_table_from_pdf(file)
        df_clean = bersihkan_dataframe(df_raw)
        df_final = transform_to_final_format(df_clean)
        result = df_final.to_dict(orient="records")
        safe_records = [sanitize_for_json(row) for row in result]
        cached_data.clear()
        cached_data.extend(safe_records)

        return jsonify({"message": "✅ File berhasil diproses", "data": safe_records})
    
    except Exception as e:
        return jsonify({"error": f"🔥 Gagal proses file: {str(e)}"}), 500

# PREVIEW
@app.route("/api/preview", methods=["GET"])
def preview_data():
    if not cached_data:
        return jsonify({"message": "Belum ada data yang dikirim"}), 200
    
    ordered = []
    for item in cached_data:
        ordered.append(OrderedDict([
            ("section", item.get("section")),
            ("point_check", item.get("point_check")),
            ("jenis_point", item.get("jenis_point")),
            ("catatan", item.get("catatan")),
            ("item_check", item.get("item_check")),
            ("control_method", item.get("control_method")),
            ("standard", item.get("standard")),
            ("jenis_pengecekan", item.get("jenis_pengecekan")),
            ("std_value", item.get("std_value")),
            ("std_min", item.get("std_min")),
            ("std_max", item.get("std_max"))
        ]))

    return Response(json.dumps(ordered, ensure_ascii=False), mimetype='application/json')

# REGISTERED ROUTES
with app.test_request_context():
    print(app.url_map)

# RUN FLASK
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=2051, threaded=True)
