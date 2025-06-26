from flask import Flask, jsonify, request

app = Flask(__name__)
cached_data = []

@app.route("/api/ping", methods=["GET"])
def ping():
    return jsonify({"message": "API aktif dan bisa diakses!"})

@app.route("/api/submit", methods=["POST"])
def submit_data():
    data = request.json
    cached_data.clear()
    cached_data.extend(data)
    return jsonify({"message": "Data received and cached!"})

@app.route("/api/preview", methods=["GET"])
def preview_data():
    return jsonify(cached_data)

with app.test_request_context():
    print(app.url_map)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
