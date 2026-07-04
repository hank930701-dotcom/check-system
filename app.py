from flask import Flask, render_template, request
import os
from utils_excel import save_to_excel

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


@app.route("/")
def form():
    return render_template("form.html")


@app.route("/submit", methods=["POST"])
def submit():
    data = request.form.to_dict()
    files = request.files

    img_paths = {}

    img_fields = [
        "soda_nozzle","soda_valve","inventory_diff","posting_screen",
        "loss_screen","k_floor","batter_table","fryer",
        "fryer_bottom","floor2","floor3","trash","toilet"
    ]

    for field in img_fields:
        file = files.get(field)
        if file and file.filename != "":
            path = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(path)
            img_paths[field] = path

    save_to_excel(data, img_paths)

    return "✅ 已提交成功，Excel已更新"


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
