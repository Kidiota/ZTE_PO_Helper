
from flask import Flask, render_template

app = Flask(__name__)

# 目录按钮配置：后续只需在这里添加更多按钮即可
LINKS = [
    {"icon": "fa-solid fa-screwdriver-wrench", "text_zh": "PO 信息批量提取", "text_en": "Get PO info", "href": "http://10.32.155.150:3000"},
    {"icon": "fa-solid fa-file-pen", "text_zh": "文档规范化自动命名", "text_en": "Document Renaming Tool", "href": "http://10.32.155.150:5000"},
    {"icon": "fa-solid fa-database", "text_zh": "开票文档批量准备工具(下载)", "text_en": "Invoicing document batch preparation tool", "href": "kayee_2.zip"},
]

@app.route("/")
def index():
    return render_template("index.html", links=LINKS)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5500, debug=True)
