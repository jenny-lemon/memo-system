from flask import Flask, request, render_template_string
import requests
from memo import login, process_order

app = Flask(__name__)


HTML = """
<h2>Memo 系統</h2>

<form method="post">
Email: <input name="email"><br>
Password: <input name="password"><br><br>

訂單編號: <input name="order"><br><br>

<button type="submit">🚀 執行</button>
</form>

<hr>

<h3>執行過程</h3>
<pre>{{log}}</pre>

<h3>結果</h3>
{{result}}
"""


@app.route("/", methods=["GET", "POST"])
def index():

    log_lines = []
    def log(x):
        log_lines.append(x)

    result = ""

    if request.method == "POST":

        session = requests.Session()

        email = request.form["email"]
        password = request.form["password"]
        order = request.form["order"]

        login(session, email, password)

        log("[登入] 成功")

        count, ok = process_order(session, order, log)

        if ok:
            result = f"✅ 成功：回寫 {count} 筆"
        else:
            result = "❌ 失敗"

    return render_template_string(HTML,
                                  log="\n".join(log_lines),
                                  result=result)


if __name__ == "__main__":
    app.run(debug=True)
