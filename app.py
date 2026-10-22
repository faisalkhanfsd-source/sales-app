import os, hmac, functools
from flask import Flask, request, session, redirect, url_for, render_template, send_file, flash
import processor

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "change-this-secret-key")
app.config["MAX_CONTENT_LENGTH"] = 60 * 1024 * 1024
USER = os.environ.get("APP_USER", "admin")
PASSWORD = os.environ.get("APP_PASSWORD", "admin123")   # ہوسٹ کرتے وقت ضرور بدلیں


def login_required(f):
    @functools.wraps(f)
    def w(*a, **k):
        if not session.get("ok"):
            return redirect(url_for("login"))
        return f(*a, **k)
    return w


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        u, p = request.form.get("username", ""), request.form.get("password", "")
        if hmac.compare_digest(u, USER) and hmac.compare_digest(p, PASSWORD):
            session["ok"] = True
            return redirect(url_for("index"))
        flash("یوزر نیم یا پاسورڈ غلط ہے")
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/", methods=["GET", "POST"])
@login_required
def index():
    if request.method == "POST":
        raw, fmt = request.files.get("raw"), request.files.get("format")
        if not raw or not fmt or not raw.filename or not fmt.filename:
            flash("دونوں فائلیں اپلوڈ کریں")
            return redirect(url_for("index"))
        try:
            out, _ = processor.build(raw.stream, fmt.stream)
        except Exception as e:
            flash("فائل پروسیس نہیں ہو سکی: %s" % e)
            return redirect(url_for("index"))
        return send_file(out, as_attachment=True, download_name="File_Format_Result.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    return render_template("index.html")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
