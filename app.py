import os, hmac, functools
from flask import Flask, request, session, redirect, url_for, render_template, send_file, flash
import processor, ncb, base64, io

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


@app.route("/ncb", methods=["GET", "POST"])
@login_required
def ncb_page():
    if request.method == "GET":
        return render_template("ncb.html", img=None)
    f = request.files.get("file")
    region = request.form.get("region", "SBL TOTAL")
    if not f or not f.filename:
        flash("فائل اپلوڈ کریں")
        return redirect(url_for("ncb_page"))
    try:
        data = ncb.read_summary(f.stream)
        png = ncb.render(data, region)
    except Exception as e:
        flash("فائل پروسیس نہیں ہو سکی: %s" % e)
        return redirect(url_for("ncb_page"))
    note = None
    if request.form.get("send"):
        try:
            note = "واٹس ایپ پر بھیج دی گئی" if ncb.send_whatsapp(png) else \
                   "واٹس ایپ کی سیٹنگ (WA_TOKEN, WA_PHONE_ID, WA_TO) نہیں لگی"
        except Exception as e:
            note = "واٹس ایپ پر بھیجنے میں ایرر: %s" % e
    return render_template("ncb.html", region=region, regions=list(data["regions"]),
                           note=note, img=base64.b64encode(png).decode())


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
