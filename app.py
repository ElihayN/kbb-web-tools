# -*- coding: utf-8 -*-
import os
import uuid
import secrets
import traceback
from functools import wraps

from flask import (
    Flask, render_template, request, redirect, url_for, send_from_directory, flash, Response
)
from werkzeug.utils import secure_filename

from logic import weekly_sms, kbb_packing_list, kbb_invoice
from logic.office import OfficeError

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_ROOT = os.path.join(BASE_DIR, "uploads")
OUTPUT_ROOT = os.path.join(BASE_DIR, "outputs")
ASSETS_DIR = os.path.join(BASE_DIR, "assets")
DEFAULT_INVOICE_TEMPLATE = os.path.join(ASSETS_DIR, "kbb_invoice_template.xlsx")

os.makedirs(UPLOAD_ROOT, exist_ok=True)
os.makedirs(OUTPUT_ROOT, exist_ok=True)
os.makedirs(ASSETS_DIR, exist_ok=True)

app = Flask(__name__)
app.secret_key = "kbb-internal-tools"  # internal LAN tool, not internet-facing
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 200MB

# --------------------------------------------------------------- basic auth
# This app handles customer phone numbers and customs invoices. The moment
# it's reachable over the open internet (even via a "hidden" temporary
# tunnel URL), anyone who finds/guesses the URL can see that data unless we
# gate it. AUTH_USER / AUTH_PASS can be overridden via environment
# variables for a real deployment; a random password is generated per
# process start otherwise so there's never a silent no-auth mode.
AUTH_USER = os.environ.get("KBB_AUTH_USER", "kbb")
AUTH_PASS = os.environ.get("KBB_AUTH_PASS") or secrets.token_urlsafe(9)
if not os.environ.get("KBB_AUTH_PASS"):
    print(f"[auth] no KBB_AUTH_PASS set — generated one for this run: {AUTH_PASS}")


def _check_auth(username, password):
    return secrets.compare_digest(username, AUTH_USER) and secrets.compare_digest(password, AUTH_PASS)


@app.before_request
def _require_auth():
    auth = request.authorization
    if not auth or not _check_auth(auth.username, auth.password):
        return Response(
            "Authentication required.", 401,
            {"WWW-Authenticate": 'Basic realm="KBB Tools"'},
        )


def new_run_dir():
    run_id = uuid.uuid4().hex[:12]
    d = os.path.join(OUTPUT_ROOT, run_id)
    os.makedirs(d, exist_ok=True)
    return run_id, d


def save_uploads(files, run_dir):
    paths = []
    for f in files:
        if not f or not f.filename:
            continue
        name = secure_filename(f.filename)
        path = os.path.join(run_dir, name)
        f.save(path)
        paths.append(path)
    return paths


@app.route("/")
def index():
    return render_template("index.html")


# --------------------------------------------------------------------- SMS
@app.route("/weekly-sms", methods=["GET", "POST"])
def weekly_sms_page():
    if request.method == "GET":
        return render_template("weekly_sms.html")

    upload = request.files.get("report_file")
    if not upload or not upload.filename:
        flash("יש להעלות קובץ דוח Operational Business.", "err")
        return render_template("weekly_sms.html")

    run_id, run_dir = new_run_dir()
    try:
        paths = save_uploads([upload], run_dir)
        result = weekly_sms.process_weekly_report(paths[0], run_dir)
        result["sms_name"] = os.path.basename(result["sms_path"])
        result["app2u_name"] = os.path.basename(result["app2u_path"])
        result["run_id"] = run_id
        return render_template("weekly_sms.html", result=result)
    except (weekly_sms.WeeklySmsError, OfficeError) as e:
        flash(str(e), "err")
        return render_template("weekly_sms.html")
    except Exception:
        flash("שגיאה לא צפויה: " + traceback.format_exc(limit=3), "err")
        return render_template("weekly_sms.html")


# ------------------------------------------------------------- packing list
@app.route("/packing-list", methods=["GET", "POST"])
def packing_list_page():
    if request.method == "GET":
        return render_template("packing_list.html")

    action = request.form.get("action", "consolidate")

    if action == "export_pdf":
        upload = request.files.get("xlsx_file")
        if not upload or not upload.filename:
            flash("יש להעלות את קובץ ה-Excel המאושר.", "err")
            return render_template("packing_list.html")
        run_id, run_dir = new_run_dir()
        try:
            paths = save_uploads([upload], run_dir)
            pdf_path = kbb_packing_list.export_pdf(paths[0], run_dir)
            result = {"pdf_name": os.path.basename(pdf_path), "run_id": run_id, "pdf_only": True}
            return render_template("packing_list.html", result=result)
        except OfficeError as e:
            flash(str(e), "err")
            return render_template("packing_list.html")
        except Exception:
            flash("שגיאה לא צפויה: " + traceback.format_exc(limit=3), "err")
            return render_template("packing_list.html")

    files = request.files.getlist("pdf_files")
    files = [f for f in files if f and f.filename]
    if not files:
        flash("יש להעלות לפחות קובץ Packing List אחד (PDF).", "err")
        return render_template("packing_list.html")

    run_id, run_dir = new_run_dir()
    try:
        paths = save_uploads(files, run_dir)
        result = kbb_packing_list.process(paths, run_dir)
        result["xlsx_name"] = os.path.basename(result["xlsx_path"])
        result["run_id"] = run_id
        return render_template("packing_list.html", result=result)
    except kbb_packing_list.PackingListError as e:
        flash(str(e), "err")
        return render_template("packing_list.html")
    except OfficeError as e:
        flash(str(e), "err")
        return render_template("packing_list.html")
    except Exception:
        flash("שגיאה לא צפויה: " + traceback.format_exc(limit=3), "err")
        return render_template("packing_list.html")


# ------------------------------------------------------------------ invoice
@app.route("/invoice", methods=["GET", "POST"])
def invoice_page():
    if request.method == "GET":
        return render_template("invoice.html")

    action = request.form.get("action", "consolidate")

    if action == "export_pdf":
        upload = request.files.get("xlsx_file")
        if not upload or not upload.filename:
            flash("יש להעלות את קובץ ה-Excel המאושר.", "err")
            return render_template("invoice.html")
        run_id, run_dir = new_run_dir()
        try:
            paths = save_uploads([upload], run_dir)
            pdf_path = kbb_invoice.export_pdf(paths[0], run_dir)
            result = {"pdf_name": os.path.basename(pdf_path), "run_id": run_id, "pdf_only": True}
            return render_template("invoice.html", result=result)
        except OfficeError as e:
            flash(str(e), "err")
            return render_template("invoice.html")
        except Exception:
            flash("שגיאה לא צפויה: " + traceback.format_exc(limit=3), "err")
            return render_template("invoice.html")

    files = request.files.getlist("pdf_files")
    files = [f for f in files if f and f.filename]
    if not files:
        flash("יש להעלות לפחות קובץ חשבונית Proforma אחד (PDF).", "err")
        return render_template("invoice.html")

    dims = (request.form.get("dims") or "").strip()
    if not dims:
        flash("שדה המידות (DIMS) אינו ניתן לחילוץ מהחשבוניות — יש למלא אותו ידנית.", "err")
        return render_template("invoice.html")

    run_id, run_dir = new_run_dir()
    try:
        paths = save_uploads(files, run_dir)

        template_upload = request.files.get("template_file")
        if template_upload and template_upload.filename:
            template_path = save_uploads([template_upload], run_dir)[0]
        else:
            template_path = DEFAULT_INVOICE_TEMPLATE

        overrides = {}
        for field in ("reference_no", "gross_weight", "defective_return", "ship_date"):
            val = (request.form.get(field) or "").strip()
            if val:
                overrides[field] = val

        result = kbb_invoice.process(paths, template_path, run_dir, dims=dims, **overrides)
        result["xlsx_name"] = os.path.basename(result["xlsx_path"])
        result["run_id"] = run_id
        return render_template("invoice.html", result=result)
    except kbb_invoice.InvoiceError as e:
        flash(str(e), "err")
        return render_template("invoice.html")
    except OfficeError as e:
        flash(str(e), "err")
        return render_template("invoice.html")
    except Exception:
        flash("שגיאה לא צפויה: " + traceback.format_exc(limit=3), "err")
        return render_template("invoice.html")


# ------------------------------------------------------------------ downloads
@app.route("/download/<run_id>/<path:filename>")
def download(run_id, filename):
    run_id = secure_filename(run_id)
    directory = os.path.join(OUTPUT_ROOT, run_id)
    return send_from_directory(directory, filename, as_attachment=True)


if __name__ == "__main__":
    # 0.0.0.0 so other computers on the local network can reach it too,
    # e.g. http://<this-computer's-LAN-IP>:5000 — and PORT is honored so
    # this same entrypoint works unchanged on Render/Railway/etc, which
    # assign the port at runtime via that env var.
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
