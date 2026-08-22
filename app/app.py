#!/usr/bin/env python3
"""
Uganda Secondary School Report Card System – Web App
Accessible browser interface for generating CBC-aligned report cards.
"""

import os
import sys
import zipfile
import shutil
from datetime import datetime
from pathlib import Path
from io import BytesIO

# Ensure project root is on the path (works locally and on cloud hosts)
ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from flask import (
    Flask, render_template, request, redirect, url_for,
    flash, send_file, send_from_directory, jsonify
)
import pandas as pd
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, logout_user, login_required, current_user
from sqlalchemy import text

# Import core logic from the existing system
from report_card_system import (
    create_excel_template,
    generate_all_reports,
    load_data,
    calculate_final_score,
    get_grade,
    SCHOOL_CONFIG,
)

app = Flask(__name__)

# Security
app.secret_key = os.environ.get(
    "SECRET_KEY_BASE",
    os.environ.get("SECRET_KEY", "uganda-report-card-secret-key-change-in-production")
)

# Database
database_url = os.environ.get("DATABASE_URL")

if database_url and database_url.startswith("postgres://"):
    database_url = database_url.replace("postgres://", "postgresql://", 1)

app.config["SQLALCHEMY_DATABASE_URI"] = database_url
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

# Upload limit: 16 MB
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024
# Initialize database and login manager
db = SQLAlchemy(app)
login_manager = LoginManager(app)
login_manager.login_view = "login"
@login_manager.user_loader
def load_user(user_id):
        return db.session.get(User, int(user_id))
# ==============================
# ADMINISTRATION MONITORING MODELS
# ==============================

class School(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), nullable=False)
    district = db.Column(db.String(100))
    active = db.Column(db.Boolean, default=True)

    users = db.relationship("User", backref="school", lazy=True)
    subscriptions = db.relationship("Subscription", backref="school", lazy=True)


class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    full_name = db.Column(db.String(150), nullable=False)
    username = db.Column(db.String(100), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(30), nullable=False, default="teacher")

    school_id = db.Column(db.Integer, db.ForeignKey("school.id"), nullable=False)

    active = db.Column(db.Boolean, default=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    mark_entries = db.relationship("MarkEntry", backref="teacher", lazy=True)


class MarkEntry(db.Model):
    id = db.Column(db.Integer, primary_key=True)

    teacher_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    student_name = db.Column(db.String(150), nullable=False)
    class_name = db.Column(db.String(50), nullable=False)
    subject = db.Column(db.String(100), nullable=False)

    mark = db.Column(db.Float)
    status = db.Column(db.String(30), default="draft")

    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(
        db.DateTime,
        default=datetime.utcnow,
        onupdate=datetime.utcnow
    )


class Subscription(db.Model):
    id = db.Column(db.Integer, primary_key=True)

    school_id = db.Column(db.Integer, db.ForeignKey("school.id"), nullable=False)

    amount = db.Column(db.Integer, default=150000)
    term = db.Column(db.String(50), nullable=False)
    year = db.Column(db.Integer, nullable=False)

    start_date = db.Column(db.DateTime)
    expiry_date = db.Column(db.DateTime)

    status = db.Column(db.String(30), default="active")


# Create the tables automatically when the application starts
with app.app_context():
    db.create_all()
BASE_DIR = Path(__file__).resolve().parent
UPLOAD_FOLDER = BASE_DIR / "uploads"
OUTPUT_FOLDER = BASE_DIR / "outputs"
TEMPLATE_FOLDER = ROOT_DIR / "templates"

UPLOAD_FOLDER.mkdir(exist_ok=True)
OUTPUT_FOLDER.mkdir(exist_ok=True)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/download-template")
def download_template():
    """Generate and download a fresh Excel template."""
    template_path = UPLOAD_FOLDER / "marks_entry_template.xlsx"
    create_excel_template(str(template_path))
    return send_file(
        template_path,
        as_attachment=True,
        download_name="Uganda_Report_Card_Template.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/generate", methods=["POST"])
def generate():
    """Upload filled Excel and generate report cards."""
    if "excel_file" not in request.files:
        flash("No file selected.", "error")
        return redirect(url_for("index"))

    file = request.files["excel_file"]
    if file.filename == "":
        flash("No file selected.", "error")
        return redirect(url_for("index"))

    if not file.filename.lower().endswith((".xlsx", ".xls")):
        flash("Please upload an Excel file (.xlsx).", "error")
        return redirect(url_for("index"))

    # Save upload
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    upload_name = f"marks_{timestamp}.xlsx"
    upload_path = UPLOAD_FOLDER / upload_name
    file.save(upload_path)

    # Create unique output folder for this run
    run_dir = OUTPUT_FOLDER / f"run_{timestamp}"
    run_dir.mkdir(exist_ok=True)

    try:
        generated = generate_all_reports(str(upload_path), str(run_dir))
        if not generated:
            flash("No report cards were generated. Check that the Students and Marks sheets have data.", "error")
            return redirect(url_for("index"))

        # Create a ZIP of all PDFs + summary
        zip_path = run_dir / "all_report_cards.zip"
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for pdf in generated:
                zf.write(pdf, arcname=Path(pdf).name)
            summary = run_dir / "class_summary.csv"
            if summary.exists():
                zf.write(summary, arcname="class_summary.csv")

        # Load summary for display
        summary_df = pd.read_csv(run_dir / "class_summary.csv") if (run_dir / "class_summary.csv").exists() else None
        summary_records = summary_df.to_dict("records") if summary_df is not None else []

        return render_template(
            "results.html",
            run_id=timestamp,
            count=len(generated),
            summary=summary_records,
            files=[Path(p).name for p in generated],
        )
    except Exception as e:
        flash(f"Error generating reports: {str(e)}", "error")
        return redirect(url_for("index"))


@app.route("/download/<run_id>/<filename>")
def download_file(run_id, filename):
    """Download a single PDF or the ZIP."""
    run_dir = OUTPUT_FOLDER / f"run_{run_id}"
    if not run_dir.exists():
        flash("Files no longer available.", "error")
        return redirect(url_for("index"))
    return send_from_directory(run_dir, filename, as_attachment=True)


@app.route("/download-all/<run_id>")
def download_all(run_id):
    """Download ZIP of all report cards for a run."""
    run_dir = OUTPUT_FOLDER / f"run_{run_id}"
    zip_path = run_dir / "all_report_cards.zip"
    if not zip_path.exists():
        flash("ZIP file not found.", "error")
        return redirect(url_for("index"))
    return send_file(
        zip_path,
        as_attachment=True,
        download_name=f"ReportCards_{run_id}.zip",
        mimetype="application/zip",
    )


@app.route("/preview/<run_id>/<filename>")
def preview_pdf(run_id, filename):
    """Serve PDF for in-browser preview."""
    run_dir = OUTPUT_FOLDER / f"run_{run_id}"
    return send_from_directory(run_dir, filename, mimetype="application/pdf")


@app.route("/grading-guide")
def grading_guide():
    return render_template("grading.html")


@app.route("/health")
def health():
    return jsonify({"status": "ok", "system": "Uganda Report Card System"})


if __name__ == "__main__":
    # Ensure a template exists
    default_template = TEMPLATE_FOLDER / "marks_entry_template.xlsx"
    if not default_template.exists():
        create_excel_template(str(default_template))

    print("=" * 60)
    print("  UGANDA SECONDARY SCHOOL REPORT CARD SYSTEM")
    print("  Web Application starting...")
    print("=" * 60)
    print("  Open in browser:  http://127.0.0.1:5000")
    print("  or               http://0.0.0.0:5000")
    print("=" * 60)
    app.run(host="0.0.0.0", port=5000, debug=False)
