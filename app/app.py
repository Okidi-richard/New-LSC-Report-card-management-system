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
from sqlalchemy import inspect, text

# Ensure project root is on the path (works locally and on cloud hosts)
ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from flask import (
    Flask, render_template, render_template_string,request, redirect, url_for,
    flash, send_file, send_from_directory, jsonify, session
)
from werkzeug.utils import secure_filename
import pandas as pd
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, logout_user, login_required, current_user
from sqlalchemy import text
from werkzeug.security import generate_password_hash

# Import core logic from the existing system
from report_card_system import (
    create_excel_template,
    generate_all_reports,
    load_data,
    calculate_final_score,
    get_grade,
    SCHOOL_CONFIG,
    build_report_card,
)

app = Flask(__name__)

# Security
app.secret_key = os.environ.get(
    "SECRET_KEY_BASE",
    os.environ.get("SECRET_KEY", "uganda-report-card-secret-key-change-in-production")
)

# Database
database_url = os.environ.get("DATABASE_URL")
if not database_url:
    # Local development fallback only. Render production must provide DATABASE_URL.
    database_url = "sqlite:///" + str(ROOT_DIR / "report_card_local.db")

if database_url and database_url.startswith("postgres://"):
    database_url = database_url.replace("postgres://", "postgresql://", 1)

app.config["SQLALCHEMY_DATABASE_URI"] = database_url
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

# Upload limit: 16 MB
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024
app.config["STUDENT_PHOTO_FOLDER"] = os.path.join(ROOT_DIR, "app", "static", "student_photos")
os.makedirs(app.config["STUDENT_PHOTO_FOLDER"], exist_ok=True)
# Initialize database and login manager
db = SQLAlchemy(app)

# Database initialization is performed lazily so a broken/missing Render
# DATABASE_URL does not prevent the web service from booting and exposing /health.
DB_READY = False
DB_ERROR = None


def init_database():
    """Create/migrate application tables without destroying existing data."""
    global DB_READY, DB_ERROR
    if DB_READY:
        return True
    try:
        db.create_all()
        inspector = inspect(db.engine)

        # Student migration
        student_columns = [c["name"] for c in inspector.get_columns("student")]
        student_migrations = {
            "photo": "VARCHAR(255)",
            "stream": "VARCHAR(50)",
            "date_of_birth": "VARCHAR(30)",
            "parent_guardian": "VARCHAR(150)",
            "contact": "VARCHAR(80)",
            "attendance": "FLOAT",
            "days_present": "INTEGER",
            "days_absent": "INTEGER",
            "class_teacher_comment": "TEXT",
            "head_teacher_comment": "TEXT",
        }
        for name, typ in student_migrations.items():
            if name not in student_columns:
                with db.engine.begin() as conn:
                    conn.execute(text(f"ALTER TABLE student ADD COLUMN {name} {typ}"))

        # Mark-entry migration
        mark_columns = [c["name"] for c in inspect(db.engine).get_columns("mark_entry")]
        mark_migrations = {
            "aoi": "FLOAT", "ca": "FLOAT", "formative": "FLOAT",
            "summative": "FLOAT", "final_score": "FLOAT", "teacher_comment": "TEXT",
            "term": "VARCHAR(50)", "academic_year": "VARCHAR(20)",
        }
        for name, typ in mark_migrations.items():
            if name not in mark_columns:
                with db.engine.begin() as conn:
                    conn.execute(text(f"ALTER TABLE mark_entry ADD COLUMN {name} {typ}"))

        # Add assignment table if it does not exist. SQLAlchemy create_all handles it
        # after the model below has been defined; call again after model definitions.
        school = School.query.filter_by(name="Safe Haven Christian High School Kalongo").first()
        if not school:
            school = School(name="Safe Haven Christian High School Kalongo", district="Agago")
            db.session.add(school)
            db.session.flush()

        admin = User.query.filter_by(username="admin").first()
        if not admin:
            admin = User(
                full_name="System Administrator", username="admin",
                password_hash=generate_password_hash(os.environ.get("DEFAULT_ADMIN_PASSWORD", "Admin@12345")),
                role="admin", school_id=school.id, active=True
            )
            db.session.add(admin)

        teacher = User.query.filter_by(username="teacher").first()
        if not teacher:
            teacher = User(
                full_name="Teacher", username="teacher",
                password_hash=generate_password_hash(os.environ.get("DEFAULT_TEACHER_PASSWORD", "Teacher@12345")),
                role="teacher", school_id=school.id, active=True
            )
            db.session.add(teacher)
        db.session.commit()
        db.create_all()
        DB_READY = True
        DB_ERROR = None
        return True
    except Exception as exc:
        db.session.rollback()
        DB_ERROR = str(exc)
        return False


@app.before_request
def ensure_database_ready():
    if request.endpoint == "health":
        return None
    if not init_database():
        return (render_template_string("""
        <!doctype html><html><head><meta name='viewport' content='width=device-width,initial-scale=1'>
        <title>Database connection</title><style>body{font-family:Arial;background:#f1f5f9;padding:30px}.box{max-width:720px;margin:auto;background:#fff;padding:25px;border-radius:12px}.err{background:#fee2e2;padding:12px;border-radius:8px;color:#991b1b}code{word-break:break-word}</style></head>
        <body><div class='box'><h2>Report Card System</h2><p>The web application is running, but the school database is currently unavailable.</p><div class='err'><b>Database error:</b><br><code>{{ error }}</code></div><p>Check the Render <b>DATABASE_URL</b> and PostgreSQL service. No existing data has been deleted by this application.</p></div></body></html>
        """, error=DB_ERROR), 503)

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_FOLDER = BASE_DIR / "uploads"
OUTPUT_FOLDER = BASE_DIR / "outputs"
TEMPLATE_FOLDER = ROOT_DIR / "templates"

UPLOAD_FOLDER.mkdir(exist_ok=True)
OUTPUT_FOLDER.mkdir(exist_ok=True)


@app.route("/")
def index():
    return render_template_string("""
    <!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1">
    <title>Uganda Report Card System</title><style>body{font-family:Arial;background:#f1f5f9;padding:25px}.box{max-width:720px;margin:40px auto;background:#fff;padding:30px;border-radius:14px;box-shadow:0 3px 15px #0001}a{display:inline-block;padding:12px 18px;background:#174a7c;color:white;text-decoration:none;border-radius:7px;margin:5px}</style></head>
    <body><div class="box"><h1>Uganda Secondary School Report Card System</h1><p>CBC-aligned school report-card management and assessment portal.</p><a href="{{url_for('login')}}">School Login</a><a href="{{url_for('health')}}">System Health</a></div></body></html>
    """)


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
    return jsonify({"status": "ok", "system": "Uganda Report Card System", "database_ready": DB_READY, "database_error": DB_ERROR})

# ============================================================
# TEACHER AND ADMINISTRATION MONITORING SYSTEM
# ============================================================

@app.route("/login", methods=["GET", "POST"])
def login():
    from werkzeug.security import check_password_hash

    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")

        user = User.query.filter_by(username=username, active=True).first()

        if user and check_password_hash(user.password_hash, password):
            session["user_id"] = user.id
            session["role"] = user.role
            login_user(user)
            return redirect(url_for("portal"))

        flash("Invalid username or password.", "error")

    return render_template_string("""
    <!DOCTYPE html>
    <html>
    <head>
        <title>School Login</title>
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <style>
            body {
                font-family: Arial, sans-serif;
                background: #f1f5f9;
                padding: 30px;
            }
            .box {
                max-width: 420px;
                margin: 50px auto;
                background: white;
                padding: 30px;
                border-radius: 12px;
                box-shadow: 0 3px 15px rgba(0,0,0,.15);
            }
            h2 { text-align: center; color: #174a7c; }
            input, button {
                width: 100%;
                padding: 13px;
                margin-top: 10px;
                box-sizing: border-box;
                border-radius: 6px;
                border: 1px solid #ccc;
            }
            button {
                background: #174a7c;
                color: white;
                border: none;
                cursor: pointer;
            }
        </style>
    </head>
    <body>
        <div class="box">
            <h2>UG Uganda Report Card System</h2>
            <p style="text-align:center;">Teacher & Administration Login</p>

            <form method="POST">
                <input type="text" name="username"
                       placeholder="Username" required>

                <input type="password" name="password"
                       placeholder="Password" required>

                <button type="submit">Login</button>
            </form>
        </div>
    </body>
    </html>
    """)


@app.route("/logout")
def logout():
    try: logout_user()
    except Exception: pass
    session.clear()
    return redirect(url_for("login"))


def current_session_user():
    uid = session.get("user_id")
    if not uid:
        return None
    user = db.session.get(User, uid)
    if not user or not user.active:
        session.clear()
        return None
    return user


def admin_required():
    user = current_session_user()
    return user if user and user.role == "admin" else None


def school_for_user(user):
    return School.query.get(user.school_id) if user else None

@app.route("/portal")
def portal():
    user_id = session.get("user_id")

    if not user_id:
        return redirect(url_for("login"))

    user = db.session.get(User, user_id)

    if not user:
        session.clear()
        return redirect(url_for("login"))

    if user.role == "admin":
        return redirect(url_for("admin_dashboard"))

    return redirect(url_for("teacher_dashboard"))


@app.route("/teacher/dashboard")
def teacher_dashboard():
    user = current_session_user()
    if not user or user.role != "teacher":
        return redirect(url_for("login"))
    assignments = TeacherAssignment.query.filter_by(teacher_id=user.id, active=True).order_by(TeacherAssignment.class_name, TeacherAssignment.subject).all()
    return render_template_string("""
    <!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1"><title>Teacher Dashboard</title>
    <style>body{font-family:Arial;background:#f1f5f9;padding:20px}.wrap{max-width:1000px;margin:auto}.card{background:#fff;padding:20px;border-radius:12px;margin-bottom:15px;box-shadow:0 2px 10px #0001}.btn{display:inline-block;padding:10px 14px;background:#174a7c;color:#fff;text-decoration:none;border-radius:6px;margin:3px}.danger{background:#b91c1c}table{width:100%;border-collapse:collapse}th,td{padding:9px;border-bottom:1px solid #ddd;text-align:left}th{background:#174a7c;color:#fff}</style></head>
    <body><div class="wrap"><div class="card"><h2>Teacher Dashboard</h2><p>Welcome, <b>{{user.full_name}}</b></p><a class="btn" href="{{url_for('teacher_marks')}}">Enter / Update Marks</a><a class="btn danger" href="{{url_for('logout')}}">Logout</a></div>
    <div class="card"><h3>My Teaching Assignments</h3>{% if assignments %}<table><tr><th>Class</th><th>Subject</th></tr>{% for a in assignments %}<tr><td>{{a.class_name}}</td><td>{{a.subject}}</td></tr>{% endfor %}</table>{% else %}<p>No assignments have been set by the administrator yet.</p>{% endif %}</div></div></body></html>
    """, user=user, assignments=assignments)

@app.route("/teacher/marks", methods=["GET", "POST"])
def teacher_marks():
    user = current_session_user()
    if not user or user.role != "teacher":
        return redirect(url_for("login"))
    assignments = TeacherAssignment.query.filter_by(teacher_id=user.id, active=True).order_by(TeacherAssignment.class_name, TeacherAssignment.subject).all()
    students = Student.query.filter_by(school_id=user.school_id, active=True).order_by(Student.class_name, Student.full_name).all()

    if request.method == "POST":
        try:
            student = Student.query.filter_by(id=int(request.form.get("student_id")), school_id=user.school_id, active=True).first()
            if not student: raise ValueError("Select a valid registered learner.")
            subject = request.form.get("subject", "").strip()
            class_name = request.form.get("class_name", student.class_name).strip()
            if not subject: raise ValueError("Select or enter a subject.")
            assignment = TeacherAssignment.query.filter_by(teacher_id=user.id, class_name=class_name, subject=subject, active=True).first()
            if assignments and not assignment: raise ValueError("This subject/class is not assigned to your account.")
            aoi, ca, summative = [float(request.form.get(k, "")) for k in ("aoi", "ca", "summative")]
            if any(x < 0 or x > 100 for x in (aoi, ca, summative)): raise ValueError("AOI, CA and Summative must each be between 0 and 100.")
            formative = round((aoi + ca) / 2.0, 2)
            final_score = round((formative * 0.20) + (summative * 0.80), 2)
            term = request.form.get("term", "Term 1").strip()
            year = request.form.get("academic_year", "2026/2027").strip()
            comment = request.form.get("teacher_comment", "").strip()
            entry = MarkEntry.query.filter_by(teacher_id=user.id, student_name=student.full_name, class_name=class_name, subject=subject, term=term, academic_year=year).first()
            if not entry:
                entry = MarkEntry(teacher_id=user.id, student_name=student.full_name, class_name=class_name, subject=subject, term=term, academic_year=year)
                db.session.add(entry)
            entry.aoi, entry.ca, entry.formative, entry.summative, entry.final_score = aoi, ca, formative, summative, final_score
            entry.teacher_comment, entry.mark, entry.status = comment, final_score, "submitted"
            db.session.commit()
            flash("Assessment saved successfully.", "success")
        except Exception as exc:
            db.session.rollback()
            flash(str(exc), "error")
        return redirect(url_for("teacher_marks"))

    return render_template_string("""
    <!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1"><title>Teacher Marks</title>
    <style>body{font-family:Arial;background:#f1f5f9;padding:20px}.card{max-width:900px;margin:auto;background:#fff;padding:24px;border-radius:12px;box-shadow:0 2px 12px #0001}.grid{display:grid;grid-template-columns:repeat(2,1fr);gap:12px}label{font-weight:bold}input,select,textarea{width:100%;box-sizing:border-box;padding:11px;border:1px solid #ccc;border-radius:6px;margin-top:5px}.full{grid-column:1/-1}.calc{padding:12px;background:#f8fafc;border:1px solid #ddd;border-radius:8px;margin:15px 0}.btn{padding:11px 16px;background:#174a7c;color:#fff;border:0;border-radius:6px;text-decoration:none;display:inline-block;margin-top:8px}.back{background:#64748b}@media(max-width:700px){.grid{grid-template-columns:1fr}.full{grid-column:auto}}</style></head>
    <body><div class="card"><h2>Teacher Assessment Entry</h2><p><b>Formative (AOI + CA)</b> is captured as two separate components.</p>
    <form method="post"><div class="grid"><div class="full"><label>Learner</label><select name="student_id" required>{% for s in students %}<option value="{{s.id}}">{{s.admission_number}} — {{s.full_name}} ({{s.class_name}})</option>{% endfor %}</select></div>
    <div><label>Class</label><select name="class_name" id="class_name">{% for a in assignments %}<option value="{{a.class_name}}">{{a.class_name}}</option>{% endfor %}{% if not assignments %}{% for s in students|unique(attribute='class_name') %}<option value="{{s.class_name}}">{{s.class_name}}</option>{% endfor %}{% endif %}</select></div>
    <div><label>Subject</label><select name="subject" id="subject">{% for a in assignments %}<option value="{{a.subject}}">{{a.subject}}</option>{% endfor %}{% if not assignments %}<option value="Agriculture">Agriculture</option>{% endif %}</select></div>
    <div><label>Term</label><select name="term"><option>Term 1</option><option>Term 2</option><option>Term 3</option></select></div><div><label>Academic Year</label><input name="academic_year" value="2026/2027"></div>
    <div><label>AOI / 100</label><input id="aoi" name="aoi" type="number" min="0" max="100" step="0.01" required></div><div><label>CA / 100</label><input id="ca" name="ca" type="number" min="0" max="100" step="0.01" required></div><div><label>Summative / 100</label><input id="summative" name="summative" type="number" min="0" max="100" step="0.01" required></div></div>
    <div class="calc"><b>Formative (AOI + CA):</b> <span id="f">0.00</span> &nbsp; | &nbsp; <b>Final Score:</b> <span id="fs">0.00</span><br><small>Current system configuration: Formative = average of AOI and CA; Final = 20% Formative + 80% Summative.</small></div>
    <label>Subject Teacher Comment</label><textarea name="teacher_comment" rows="3"></textarea><br><button class="btn" type="submit">Save Assessment</button> <a class="btn back" href="{{url_for('teacher_dashboard')}}">Back</a></form></div>
    <script>function c(){let a=+document.getElementById('aoi').value||0,ca=+document.getElementById('ca').value||0,s=+document.getElementById('summative').value||0,f=(a+ca)/2,fs=f*.2+s*.8;document.getElementById('f').textContent=f.toFixed(2);document.getElementById('fs').textContent=fs.toFixed(2)}['aoi','ca','summative'].forEach(x=>document.getElementById(x).addEventListener('input',c));</script></body></html>
    """, user=user, assignments=assignments, students=students)

@app.route("/admin/dashboard")
def admin_dashboard():
    user_id = session.get("user_id")
    if not user_id:
        return redirect(url_for("login"))

    user = db.session.get(User, user_id)
    if not user or user.role != "admin":
        return redirect(url_for("login"))

    entries = MarkEntry.query.order_by(MarkEntry.created_at.desc()).all()
    students_count = Student.query.filter_by(school_id=user.school_id, active=True).count()
    teachers_count = User.query.filter_by(school_id=user.school_id, role="teacher", active=True).count()

    return render_template_string("""
    <!DOCTYPE html>
    <html>
    <head>
      <title>Administration Dashboard</title>
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <style>
        body{font-family:Arial,sans-serif;background:#f1f5f9;padding:15px}
        .wrap{max-width:1200px;margin:auto}
        .cards{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:15px}
        .card{background:white;padding:18px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,.08)}
        .stat{font-size:28px;font-weight:bold;color:#174a7c}
        table{width:100%;border-collapse:collapse;min-width:1000px}
        th,td{padding:9px;border-bottom:1px solid #ddd;text-align:left}
        th{background:#174a7c;color:white}
        .table-wrap{overflow:auto}
        .btn{display:inline-block;padding:10px 14px;background:#174a7c;color:white;text-decoration:none;border-radius:6px;margin:3px}
        .logout{background:#b91c1c}
        @media(max-width:700px){.cards{grid-template-columns:1fr}}
      </style>
    </head>
    <body>
      <div class="wrap">
        <h2>School Administration Dashboard</h2>
        <p>Welcome, <strong>{{user.full_name}}</strong></p>
        <div class="cards">
          <div class="card"><div>Registered learners</div><div class="stat">{{students_count}}</div></div>
          <div class="card"><div>Active teachers</div><div class="stat">{{teachers_count}}</div></div>
          <div class="card"><div>Assessment entries</div><div class="stat">{{entries|length}}</div></div>
        </div>
        <div class="card">
          <a class="btn" href="{{url_for('manage_students')}}">Manage Students</a>
          <a class="btn" href="{{url_for('manage_teachers')}}">Teachers & Assignments</a>
          <a class="btn" href="{{url_for('subscription')}}">Subscription</a>
          <a class="btn logout" href="{{url_for('logout')}}">Logout</a>
        </div>
        <div class="card">
          <h3>Teacher Assessment Submissions</h3>
          {% if entries %}
          <div class="table-wrap"><table>
            <tr><th>Teacher</th><th>Student</th><th>Class</th><th>Subject</th><th>AOI</th><th>CA</th><th>Formative</th><th>Summative</th><th>Final</th><th>Status</th><th>Date</th></tr>
            {% for entry in entries %}
            <tr>
              <td>{{entry.teacher.full_name}}</td>
              <td>{{entry.student_name}}</td>
              <td>{{entry.class_name}}</td>
              <td>{{entry.subject}}</td>
              <td>{{entry.aoi if entry.aoi is not none else "—"}}</td>
              <td>{{entry.ca if entry.ca is not none else "—"}}</td>
              <td>{{entry.formative if entry.formative is not none else "—"}}</td>
              <td>{{entry.summative if entry.summative is not none else "—"}}</td>
              <td>{{entry.final_score if entry.final_score is not none else entry.mark}}</td>
              <td>{{entry.status}}</td>
              <td>{{entry.created_at}}</td>
            </tr>
            {% endfor %}
          </table></div>
          {% else %}<p>No assessment entries have been submitted yet.</p>{% endif %}
        </div>
      </div>
    </body>
    </html>
    """, user=user, entries=entries, students_count=students_count, teachers_count=teachers_count)

# ==============================
# ADMINISTRATION MONITORING
# ==============================


@app.route("/admin/teachers", methods=["GET", "POST"])
def manage_teachers():
    admin = admin_required()
    if not admin: return redirect(url_for("login"))
    if request.method == "POST":
        action = request.form.get("action")
        if action == "add":
            username = request.form.get("username", "").strip()
            full_name = request.form.get("full_name", "").strip()
            password = request.form.get("password", "").strip()
            if username and full_name and password and not User.query.filter_by(username=username).first():
                db.session.add(User(full_name=full_name, username=username, password_hash=generate_password_hash(password), role="teacher", school_id=admin.school_id, active=True))
                db.session.commit(); flash("Teacher account created.", "success")
        elif action == "assign":
            tid = int(request.form["teacher_id"]); cls = request.form["class_name"].strip(); subject = request.form["subject"].strip()
            teacher = User.query.filter_by(id=tid, school_id=admin.school_id, role="teacher").first()
            if teacher and cls and subject and not TeacherAssignment.query.filter_by(teacher_id=tid, class_name=cls, subject=subject, active=True).first():
                db.session.add(TeacherAssignment(teacher_id=tid, class_name=cls, subject=subject)); db.session.commit(); flash("Teaching assignment added.", "success")
        return redirect(url_for("manage_teachers"))
    teachers = User.query.filter_by(school_id=admin.school_id, role="teacher").order_by(User.full_name).all()
    assignments = TeacherAssignment.query.join(User).filter(User.school_id==admin.school_id, TeacherAssignment.active==True).order_by(User.full_name, TeacherAssignment.class_name, TeacherAssignment.subject).all()
    return render_template_string("""
    <!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1"><title>Teachers</title><style>body{font-family:Arial;background:#f1f5f9;padding:20px}.wrap{max-width:1000px;margin:auto}.card{background:#fff;padding:20px;border-radius:12px;margin-bottom:15px}.grid{display:grid;grid-template-columns:repeat(3,1fr);gap:10px}input,select,button{padding:10px;width:100%;box-sizing:border-box;margin-top:5px}.btn{display:inline-block;background:#174a7c;color:#fff;padding:10px 14px;border-radius:6px;text-decoration:none;border:0}table{width:100%;border-collapse:collapse}th,td{padding:9px;border-bottom:1px solid #ddd;text-align:left}th{background:#174a7c;color:#fff}@media(max-width:700px){.grid{grid-template-columns:1fr}}</style></head><body><div class="wrap"><a href="{{url_for('admin_dashboard')}}">← Dashboard</a><div class="card"><h2>Teacher Accounts</h2><form method="post"><input type="hidden" name="action" value="add"><div class="grid"><div><label>Full name</label><input name="full_name" required></div><div><label>Username</label><input name="username" required></div><div><label>Temporary password</label><input name="password" required></div></div><button class="btn">Create Teacher</button></form></div><div class="card"><h2>Assign Subject/Class</h2><form method="post"><input type="hidden" name="action" value="assign"><div class="grid"><div><label>Teacher</label><select name="teacher_id">{% for t in teachers %}<option value="{{t.id}}">{{t.full_name}} ({{t.username}})</option>{% endfor %}</select></div><div><label>Class</label><input name="class_name" placeholder="S.4" required></div><div><label>Subject</label><input name="subject" placeholder="Agriculture" required></div></div><button class="btn">Add Assignment</button></form></div><div class="card"><h2>Current Assignments</h2><table><tr><th>Teacher</th><th>Class</th><th>Subject</th></tr>{% for a in assignments %}<tr><td>{{a.teacher.full_name}}</td><td>{{a.class_name}}</td><td>{{a.subject}}</td></tr>{% else %}<tr><td colspan="3">No assignments yet.</td></tr>{% endfor %}</table></div></div></body></html>
    """, teachers=teachers, assignments=assignments)


@app.route("/admin/subscription", methods=["GET", "POST"])
def subscription():
    admin = admin_required()
    if not admin: return redirect(url_for("login"))
    if request.method == "POST":
        try:
            amount = int(request.form.get("amount", 150000)); term = request.form.get("term", "Term 1"); year = int(request.form.get("year", datetime.now().year)); days = int(request.form.get("days", 120))
            now = datetime.utcnow()
            db.session.add(Subscription(school_id=admin.school_id, amount=amount, term=term, year=year, start_date=now, expiry_date=now.fromtimestamp(now.timestamp()+days*86400), status="active")); db.session.commit(); flash("Subscription activated.", "success")
        except Exception as exc: db.session.rollback(); flash(str(exc), "error")
    subs = Subscription.query.filter_by(school_id=admin.school_id).order_by(Subscription.id.desc()).all()
    return render_template_string("""<!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1"><title>Subscription</title><style>body{font-family:Arial;background:#f1f5f9;padding:20px}.card{max-width:850px;margin:auto;background:#fff;padding:22px;border-radius:12px;margin-bottom:15px}input,select,button{padding:10px;margin:5px;width:calc(100% - 10px);box-sizing:border-box}.btn{background:#174a7c;color:#fff;border:0;border-radius:6px}</style></head><body><div class="card"><a href="{{url_for('admin_dashboard')}}">← Dashboard</a><h2>School Subscription</h2><form method="post"><input name="amount" type="number" value="150000"><select name="term"><option>Term 1</option><option>Term 2</option><option>Term 3</option></select><input name="year" type="number" value="{{year}}"><input name="days" type="number" value="120"><button class="btn">Activate Subscription</button></form></div><div class="card"><h3>Subscription History</h3>{% for s in subs %}<p>{{s.term}} {{s.year}} — UGX {{s.amount}} — {{s.status}} — expires {{s.expiry_date}}</p>{% else %}<p>No subscription records.</p>{% endfor %}</div></body></html>""", subs=subs, year=datetime.now().year)


def build_student_report(student, entries, school, term, year, output_path):
    marks=[]
    for e in entries:
        marks.append({
            "Subject": e.subject,
            "AOI (out of 100)": e.aoi,
            "CA (out of 100)": e.ca,
            "Formative (out of 100)": e.formative,
            "Summative (out of 100)": e.summative,
            "Subject Teacher": e.teacher.full_name if e.teacher else "",
            "Teacher Comment (optional)": e.teacher_comment or "",
        })
    student_dict={
        "Surname": student.full_name, "First Name":"", "Other Names":"",
        "Admission No":student.admission_number, "LIN":student.lin or "",
        "Class":student.class_name, "Stream":student.stream or "", "Sex":student.gender or "",
        "Date of Birth":student.date_of_birth or "", "Parent/Guardian":student.parent_guardian or "",
        "Contact":student.contact or "", "Attendance (%)":student.attendance,
        "Days Present":student.days_present, "Days Absent":student.days_absent,
        "Class Teacher Comment":student.class_teacher_comment or "",
        "Head Teacher Comment":student.head_teacher_comment or "",
        "Photo Path": str(Path(app.config["STUDENT_PHOTO_FOLDER"]) / student.photo) if student.photo else "",
    }
    school_dict={"name":school.name, "motto":os.environ.get("SCHOOL_MOTTO", "Achieving excellence together"), "address":school.district or "", "phone":os.environ.get("SCHOOL_PHONE", ""), "email":os.environ.get("SCHOOL_EMAIL", "")}
    build_report_card(student_dict, marks, school_dict, term, year, datetime.now().strftime("%d/%m/%Y"), str(output_path))


@app.route("/admin/reports/student/<int:student_id>")
def student_report(student_id):
    admin=admin_required()
    if not admin: return redirect(url_for("login"))
    student=Student.query.filter_by(id=student_id, school_id=admin.school_id, active=True).first_or_404()
    term=request.args.get("term", "Term 1"); year=request.args.get("year", "2026/2027")
    entries=MarkEntry.query.filter_by(student_name=student.full_name, class_name=student.class_name, term=term, academic_year=year).order_by(MarkEntry.subject).all()
    if not entries: flash("No marks found for this learner for the selected term/year.", "error"); return redirect(url_for("manage_students"))
    out=OUTPUT_FOLDER/f"ReportCard_{student.admission_number.replace('/','-')}_{datetime.now().strftime('%Y%m%d%H%M%S')}.pdf"
    build_student_report(student, entries, school_for_user(admin), term, year, out)
    return send_file(out, as_attachment=True, download_name=out.name, mimetype="application/pdf")


@app.route("/admin/reports/class/<class_name>")
def class_reports(class_name):
    admin=admin_required()
    if not admin: return redirect(url_for("login"))
    term=request.args.get("term", "Term 1"); year=request.args.get("year", "2026/2027")
    students=Student.query.filter_by(school_id=admin.school_id, class_name=class_name, active=True).order_by(Student.full_name).all()
    run=OUTPUT_FOLDER/f"class_{secure_filename(class_name)}_{datetime.now().strftime('%Y%m%d%H%M%S')}"; run.mkdir(parents=True,exist_ok=True)
    made=[]
    for s in students:
        entries=MarkEntry.query.filter_by(student_name=s.full_name,class_name=s.class_name,term=term,academic_year=year).all()
        if entries:
            out=run/f"ReportCard_{secure_filename(s.admission_number)}_{secure_filename(s.full_name)}.pdf"; build_student_report(s,entries,school_for_user(admin),term,year,out); made.append(out)
    if not made: flash("No report-card marks found for this class/term/year.","error"); return redirect(url_for("manage_students"))
    z=run/"all_report_cards.zip"
    with zipfile.ZipFile(z,"w",zipfile.ZIP_DEFLATED) as f:
        for p in made: f.write(p,arcname=p.name)
    return send_file(z,as_attachment=True,download_name=f"ReportCards_{class_name}_{term}_{year}.zip",mimetype="application/zip")

@app.route("/admin/students", methods=["GET", "POST"])
def manage_students():
    user_id = session.get("user_id")

    if not user_id:
        return redirect(url_for("login"))

    user = db.session.get(User, user_id)

    if not user or user.role != "admin":
        return redirect(url_for("login"))

    if request.method == "POST":
        admission_number = request.form.get("admission_number", "").strip()
        lin = request.form.get("lin", "").strip() or None
        full_name = request.form.get("full_name", "").strip()
        class_name = request.form.get("class_name", "").strip()
        gender = request.form.get("gender", "").strip()

        photo = None
        photo_file = request.files.get("photo")

        if photo_file and photo_file.filename:
            filename = secure_filename(photo_file.filename)
            photo_file.save(os.path.join(app.config["STUDENT_PHOTO_FOLDER"], filename))
            photo = filename

        if admission_number and full_name and class_name:
            existing = Student.query.filter_by(
                admission_number=admission_number,
                school_id=user.school_id
            ).first()

            if not existing:
                student = Student(
                    admission_number=admission_number,
                                lin=lin,
                    full_name=full_name,
                    class_name=class_name,
                    gender=gender,
                    photo=photo,
                    school_id=user.school_id,
                    active=True
                )

                db.session.add(student)
                db.session.commit()

        return redirect(url_for("manage_students"))

    students = Student.query.filter_by(
        school_id=user.school_id,
        active=True
    ).order_by(Student.full_name.asc()).all()

    return render_template_string("""
    <!DOCTYPE html>
    <html>
    <head>
        <title>Manage Students</title>
        <style>
            body {
                font-family: Arial, sans-serif;
                margin: 30px;
                background: #f4f6f8;
            }

            .container {
                max-width: 1100px;
                margin: auto;
                background: white;
                padding: 25px;
                border-radius: 10px;
            }

            h2 {
                color: #174a7c;
            }

            input, select {
                padding: 10px;
                margin: 5px;
                border: 1px solid #ccc;
                border-radius: 5px;
            }

            button {
                padding: 10px 18px;
                background: #174a7c;
                color: white;
                border: none;
                border-radius: 5px;
                cursor: pointer;
            }

            table {
                width: 100%;
                border-collapse: collapse;
                margin-top: 25px;
            }

            th, td {
                padding: 10px;
                border-bottom: 1px solid #ddd;
                text-align: left;
            }

            th {
                background: #174a7c;
                color: white;
            }

            .back {
                display: inline-block;
                margin-bottom: 20px;
                text-decoration: none;
                color: #174a7c;
            }
        </style>
    </head>

    <body>
        <div class="container">

            <a class="back" href="{{ url_for('admin_dashboard') }}">
                ← Back to Dashboard
            </a>

            <h2>Student Management</h2>

           <form method="POST" enctype="multipart/form-data">

                <input
                    type="text"
                    name="admission_number"
                    placeholder="Admission Number"
                    required
                >
        <input
            type="text"
            name="lin"
            placeholder="Student LIN"
        >
                <input
                    type="text"
                    name="full_name"
                    placeholder="Student Full Name"
                    required
                >

               <input
    type="file"
    name="photo"
    accept="image/*"
> 
                <input
                    type="text"
                    name="class_name"
                    placeholder="Class (e.g. S.1)"
                    required
                >

                <select name="gender">
                    <option value="">Gender</option>
                    <option value="Male">Male</option>
                    <option value="Female">Female</option>
                </select>

                <button type="submit">Add Student</button>

            </form>

            <table>
                <tr>
                    <th>No.</th>
                    <th>Admission Number</th>
<th>Student LIN</th>

                    <th>Student Name</th>
                    <th>Photo</th>
                    <th>Class</th>
                    <th>Gender</th><th>Report</th>
                </tr>

                {% for student in students %}
                <tr>
                    <td>{{ loop.index }}</td>
                    <td>{{ student.admission_number }}</td>
                    <td>{{ student.lin or "" }}</td>
                    <td>{{ student.full_name }}</td>
                    <td>
    {% if student.photo %}
        <img src="{{ url_for('static', filename='student_photos/' + student.photo) }}"
             width="50" height="50"
             style="object-fit: cover; border-radius: 5px;">
    {% else %}
        No photo
    {% endif %}
</td>
                    <td>{{ student.class_name }}</td>
                    <td>{{ student.gender or "" }}</td>
                    <td><a href="{{url_for('student_report', student_id=student.id)}}">Report</a></td>
                </tr>
                {% else %}
                <tr>
                    <td colspan="7">No students registered yet.</td>
                </tr>
                {% endfor %}

            </table>

        </div>
    </body>
    </html>
    """, students=students)
# V6 features: teacher assignments, AOI+CA entry, term/year marks, admin report generation.

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
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=False)
