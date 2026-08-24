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
    Flask, render_template, render_template_string,request, redirect, url_for,
    flash, send_file, send_from_directory, jsonify, session
)
import pandas as pd
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, logout_user, login_required, current_user
from sqlalchemy import inspect, text
from werkzeug.security import generate_password_hash

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

    formative = db.Column(db.Float, nullable=True)
    summative = db.Column(db.Float, nullable=True)
    mark = db.Column(db.Float, nullable=True)
    status = db.Column(db.String(30), default="draft")

    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(
        db.DateTime,
        default=datetime.utcnow,
        onupdate=datetime.utcnow
    )


class Student(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    admission_number = db.Column(db.String(50), unique=True, nullable=False)
    full_name = db.Column(db.String(150), nullable=False)
    class_name = db.Column(db.String(50), nullable=False)
    gender = db.Column(db.String(20))
    school_id = db.Column(db.Integer, db.ForeignKey("school.id"), nullable=False)
    active = db.Column(db.Boolean, default=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


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
    db.create_all()    # Add new mark columns to existing databases if they do not exist
    inspector = inspect(db.engine)
    mark_columns = {
        column["name"]
        for column in inspector.get_columns("mark_entry")
    }

    with db.engine.begin() as conn:
        if "formative" not in mark_columns:
            conn.execute(
                text("ALTER TABLE mark_entry ADD COLUMN formative FLOAT")
            )

        if "summative" not in mark_columns:
            conn.execute(
                text("ALTER TABLE mark_entry ADD COLUMN summative FLOAT")
            )
        # Create the first school and administrator account
    school = School.query.filter_by(
        name="Safe Haven Christian High School Kalongo"
    ).first()

    if not school:
        school = School(
            name="Safe Haven Christian High School Kalongo",
            district="Agago"
        )
        db.session.add(school)
        db.session.flush()

    admin = User.query.filter_by(username="admin").first()

    if not admin:
        admin = User(
            full_name="System Administrator",
            username="admin",
            password_hash=generate_password_hash("Admin@12345"),
            role="admin",
            school_id=school.id,
            active=True
        )
        db.session.add(admin)
        db.session.commit()
    teacher = User.query.filter_by(username="teacher").first()

    if not teacher:
        teacher = User(
            full_name="Teacher",
            username="teacher",
            password_hash=generate_password_hash("Teacher@12345"),
            role="teacher",
            school_id=school.id,
            active=True
        )
        db.session.add(teacher)
        db.session.commit()
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
    session.clear()
    return redirect(url_for("login"))


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
    user_id = session.get("user_id")

    if not user_id:
        return redirect(url_for("login"))

    user = db.session.get(User, user_id)

    if not user or user.role != "teacher":
        return redirect(url_for("login"))

    return render_template_string("""
    <!DOCTYPE html>
    <html>
    <head>
        <title>Teacher Dashboard</title>
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <style>
            body {
                font-family: Arial, sans-serif;
                background: #f1f5f9;
                padding: 20px;
            }
            .card {
                max-width: 700px;
                margin: auto;
                background: white;
                padding: 25px;
                border-radius: 12px;
                box-shadow: 0 3px 15px rgba(0,0,0,.12);
            }
            .btn {
                display: block;
                padding: 14px;
                margin: 12px 0;
                background: #174a7c;
                color: white;
                text-decoration: none;
                border-radius: 7px;
                text-align: center;
            }
            .logout {
                background: #b91c1c;
            }
        </style>
    </head>
    <body>
        <div class="card">
            <h2>Teacher Dashboard</h2>
            <p>Welcome, <strong>{{ user.full_name }}</strong></p>

            <a class="btn" href="{{ url_for('teacher_marks') }}">
                Enter / Submit Marks
            </a>

            <a class="btn logout" href="{{ url_for('logout') }}">
                Logout
            </a>
        </div>
    </body>
    </html>
    """, user=user)


@app.route("/teacher/marks", methods=["GET", "POST"])
def teacher_marks():
    user_id = session.get("user_id")

    if not user_id:
        return redirect(url_for("login"))

    user = db.session.get(User, user_id)

    if not user or user.role != "teacher":
        return redirect(url_for("login"))

    # Get students belonging to the teacher's school
    students = Student.query.filter_by(
        school_id=user.school_id,
        active=True
    ).order_by(Student.class_name, Student.full_name).all()

    if request.method == "POST":
        student_id = request.form.get("student_id", "").strip()
        subject = request.form.get("subject", "").strip()
        mark_value = request.form.get("mark", "").strip()

        if not student_id or not subject or not mark_value:
            return "Please select a student, enter the subject and enter the mark.", 400

        try:
            mark = float(mark_value)

            if mark < 0 or mark > 100:
                return "Mark must be between 0 and 100.", 400

            student = db.session.get(Student, int(student_id))

            if not student:
                return "Selected student was not found.", 404

            # Make sure the student belongs to the same school
            if student.school_id != user.school_id:
                return "You cannot enter marks for this student.", 403

            entry = MarkEntry(
                teacher_id=user.id,
                student_name=student.full_name,
                class_name=student.class_name,
                subject=subject,
                mark=mark,
                status="submitted"
            )

            db.session.add(entry)
            db.session.commit()

            return redirect(url_for("teacher_marks"))

        except ValueError:
            return "Invalid mark or student selection.", 400

        except Exception:
            db.session.rollback()
            return "Unable to save the mark. Please try again.", 500

    return render_template_string("""
<!DOCTYPE html>
<html>
<head>
    <title>Teacher Mark Entry</title>

    <style>
        body {
            font-family: Arial, sans-serif;
            background: #f1f5f9;
            margin: 0;
            padding: 40px 20px;
        }

        .container {
            max-width: 700px;
            margin: auto;
            background: white;
            padding: 30px;
            border-radius: 12px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.10);
        }

        h1 {
            margin-bottom: 30px;
            color: #111827;
        }

        label {
            display: block;
            font-weight: bold;
            margin-top: 18px;
            margin-bottom: 8px;
        }

        select,
        input {
            width: 100%;
            box-sizing: border-box;
            padding: 13px;
            border: 1px solid #b8b8b8;
            border-radius: 6px;
            font-size: 15px;
        }

        select:focus,
        input:focus {
            outline: none;
            border: 2px solid #2563eb;
        }

        button {
            width: 100%;
            margin-top: 25px;
            padding: 14px;
            border: none;
            border-radius: 6px;
            background: #245b8f;
            color: white;
            font-size: 16px;
            cursor: pointer;
        }

        button:hover {
            background: #1d4f7d;
        }

        .back {
            display: inline-block;
            margin-top: 20px;
            color: #145ca8;
            text-decoration: none;
        }

        .back:hover {
            text-decoration: underline;
        }
    </style>
</head>

<body>

<div class="container">

    <h1>Teacher Mark Entry</h1>

    <form method="POST">

        <label>Student</label>

        <select name="student_id" required>
            <option value="">Select student</option>

            {% for student in students %}
                <option value="{{ student.id }}">
                    {{ student.full_name }} — {{ student.class_name }}
                </option>
            {% endfor %}

        </select>

        <label>Subject</label>

        <input
            type="text"
            name="subject"
            placeholder="Enter subject"
            required
        >

        <label>Mark</label>

        <input
            type="number"
            name="mark"
            min="0"
            max="100"
            step="0.01"
            placeholder="Mark out of 100"
            required
        >

        <button type="submit">
            Submit Mark
        </button>

    </form>

    <a class="back" href="{{ url_for('teacher_dashboard') }}">
        ← Back to Dashboard
    </a>

</div>

</body>
</html>
""", user=user, students=students)
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
        full_name = request.form.get("full_name", "").strip()
        class_name = request.form.get("class_name", "").strip()
        gender = request.form.get("gender", "").strip()

        if admission_number and full_name and class_name:
            student = Student(
                admission_number=admission_number,
                full_name=full_name,
                class_name=class_name,
                gender=gender,
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
                background: #f1f5f9;
                padding: 30px;
            }

            .container {
                max-width: 1000px;
                margin: auto;
                background: white;
                padding: 30px;
                border-radius: 10px;
            }

            input, select {
                padding: 10px;
                margin: 5px;
                width: 20%;
            }

            button {
                padding: 10px 20px;
                background: #1d4f82;
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
                background: #1d4f82;
                color: white;
            }

            a {
                color: #1d4f82;
            }
        </style>
    </head>

    <body>

        <div class="container">

            <h2>Manage Students</h2>

            <p>
                <a href="{{ url_for('admin_dashboard') }}">
                    ← Back to Administration Dashboard
                </a>
            </p>

            <h3>Add Student</h3>

            <form method="POST">

                <input
                    type="text"
                    name="admission_number"
                    placeholder="Admission Number"
                    required
                >

                <input
                    type="text"
                    name="full_name"
                    placeholder="Student Full Name"
                    required
                >

                <input
                    type="text"
                    name="class_name"
                    placeholder="Class e.g. S.4"
                    required
                >

                <select name="gender">
                    <option value="">Gender</option>
                    <option value="Male">Male</option>
                    <option value="Female">Female</option>
                </select>

                <button type="submit">
                    Add Student
                </button>

            </form>

            <h3>Registered Students</h3>

            {% if students %}

            <table>

                <tr>
                    <th>Admission Number</th>
                    <th>Student Name</th>
                    <th>Class</th>
                    <th>Gender</th>
                </tr>

                {% for student in students %}

                <tr>
                    <td>{{ student.admission_number }}</td>
                    <td>{{ student.full_name }}</td>
                    <td>{{ student.class_name }}</td>
                    <td>{{ student.gender or "" }}</td>
                </tr>

                {% endfor %}

            </table>

            {% else %}

            <p>No students have been registered yet.</p>

            {% endif %}

        </div>

    </body>
    </html>
    """, students=students)


@app.route("/admin/dashboard")
def admin_dashboard():
    user_id = session.get("user_id")

    if not user_id:
        return redirect(url_for("login"))

    user = db.session.get(User, user_id)

    if not user or user.role != "admin":
        return redirect(url_for("login"))

    entries = MarkEntry.query.order_by(
        MarkEntry.created_at.desc()
    ).all()

    return render_template_string("""
    <!DOCTYPE html>
    <html>
    <head>
        <title>Administration Dashboard</title>
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <style>
            body {
                font-family: Arial, sans-serif;
                background: #f1f5f9;
                padding: 15px;
            }
            .card {
                background: white;
                padding: 20px;
                border-radius: 12px;
                overflow-x: auto;
            }
            table {
                width: 100%;
                border-collapse: collapse;
                min-width: 700px;
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
            .logout {
                display: inline-block;
                margin-bottom: 15px;
                background: #b91c1c;
                color: white;
                padding: 10px 15px;
                text-decoration: none;
                border-radius: 6px;
            }
        </style>
    </head>
    <body>
        <h2>School Administration Dashboard</h2>

       <a class="button" href="{{ url_for('manage_students') }}">Manage Students</a>
<a class="logout" href="{{ url_for('logout') }}">Logout</a>

        <div class="card">
            <h3>Teacher Mark Submissions</h3>

            {% if entries %}
            <table>
                <tr>
                    <th>Teacher</th>
                    <th>Student</th>
                    <th>Class</th>
                    <th>Subject</th>
                    <th>Mark</th>
                    <th>Status</th>
                    <th>Date</th>
                </tr>

                {% for entry in entries %}
                <tr>
                    <td>{{ entry.teacher.full_name }}</td>
                    <td>{{ entry.student_name }}</td>
                    <td>{{ entry.class_name }}</td>
                    <td>{{ entry.subject }}</td>
                    <td>{{ entry.mark }}</td>
                    <td>{{ entry.status }}</td>
                    <td>{{ entry.created_at }}</td>
                </tr>
                {% endfor %}
            </table>

            {% else %}
            <p>No marks have been submitted yet.</p>
            {% endif %}
        </div>
    </body>
    </html>
    """, entries=entries)

# ==============================
# ADMINISTRATION MONITORING
# ==============================


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
