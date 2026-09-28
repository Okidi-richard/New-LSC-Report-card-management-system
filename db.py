"""
Database and authentication models for the Uganda Report Card System.
Uses PostgreSQL on Render via DATABASE_URL and SQLite locally when no URL is set.
"""
import os
from datetime import datetime, date, timedelta
from functools import wraps

from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash
from sqlalchemy import (
    create_engine, Column, Integer, String, DateTime, Date, Boolean,
    ForeignKey, Text, Numeric, UniqueConstraint, inspect, text
)
from sqlalchemy.orm import declarative_base, relationship, sessionmaker, scoped_session

Base = declarative_base()

def database_url():
    url = os.environ.get("DATABASE_URL", "").strip()
    if url:
        if url.startswith("postgres://"):
            url = "postgresql+psycopg2://" + url[len("postgres://"):]
        elif url.startswith("postgresql://"):
            url = "postgresql+psycopg2://" + url[len("postgresql://"):]
        return url
    return "sqlite:///report_card_local.db"

ENGINE = create_engine(database_url(), pool_pre_ping=True)
SessionLocal = scoped_session(sessionmaker(bind=ENGINE, autoflush=False, expire_on_commit=False))

class School(Base):
    __tablename__ = "schools"
    id = Column(Integer, primary_key=True)
    name = Column(String(200), nullable=False)
    district = Column(String(120))
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)
    users = relationship("User", back_populates="school", cascade="all, delete-orphan")
    subscriptions = relationship("Subscription", back_populates="school", cascade="all, delete-orphan")

class User(Base, UserMixin):
    __tablename__ = "users"
    id = Column(Integer, primary_key=True)
    school_id = Column(Integer, ForeignKey("schools.id"), nullable=False)
    full_name = Column(String(160), nullable=False)
    email = Column(String(160), nullable=False)
    password_hash = Column(String(255), nullable=False)
    role = Column(String(30), nullable=False, default="teacher")  # admin / teacher
    active = Column(Boolean, default=True, nullable=False)
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)
    school = relationship("School", back_populates="users")
    __table_args__ = (UniqueConstraint("school_id", "email", name="uq_school_user_email"),)

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

class Subscription(Base):
    __tablename__ = "subscriptions"
    id = Column(Integer, primary_key=True)
    school_id = Column(Integer, ForeignKey("schools.id"), nullable=False)
    amount_ugx = Column(Integer, nullable=False, default=150000)
    status = Column(String(20), nullable=False, default="pending")  # pending/active/expired
    payment_reference = Column(String(120))
    start_date = Column(Date)
    end_date = Column(Date)
    activated_by = Column(Integer, ForeignKey("users.id"))
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)
    school = relationship("School", back_populates="subscriptions")
    __table_args__ = (UniqueConstraint("school_id", "payment_reference", name="uq_school_payment_reference"),)


class Student(Base):
    __tablename__ = "students"
    id = Column(Integer, primary_key=True)
    school_id = Column(Integer, ForeignKey("schools.id"), nullable=False)
    admission_no = Column(String(60), nullable=False)
    lin = Column(String(80))
    surname = Column(String(100), nullable=False)
    first_name = Column(String(100), nullable=False)
    other_names = Column(String(100))
    sex = Column(String(10))
    class_name = Column(String(30), nullable=False)
    stream = Column(String(30))
    date_of_birth = Column(String(30))
    parent_guardian = Column(String(180))
    contact = Column(String(80))
    attendance = Column(Numeric(5,2))
    days_present = Column(Integer)
    days_absent = Column(Integer)
    class_teacher_comment = Column(Text)
    head_teacher_comment = Column(Text)
    photo_path = Column(String(500))
    active = Column(Boolean, default=True, nullable=False)
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)
    school = relationship("School")
    __table_args__ = (
        UniqueConstraint("school_id", "admission_no", name="uq_school_student_admission"),
    )

class MarkEntry(Base):
    __tablename__ = "mark_entries"
    id = Column(Integer, primary_key=True)
    school_id = Column(Integer, ForeignKey("schools.id"), nullable=False)
    teacher_id = Column(Integer, ForeignKey("users.id"), nullable=False)
    admission_no = Column(String(60), nullable=False)
    student_name = Column(String(180), nullable=False)
    class_name = Column(String(30), nullable=False)
    subject = Column(String(120), nullable=False)
    aoi = Column(Numeric(5,2))
    ca = Column(Numeric(5,2))
    formative = Column(Numeric(5,2))  # derived from AOI + CA
    summative = Column(Numeric(5,2))
    final_score = Column(Numeric(5,2))
    teacher_comment = Column(Text)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow, nullable=False)
    teacher = relationship("User")
    school = relationship("School")

class AuditLog(Base):
    __tablename__ = "audit_logs"
    id = Column(Integer, primary_key=True)
    school_id = Column(Integer, ForeignKey("schools.id"), nullable=False)
    user_id = Column(Integer, ForeignKey("users.id"))
    action = Column(String(120), nullable=False)
    details = Column(Text)
    created_at = Column(DateTime, default=datetime.utcnow, nullable=False)

def init_db():
    Base.metadata.create_all(ENGINE)
    # Backward-compatible migration for databases created by earlier versions.
    inspector = inspect(ENGINE)
    if "mark_entries" in inspector.get_table_names():
        existing = {c["name"] for c in inspector.get_columns("mark_entries")}
        with ENGINE.begin() as conn:
            if "aoi" not in existing:
                conn.execute(text("ALTER TABLE mark_entries ADD COLUMN aoi NUMERIC(5,2)"))
            if "ca" not in existing:
                conn.execute(text("ALTER TABLE mark_entries ADD COLUMN ca NUMERIC(5,2)"))
    s = SessionLocal()
    try:
        school = s.query(School).first()
        school_name = os.environ.get("SCHOOL_NAME", "Safe Haven Christian High School Kalongo")
        district = os.environ.get("SCHOOL_DISTRICT", "Agago District")
        if not school:
            school = School(name=school_name, district=district)
            s.add(school)
            s.flush()
        else:
            if os.environ.get("SCHOOL_NAME"):
                school.name = school_name
            if os.environ.get("SCHOOL_DISTRICT"):
                school.district = district

        admin_email = os.environ.get("ADMIN_EMAIL", "admin@safehaven.local").strip().lower()
        admin_password = os.environ.get("ADMIN_PASSWORD", "ChangeMe123!")
        admin = s.query(User).filter_by(school_id=school.id, email=admin_email).first()
        if not admin:
            admin = User(
                school_id=school.id,
                full_name=os.environ.get("ADMIN_NAME", "School Administrator"),
                email=admin_email,
                role="admin",
            )
            admin.set_password(admin_password)
            s.add(admin)
        s.commit()
        return school.id
    finally:
        s.close()

def get_session():
    return SessionLocal()

def subscription_status(s, school_id):
    today = date.today()
    subs = (
        s.query(Subscription)
        .filter(Subscription.school_id == school_id)
        .order_by(Subscription.end_date.desc().nullslast(), Subscription.id.desc())
        .all()
    )
    for sub in subs:
        if sub.status == "active" and sub.end_date and sub.end_date >= today:
            return sub
        if sub.status == "active" and sub.end_date and sub.end_date < today:
            sub.status = "expired"
    s.commit()
    return None

def school_access_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        from flask import redirect, url_for, flash
        from flask_login import current_user
        if not current_user.is_authenticated:
            return redirect(url_for("login"))
        if current_user.role == "admin":
            return view(*args, **kwargs)
        s = get_session()
        try:
            active = subscription_status(s, current_user.school_id)
        finally:
            s.close()
        if not active:
            flash("The school's term subscription is not active. Please contact the school administrator.", "error")
            return redirect(url_for("subscription"))
        return view(*args, **kwargs)
    return wrapped
