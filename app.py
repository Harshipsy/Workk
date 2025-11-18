# app.py
import os
from datetime import datetime, timedelta, time as dtime, date as ddate
from io import BytesIO

import pytz
from flask import (
    Flask, abort, flash, jsonify, redirect, render_template,
    request, send_file, url_for
)
from flask_login import (
    LoginManager, UserMixin, current_user, login_required,
    login_user, logout_user
)
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
from sqlalchemy import func
from sqlalchemy import event
from sqlalchemy.engine import Engine
import calendar
from datetime import date as pydate
from openpyxl import Workbook     # <-- used to generate Excel without pandas


@event.listens_for(Engine, "connect")
def enable_sqlite_fk(dbapi_connection, connection_record):
    cursor = dbapi_connection.cursor()
    cursor.execute("PRAGMA foreign_keys=ON")
    cursor.close()

# ---------------- CONFIG ----------------
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DB_DIR = os.path.join(BASE_DIR, "data")
os.makedirs(DB_DIR, exist_ok=True)
DB_PATH = os.path.join(DB_DIR, "worktracker.db")

app = Flask(__name__)
app.config["SECRET_KEY"] = "ABC"  # change in production
app.config["SQLALCHEMY_DATABASE_URI"] = f"sqlite:///{DB_PATH}"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)
login_manager = LoginManager(app)
login_manager.login_view = "login"

# Timezone: IST
TZ = pytz.timezone("Asia/Kolkata")


# ---------------- MODELS ----------------
class User(db.Model, UserMixin):
    __tablename__ = "user"
    id = db.Column(db.Integer, primary_key=True)
    fullname = db.Column(db.String(150), nullable=False)
    username = db.Column(db.String(150), unique=True, nullable=False)
    password_hash = db.Column(db.String(256), nullable=False)
    email = db.Column(db.String(150))
    is_admin = db.Column(db.Boolean, default=False)

    db.relationship("WorkItem", backref="user", cascade="all, delete-orphan", passive_deletes=True)

    def check_password(self, pwd):
        return check_password_hash(self.password_hash, pwd)


class DeletedUsername(db.Model):
    __tablename__ = "deleted_username"
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(150), unique=True, nullable=False)


class WorkItem(db.Model):
    __tablename__ = "work_item"
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(
        db.Integer,
        db.ForeignKey("user.id", ondelete="CASCADE"),
        nullable=False
    )
    work_date = db.Column(db.Date, nullable=False)
    section = db.Column(db.String(30), nullable=False)
    content = db.Column(db.String(512), nullable=False)
    created_at = db.Column(db.DateTime, nullable=False, default=lambda: datetime.now(TZ))


class SubmissionLock(db.Model):
    __tablename__ = "submission_lock"
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(
        db.Integer,
        db.ForeignKey("user.id", ondelete="CASCADE"),
        nullable=False
    )
    work_date = db.Column(db.Date, nullable=False)
    submitted_on = db.Column(db.DateTime, nullable=True)
    leave = db.Column(db.Boolean, default=False)

    __table_args__ = (db.UniqueConstraint('user_id', 'work_date', name='_user_date_uc'),)


# ---------------- LOGIN ----------------
@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))


# ---------------- UTIL FUNCTIONS ----------------
def today_ist():
    return datetime.now(TZ).date()

def end_of_day(dt_date: ddate):
    return TZ.localize(datetime.combine(dt_date, dtime(23, 59, 59)))

def can_submit_for_date(target_date: ddate):
    today = today_ist()
    if target_date > today:
        return False
    return True

def ensure_aware(dt):
    if dt is None:
        return None
    if dt.tzinfo is None:
        return TZ.localize(dt)
    return dt.astimezone(TZ)

def is_late(submitted_on, work_date):
    if not submitted_on:
        return True
    submitted = ensure_aware(submitted_on)
    return submitted > end_of_day(work_date)


# ---------------- INIT DB + DEFAULT ADMIN ----------------
with app.app_context():
    db.create_all()

    admin = User.query.filter_by(username="Admin").first()
    if not admin:
        admin = User(
            fullname="Admin",
            username="Admin",
            password_hash=generate_password_hash("123"),
            email="admin@example.com",
            is_admin=True
        )
        db.session.add(admin)
        db.session.commit()


# ---------------- AUTH ROUTES ----------------
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        pwd = request.form.get("password", "")
        user = User.query.filter(func.lower(User.username) == username.lower()).first()
        if user and user.check_password(pwd):
            login_user(user)
            return redirect(url_for("dashboard"))
        flash("Invalid username or password", "danger")
    return render_template("login.html")


@app.route("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for("login"))


# ---------------- DASHBOARD ----------------
@app.route("/")
@login_required
def dashboard():
    if current_user.is_admin:
        return redirect(url_for("admin_dashboard"))

    sel_date_str = request.args.get("date")
    sel_date = datetime.strptime(sel_date_str, "%Y-%m-%d").date() if sel_date_str else today_ist()

    lock = SubmissionLock.query.filter_by(user_id=current_user.id, work_date=sel_date).first()

    items = WorkItem.query.filter_by(user_id=current_user.id, work_date=sel_date).order_by(WorkItem.id.asc()).all()

    month_val = request.args.get("score_month")
    if month_val:
        y, m = map(int, month_val.split("-"))
    else:
        y, m = today_ist().year, today_ist().month

    score_pct = calculate_user_score(current_user.id, y, m)

    return render_template("employee_dashboard.html",
                           sel_date=sel_date,
                           lock=lock,
                           items=items,
                           today=today_ist(),
                           timedelta=timedelta,
                           score_pct=score_pct,
                           score_month=f"{y}-{m:02d}")


# ---------------- EXPORT (NO PANDAS) ----------------
@app.route("/export", methods=["GET", "POST"])
@login_required
def export():
    if request.method == "POST":
        start = request.form.get("start")
        end = request.form.get("end")
        user_id = request.form.get("user_id")

        if not start or not end:
            flash("Select start and end dates", "danger")
            return redirect(request.referrer or url_for("dashboard"))

        start_d = datetime.strptime(start, "%Y-%m-%d").date()
        end_d = datetime.strptime(end, "%Y-%m-%d").date()

        query = WorkItem.query.join(User, User.id == WorkItem.user_id)

        if current_user.is_admin:
            if user_id and user_id != "all":
                query = query.filter(WorkItem.user_id == int(user_id))
        else:
            query = query.filter(WorkItem.user_id == current_user.id)

        query = query.filter(
            WorkItem.work_date >= start_d,
            WorkItem.work_date <= end_d
        ).order_by(WorkItem.work_date.asc())

        rows = query.all()

        # ----------- CREATE EXCEL WITHOUT PANDAS -----------
        wb = Workbook()
        ws = wb.active
        ws.title = "WorkItems"

        headers = ["Name", "Username", "Work Date", "Section", "Content", "Posting Time"]
        ws.append(headers)

        for r in rows:
            user = User.query.get(r.user_id)
            ws.append([
                user.fullname,
                user.username,
                r.work_date.isoformat(),
                r.section,
                r.content,
                ensure_aware(r.created_at).strftime("%Y-%m-%d %H:%M:%S") if r.created_at else ""
            ])

        output = BytesIO()
        wb.save(output)
        output.seek(0)

        return send_file(
            output,
            download_name=f"workitems_{start}_to_{end}.xlsx",
            as_attachment=True,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    employees = User.query.filter_by(is_admin=False).all() if current_user.is_admin else []
    return render_template("export.html", employees=employees)


# ---------------- RUN ----------------
if __name__ == "__main__":
    app.run(debug=True)
