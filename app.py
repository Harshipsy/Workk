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
from openpyxl import Workbook


@event.listens_for(Engine, "connect")
def enable_sqlite_fk(dbapi_connection, connection_record):
    cursor = dbapi_connection.cursor()
    cursor.execute("PRAGMA foreign_keys=ON")
    cursor.close()

# ---------------- CONFIG ----------------
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DB_DIR = os.path.join(BASE_DIR, "data")
os.makedirs(DB_DIR, exist_ok=True)
db_url = os.environ.get("DATABASE_URL")
if db_url.startswith("postgres://"):
    # Render sometimes gives deprecated URL form
    db_url = db_url.replace("postgres://", "postgresql://", 1)

app.config["SQLALCHEMY_DATABASE_URI"] = db_url

app = Flask(__name__)
app.config["SECRET_KEY"] = "ABC"

# ---------------- PostgreSQL CONFIG ----------------
db_url = os.environ.get("DATABASE_URL")
if db_url and db_url.startswith("postgres://"):
    db_url = db_url.replace("postgres://", "postgresql://", 1)

app.config["SQLALCHEMY_DATABASE_URI"] = db_url
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

# Initialize DB
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


def calculate_user_score(user_id, year, month):
    today = today_ist()
    _, last_day = calendar.monthrange(year, month)
    start_date = ddate(year, month, 1)
    end_of_month = ddate(year, month, last_day)
    considered_end = min(end_of_month, today)

    dates = []
    d = start_date
    while d <= considered_end:
        dates.append(d)
        d += timedelta(days=1)

    locks = SubmissionLock.query.filter(
        SubmissionLock.user_id == user_id,
        SubmissionLock.work_date >= start_date,
        SubmissionLock.work_date <= considered_end
    ).all()
    lock_map = {l.work_date: l for l in locks}

    on_time = 0
    late = 0
    leave = 0
    not_sub = 0

    for d in dates:
        if d.weekday() == 6:
            continue
        lock = lock_map.get(d)
        if not lock:
            not_sub += 1
        else:
            if lock.leave:
                leave += 1
            else:
                if not is_late(lock.submitted_on, d):
                    on_time += 1
                else:
                    late += 1

    working_days = len([d for d in dates if d.weekday() != 6])
    denom = working_days - leave
    if denom <= 0:
        pct = 100
    else:
        pct = round((on_time / denom) * 100)

    return pct
# ---------------- INIT DB + DEFAULT ADMIN ----------------
with app.app_context():
    db.create_all()

    admin = User.query.filter_by(username="Admin").first()
    if not admin:
        admin = User(
            fullname="Admin",
            username="Admin",
            password_hash=generate_password_hash("2264"),
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


# ---------------- EMPLOYEE DASHBOARD ----------------
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


# ---------------- SUBMIT ----------------
@app.route("/submit/<date_str>", methods=["POST"])
@login_required
def submit(date_str):
    if current_user.is_admin:
        abort(403)
    target_date = datetime.strptime(date_str, "%Y-%m-%d").date()

    existing_lock = SubmissionLock.query.filter_by(user_id=current_user.id, work_date=target_date).first()
    if existing_lock:
        flash("Already submitted for that date (locked).", "warning")
        return redirect(url_for("dashboard", date=target_date.isoformat()))

    if not can_submit_for_date(target_date):
        flash("Submission window for that date has closed.", "danger")
        return redirect(url_for("dashboard", date=target_date.isoformat()))

    completed = request.form.getlist("completed[]")
    under = request.form.getlist("underprocess[]")
    misc = request.form.getlist("misc[]")

    to_add = []
    now = datetime.now(TZ)
    for line in completed:
        line = (line or "").strip()
        if line:
            to_add.append(WorkItem(user_id=current_user.id, work_date=target_date, section="completed", content=line, created_at=now))
    for line in under:
        line = (line or "").strip()
        if line:
            to_add.append(WorkItem(user_id=current_user.id, work_date=target_date, section="underprocess", content=line, created_at=now))
    for line in misc:
        line = (line or "").strip()
        if line:
            to_add.append(WorkItem(user_id=current_user.id, work_date=target_date, section="misc", content=line, created_at=now))

    if to_add:
        db.session.add_all(to_add)

    lock = SubmissionLock(user_id=current_user.id, work_date=target_date, submitted_on=now.astimezone(TZ), leave=False)
    db.session.add(lock)
    db.session.commit()

    flash("Submitted and locked for that date.", "success")
    return redirect(url_for("dashboard", date=target_date.isoformat()))


# ---------------- MARK LEAVE ----------------
@app.route("/mark_leave/<date_str>", methods=["POST", "GET"])
@login_required
def mark_leave(date_str):
    if current_user.is_admin:
        abort(403)
    target_date = datetime.strptime(date_str, "%Y-%m-%d").date()

    if not can_submit_for_date(target_date):
        flash("Leave marking window closed for that date.", "danger")
        return redirect(url_for("dashboard", date=target_date.isoformat()))

    existing = SubmissionLock.query.filter_by(user_id=current_user.id, work_date=target_date).first()
    if existing:
        flash("Already marked/submitted for that date.", "warning")
        return redirect(url_for("dashboard", date=target_date.isoformat()))

    lock = SubmissionLock(user_id=current_user.id, work_date=target_date, submitted_on=datetime.now(TZ), leave=True)
    db.session.add(lock)
    db.session.commit()
    flash("Marked as On Leave for that date.", "success")
    return redirect(url_for("dashboard", date=target_date.isoformat()))


# ---------------- CALENDAR STATUS ----------------
@app.route("/calendar_status", methods=["GET"])
@login_required
def calendar_status():
    start = request.args.get("start")
    end = request.args.get("end")

    today = today_ist()

    if not start or not end:
        start = today.replace(day=1).isoformat()
        end = (today.replace(day=28) + timedelta(days=10)).isoformat()

    start_date = datetime.strptime(start, "%Y-%m-%d").date()
    end_date = datetime.strptime(end, "%Y-%m-%d").date()

    locks = SubmissionLock.query.filter(
        SubmissionLock.user_id == current_user.id,
        SubmissionLock.work_date >= start_date,
        SubmissionLock.work_date <= end_date
    ).all()

    lock_map = {l.work_date: l for l in locks}
    result = {}

    for offset in range((end_date - start_date).days + 1):
        d = start_date + timedelta(days=offset)
        key = d.isoformat()

        if d > today:
            result[key] = "grey"
            continue

        lock = lock_map.get(d)

        if d == today:
            if not lock:
                result[key] = "yellow"
            else:
                if lock.leave:
                    result[key] = "leave"
                else:
                    result[key] = "green" if not is_late(lock.submitted_on, d) else "orange"
            continue

        if not lock:
            result[key] = "red"
        else:
            if lock.leave:
                result[key] = "leave"
            else:
                result[key] = "green" if not is_late(lock.submitted_on, d) else "orange"

    return jsonify(result)


# ---------------- ADMIN DASHBOARD ----------------
@app.route("/admin")
@login_required
def admin_dashboard():
    if not current_user.is_admin:
        return redirect(url_for("dashboard"))

    date_str = request.args.get("date")
    selected_date = datetime.strptime(date_str, "%Y-%m-%d").date() if date_str else today_ist()

    min_date = (selected_date - timedelta(days=365)).isoformat()
    max_date = today_ist().isoformat()

    users = User.query.filter_by(is_admin=False).all()
    locks = SubmissionLock.query.filter_by(work_date=selected_date).all()
    lock_map = {l.user_id: l for l in locks}
    all_items = WorkItem.query.filter_by(work_date=selected_date).all()

    summary = []
    for u in users:
        user_items = [it for it in all_items if it.user_id == u.id]
        lock = lock_map.get(u.id)
        submitted_on = lock.submitted_on if lock else None
        leave = lock.leave if lock else False
        late = submitted_on and not leave and is_late(submitted_on, selected_date)

        summary.append({
            "name": u.fullname,
            "work_items": user_items,
            "submitted_on": submitted_on,
            "leave": leave,
            "late": late
        })

    summary.sort(key=lambda x: (x["submitted_on"] is None, x["submitted_on"]), reverse=True)

    return render_template("admin_dashboard.html",
                           summary=summary,
                           selected_date=selected_date,
                           min_date=min_date,
                           max_date=max_date)


# ---------------- ADMIN SCORECARD ----------------
@app.route("/admin/scorecard", methods=["GET", "POST"])
@login_required
def admin_scorecard():
    if not current_user.is_admin:
        return redirect(url_for("dashboard"))

    month_val = request.values.get("month")
    today = today_ist()

    if not month_val:
        month_val = today.strftime("%Y-%m")

    try:
        year, month = map(int, month_val.split("-"))
    except Exception:
        flash("Invalid month selected", "danger")
        return redirect(url_for("admin_dashboard"))

    _, last_day = calendar.monthrange(year, month)
    start_date = pydate(year, month, 1)
    end_of_month = pydate(year, month, last_day)
    considered_end = min(end_of_month, today)

    dates = []
    if start_date <= considered_end:
        d = start_date
        while d <= considered_end:
            dates.append(d)
            d += timedelta(days=1)

    users = User.query.filter_by(is_admin=False).all()
    locks = SubmissionLock.query.filter(
        SubmissionLock.work_date >= start_date,
        SubmissionLock.work_date <= considered_end
    ).all()

    lock_map = {(l.user_id, l.work_date): l for l in locks}
    rows = []

    for u in users:
        on_time = 0
        late = 0
        leave = 0
        not_submitted = 0

        for d in dates:
            if d.weekday() == 6:
                continue

            lock = lock_map.get((u.id, d))
            if not lock:
                not_submitted += 1
            else:
                if lock.leave:
                    leave += 1
                else:
                    if not is_late(lock.submitted_on, d):
                        on_time += 1
                    else:
                        late += 1

        working_days = len([d for d in dates if d.weekday() != 6])
        denom = working_days - leave

        pct = 100 if denom <= 0 else round((on_time / denom) * 100)

        rows.append({
            "username": u.username,
            "fullname": u.fullname,
            "on_time": on_time,
            "late": late,
            "leave": leave,
            "not_submitted": not_submitted,
            "denom": denom,
            "pct": pct
        })

    rows.sort(key=lambda x: x["pct"], reverse=True)

    month_title = start_date.strftime("%B %Y")

    return render_template("admin_scorecard.html",
                           rows=rows,
                           month_val=month_val,
                           month_title=month_title,
                           start_date=start_date,
                           end_date=considered_end)


# ---------------- EMPLOYEE MANAGEMENT ----------------
@app.route("/admin/employees", methods=["GET", "POST"])
@login_required
def admin_employees():
    if not current_user.is_admin:
        abort(403)

    if request.method == "POST":
        fullname = request.form.get("fullname", "").strip()
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        email = request.form.get("email", "").strip()

        if not (fullname and username and password and email):
            flash("All fields are required.", "danger")
            return redirect(url_for("admin_employees"))

        existing = User.query.filter(func.lower(User.username) == username.lower()).first()
        if existing:
            flash("This username already exists.", "danger")
            return redirect(url_for("admin_employees"))

        prev = DeletedUsername.query.filter(func.lower(DeletedUsername.username) == username.lower()).first()
        if prev:
            flash("This username was used before and cannot be reused.", "danger")
            return redirect(url_for("admin_employees"))

        u = User(
            fullname=fullname,
            username=username,
            password_hash=generate_password_hash(password),
            email=email,
            is_admin=False
        )
        db.session.add(u)
        db.session.commit()
        flash("Employee added", "success")
        return redirect(url_for("admin_employees"))

    employees = User.query.filter_by(is_admin=False).all()
    return render_template("admin_employees.html", employees=employees)


@app.route("/admin/employee/<int:user_id>/edit", methods=["GET", "POST"])
@login_required
def admin_employee_edit(user_id):
    if not current_user.is_admin:
        abort(403)

    u = User.query.get_or_404(user_id)
    if request.method == "POST":
        u.fullname = request.form.get("fullname", u.fullname).strip()
        u.email = request.form.get("email", u.email).strip()
        pwd = request.form.get("password")
        if pwd:
            u.password_hash = generate_password_hash(pwd)
        db.session.commit()
        flash("Saved", "success")
        return redirect(url_for("admin_employees"))
    return render_template("admin_employee_edit.html", u=u)


@app.route("/admin/employee/<int:user_id>/delete", methods=["POST"])
@login_required
def admin_employee_delete(user_id):
    if not current_user.is_admin:
        abort(403)

    u = User.query.get_or_404(user_id)

    try:
        reserved = DeletedUsername(username=u.username)
        db.session.add(reserved)
    except Exception:
        db.session.rollback()
        reserved = DeletedUsername.query.filter_by(username=u.username).first()
        if not reserved:
            reserved = None

    db.session.delete(u)
    db.session.commit()

    flash("Deleted employee and all their data.", "success")
    return redirect(url_for("admin_employees"))


# ---------------- EXPORT — REWRITTEN WITHOUT PANDAS ----------------
from openpyxl import Workbook

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

        query = query.filter(WorkItem.work_date >= start_d, WorkItem.work_date <= end_d).order_by(WorkItem.work_date.asc())
        rows = query.all()

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



