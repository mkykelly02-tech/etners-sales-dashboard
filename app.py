import os
from datetime import date
from functools import wraps

import psycopg2
import psycopg2.extras
from dotenv import load_dotenv
from flask import Flask, abort, g, redirect, render_template, request, session, url_for
from werkzeug.security import check_password_hash, generate_password_hash

load_dotenv()

DATABASE_URL = os.environ["DATABASE_URL"]
STAGES = ["협상중", "계약완료", "진행중", "완료"]

app = Flask(__name__)
app.secret_key = os.environ["SECRET_KEY"]


def get_db():
    if "db" not in g:
        g.db = psycopg2.connect(DATABASE_URL, cursor_factory=psycopg2.extras.RealDictCursor)
    return g.db


@app.teardown_appcontext
def close_db(exception=None):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def init_db():
    conn = psycopg2.connect(DATABASE_URL)
    try:
        with conn, conn.cursor() as cur:
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS users (
                    id SERIAL PRIMARY KEY,
                    username TEXT NOT NULL UNIQUE,
                    password_hash TEXT NOT NULL,
                    display_name TEXT NOT NULL,
                    department TEXT NOT NULL DEFAULT '미지정',
                    role TEXT NOT NULL DEFAULT 'user' CHECK (role IN ('admin','user')),
                    created_at TIMESTAMPTZ NOT NULL DEFAULT now()
                )
                """
            )
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS contracts (
                    id SERIAL PRIMARY KEY,
                    title TEXT NOT NULL,
                    client TEXT NOT NULL,
                    amount NUMERIC(14,0) NOT NULL,
                    contract_date DATE NOT NULL,
                    stage TEXT NOT NULL DEFAULT '협상중' CHECK (stage IN ('협상중','계약완료','진행중','완료')),
                    owner_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
                    created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
                    updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
                )
                """
            )
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS targets (
                    id SERIAL PRIMARY KEY,
                    year_month TEXT NOT NULL,
                    department TEXT,
                    target_amount NUMERIC(14,0) NOT NULL,
                    UNIQUE (year_month, department)
                )
                """
            )
    finally:
        conn.close()


# ---------- auth helpers ----------

def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if "user_id" not in session:
            return redirect(url_for("login"))
        return view(*args, **kwargs)

    return wrapped


def admin_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if "user_id" not in session:
            return redirect(url_for("login"))
        if session.get("role") != "admin":
            abort(403)
        return view(*args, **kwargs)

    return wrapped


def current_user_dict():
    return {
        "id": session.get("user_id"),
        "display_name": session.get("display_name"),
        "department": session.get("department"),
        "role": session.get("role"),
    }


# ---------- auth routes ----------

@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "GET":
        return render_template("register.html")

    username = request.form.get("username", "").strip()
    password = request.form.get("password", "")
    display_name = request.form.get("display_name", "").strip()
    department = request.form.get("department", "").strip() or "미지정"

    if not username or not password or not display_name:
        return render_template("register.html", error="아이디, 비밀번호, 이름을 모두 입력하세요.")

    db = get_db()
    with db.cursor() as cur:
        cur.execute("SELECT id FROM users WHERE username = %s", (username,))
        if cur.fetchone():
            return render_template("register.html", error="이미 사용 중인 아이디입니다.")

        cur.execute("SELECT COUNT(*) AS c FROM users")
        is_first_user = cur.fetchone()["c"] == 0
        role = "admin" if is_first_user else "user"

        cur.execute(
            """
            INSERT INTO users (username, password_hash, display_name, department, role)
            VALUES (%s, %s, %s, %s, %s) RETURNING id
            """,
            (username, generate_password_hash(password), display_name, department, role),
        )
        user_id = cur.fetchone()["id"]
    db.commit()

    session["user_id"] = user_id
    session["display_name"] = display_name
    session["department"] = department
    session["role"] = role
    return redirect(url_for("dashboard"))


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "GET":
        return render_template("login.html")

    username = request.form.get("username", "").strip()
    password = request.form.get("password", "")

    db = get_db()
    with db.cursor() as cur:
        cur.execute(
            "SELECT id, password_hash, display_name, department, role FROM users WHERE username = %s",
            (username,),
        )
        user = cur.fetchone()

    if not user or not check_password_hash(user["password_hash"], password):
        return render_template("login.html", error="아이디 또는 비밀번호가 올바르지 않습니다.")

    session["user_id"] = user["id"]
    session["display_name"] = user["display_name"]
    session["department"] = user["department"]
    session["role"] = user["role"]
    return redirect(url_for("dashboard"))


@app.route("/logout", methods=["POST"])
def logout():
    session.clear()
    return redirect(url_for("login"))


# ---------- dashboard ----------

@app.route("/")
@login_required
def dashboard():
    db = get_db()
    this_month = date.today().strftime("%Y-%m")

    with db.cursor() as cur:
        if session["role"] == "admin":
            # company-wide monthly trend (last 6 months)
            cur.execute(
                """
                SELECT to_char(contract_date, 'YYYY-MM') AS ym, SUM(amount)::float AS total
                FROM contracts
                WHERE contract_date >= (CURRENT_DATE - INTERVAL '6 months')
                GROUP BY ym ORDER BY ym
                """
            )
            monthly_trend = cur.fetchall()

            cur.execute(
                """
                SELECT u.display_name, u.department, COALESCE(SUM(c.amount), 0)::float AS total
                FROM users u
                LEFT JOIN contracts c ON c.owner_id = u.id
                GROUP BY u.id, u.display_name, u.department
                ORDER BY total DESC
                """
            )
            ranking = cur.fetchall()

            cur.execute(
                "SELECT stage, COUNT(*) AS cnt, COALESCE(SUM(amount),0)::float AS total FROM contracts GROUP BY stage"
            )
            stage_rows = cur.fetchall()

            cur.execute(
                "SELECT COALESCE(SUM(amount), 0)::float AS total FROM contracts WHERE to_char(contract_date, 'YYYY-MM') = %s",
                (this_month,),
            )
            month_total = cur.fetchone()["total"]

            cur.execute(
                "SELECT target_amount FROM targets WHERE year_month = %s AND department IS NULL",
                (this_month,),
            )
            target_row = cur.fetchone()
            target_amount = target_row["target_amount"] if target_row else None

            cur.execute(
                """
                SELECT c.*, u.display_name AS owner_name
                FROM contracts c JOIN users u ON u.id = c.owner_id
                ORDER BY c.created_at DESC LIMIT 8
                """
            )
            recent = cur.fetchall()
        else:
            cur.execute(
                """
                SELECT to_char(contract_date, 'YYYY-MM') AS ym, SUM(amount)::float AS total
                FROM contracts
                WHERE owner_id = %s AND contract_date >= (CURRENT_DATE - INTERVAL '6 months')
                GROUP BY ym ORDER BY ym
                """,
                (session["user_id"],),
            )
            monthly_trend = cur.fetchall()
            ranking = []

            cur.execute(
                "SELECT stage, COUNT(*) AS cnt, COALESCE(SUM(amount),0)::float AS total FROM contracts WHERE owner_id = %s GROUP BY stage",
                (session["user_id"],),
            )
            stage_rows = cur.fetchall()

            cur.execute(
                "SELECT COALESCE(SUM(amount), 0)::float AS total FROM contracts WHERE owner_id = %s AND to_char(contract_date, 'YYYY-MM') = %s",
                (session["user_id"], this_month),
            )
            month_total = cur.fetchone()["total"]

            cur.execute(
                "SELECT target_amount FROM targets WHERE year_month = %s AND department = %s",
                (this_month, session["department"]),
            )
            target_row = cur.fetchone()
            target_amount = target_row["target_amount"] if target_row else None

            cur.execute(
                "SELECT * FROM contracts WHERE owner_id = %s ORDER BY created_at DESC LIMIT 8",
                (session["user_id"],),
            )
            recent = cur.fetchall()

    stage_map = {row["stage"]: {"cnt": row["cnt"], "total": row["total"]} for row in stage_rows}
    stage_data = [stage_map.get(s, {"cnt": 0, "total": 0}) for s in STAGES]

    achievement_pct = None
    if target_amount and float(target_amount) > 0:
        achievement_pct = round(float(month_total) / float(target_amount) * 100, 1)

    return render_template(
        "dashboard.html",
        user=current_user_dict(),
        this_month=this_month,
        month_total=month_total,
        target_amount=target_amount,
        achievement_pct=achievement_pct,
        monthly_trend=monthly_trend,
        ranking=ranking,
        stages=STAGES,
        stage_data=stage_data,
        recent=recent,
    )


# ---------- contracts CRUD ----------

@app.route("/contracts")
@login_required
def contracts_list():
    db = get_db()
    with db.cursor() as cur:
        if session["role"] == "admin":
            cur.execute(
                """
                SELECT c.*, u.display_name AS owner_name, u.department AS owner_department
                FROM contracts c JOIN users u ON u.id = c.owner_id
                ORDER BY c.contract_date DESC
                """
            )
        else:
            cur.execute(
                """
                SELECT c.*, u.display_name AS owner_name, u.department AS owner_department
                FROM contracts c JOIN users u ON u.id = c.owner_id
                WHERE c.owner_id = %s
                ORDER BY c.contract_date DESC
                """,
                (session["user_id"],),
            )
        contracts = cur.fetchall()
    return render_template("contracts_list.html", contracts=contracts, user=current_user_dict())


def _contract_form_context(error=None, contract=None):
    return {
        "user": current_user_dict(),
        "stages": STAGES,
        "error": error,
        "contract": contract,
        "today": date.today().isoformat(),
    }


@app.route("/contracts/new", methods=["GET", "POST"])
@login_required
def contract_new():
    if request.method == "GET":
        return render_template("contract_form.html", **_contract_form_context())

    title = request.form.get("title", "").strip()
    client = request.form.get("client", "").strip()
    amount = request.form.get("amount", "").strip()
    contract_date = request.form.get("contract_date", "").strip()
    stage = request.form.get("stage", STAGES[0])

    if not title or not client or not amount or not contract_date:
        return render_template("contract_form.html", **_contract_form_context(error="필수 항목을 모두 입력하세요."))

    db = get_db()
    with db.cursor() as cur:
        cur.execute(
            """
            INSERT INTO contracts (title, client, amount, contract_date, stage, owner_id)
            VALUES (%s, %s, %s, %s, %s, %s)
            """,
            (title, client, amount, contract_date, stage, session["user_id"]),
        )
    db.commit()
    return redirect(url_for("contracts_list"))


def _get_owned_contract(contract_id):
    db = get_db()
    with db.cursor() as cur:
        cur.execute("SELECT * FROM contracts WHERE id = %s", (contract_id,))
        contract = cur.fetchone()
    if not contract:
        abort(404)
    if session["role"] != "admin" and contract["owner_id"] != session["user_id"]:
        abort(403)
    return contract


@app.route("/contracts/<int:contract_id>/edit", methods=["GET", "POST"])
@login_required
def contract_edit(contract_id):
    contract = _get_owned_contract(contract_id)

    if request.method == "GET":
        return render_template("contract_form.html", **_contract_form_context(contract=contract))

    title = request.form.get("title", "").strip()
    client = request.form.get("client", "").strip()
    amount = request.form.get("amount", "").strip()
    contract_date = request.form.get("contract_date", "").strip()
    stage = request.form.get("stage", STAGES[0])

    if not title or not client or not amount or not contract_date:
        return render_template(
            "contract_form.html", **_contract_form_context(error="필수 항목을 모두 입력하세요.", contract=contract)
        )

    db = get_db()
    with db.cursor() as cur:
        cur.execute(
            """
            UPDATE contracts SET title=%s, client=%s, amount=%s, contract_date=%s, stage=%s, updated_at=now()
            WHERE id = %s
            """,
            (title, client, amount, contract_date, stage, contract_id),
        )
    db.commit()
    return redirect(url_for("contracts_list"))


@app.route("/contracts/<int:contract_id>/delete", methods=["POST"])
@login_required
def contract_delete(contract_id):
    _get_owned_contract(contract_id)
    db = get_db()
    with db.cursor() as cur:
        cur.execute("DELETE FROM contracts WHERE id = %s", (contract_id,))
    db.commit()
    return redirect(url_for("contracts_list"))


# ---------- admin: users & targets ----------

@app.route("/admin/users")
@admin_required
def admin_users():
    db = get_db()
    with db.cursor() as cur:
        cur.execute("SELECT id, username, display_name, department, role FROM users ORDER BY id")
        users = cur.fetchall()
    return render_template("admin_users.html", users=users, user=current_user_dict())


@app.route("/admin/users/<int:user_id>/role", methods=["POST"])
@admin_required
def admin_user_role(user_id):
    new_role = request.form.get("role")
    if new_role not in ("admin", "user"):
        abort(400)
    db = get_db()
    with db.cursor() as cur:
        cur.execute("UPDATE users SET role = %s WHERE id = %s", (new_role, user_id))
    db.commit()
    return redirect(url_for("admin_users"))


@app.route("/admin/targets", methods=["GET", "POST"])
@admin_required
def admin_targets():
    db = get_db()
    if request.method == "POST":
        year_month = request.form.get("year_month", "").strip()
        department = request.form.get("department", "").strip() or None
        target_amount = request.form.get("target_amount", "").strip()
        if year_month and target_amount:
            with db.cursor() as cur:
                cur.execute(
                    """
                    INSERT INTO targets (year_month, department, target_amount)
                    VALUES (%s, %s, %s)
                    ON CONFLICT (year_month, department)
                    DO UPDATE SET target_amount = EXCLUDED.target_amount
                    """,
                    (year_month, department, target_amount),
                )
            db.commit()
        return redirect(url_for("admin_targets"))

    with db.cursor() as cur:
        cur.execute("SELECT * FROM targets ORDER BY year_month DESC, department NULLS FIRST")
        targets = cur.fetchall()
        cur.execute("SELECT DISTINCT department FROM users ORDER BY department")
        departments = [r["department"] for r in cur.fetchall()]
    return render_template(
        "admin_targets.html", targets=targets, departments=departments, user=current_user_dict(), today=date.today().isoformat()
    )


init_db()

if __name__ == "__main__":
    app.run(debug=True)
