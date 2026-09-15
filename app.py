import base64
import json
import mimetypes
import os
import re
import uuid
from datetime import date
from functools import wraps
from io import BytesIO
from urllib.parse import quote

import psycopg2
import psycopg2.extras
import requests
from dotenv import load_dotenv
from flask import Flask, Response, abort, g, redirect, render_template, request, session, url_for
from openpyxl import Workbook
from werkzeug.security import check_password_hash, generate_password_hash

load_dotenv()

DATABASE_URL = os.environ["DATABASE_URL"]


def _clean_env(value):
    """Strip whitespace and drop any non-ASCII byte.

    Env vars pasted through a web dashboard can pick up invisible characters
    (smart quotes, NBSPs, stray newlines) that are invalid in an HTTP header
    and make `requests` blow up with a latin-1 UnicodeEncodeError. A Supabase
    URL/JWT never legitimately contains non-ASCII characters, so it's safe to
    just filter them out rather than fail.
    """
    if not value:
        return None
    cleaned = "".join(ch for ch in value if ord(ch) < 128).strip()
    return cleaned or None


# Optional: only the 계약 서류 업로드 기능 needs these. Read lazily (not at import
# time) so a missing value can't take down login/contracts/dashboard for everyone.
SUPABASE_URL = _clean_env(os.environ.get("SUPABASE_URL"))
if SUPABASE_URL:
    SUPABASE_URL = SUPABASE_URL.rstrip("/")
SUPABASE_SERVICE_KEY = _clean_env(os.environ.get("SUPABASE_SERVICE_KEY"))
STORAGE_BUCKET = "contract-files"

# Optional: only the AI 문서대조 기능 needs this. Without it, uploads just skip
# straight to match_status = '미확인' instead of failing.
ANTHROPIC_API_KEY = _clean_env(os.environ.get("ANTHROPIC_API_KEY"))
ANTHROPIC_MODEL = "claude-sonnet-5"
EXTRACTABLE_MIME_TYPES = {
    "application/pdf": "document",
    "image/png": "image",
    "image/jpeg": "image",
    "image/jpg": "image",
}

STAGES = ["협상중", "계약완료", "진행중", "완료"]
DOC_TYPES = ["계약서", "세금계산서", "기타"]


class StorageNotConfigured(RuntimeError):
    pass


app = Flask(__name__)
app.secret_key = os.environ["SECRET_KEY"]


@app.errorhandler(StorageNotConfigured)
def handle_storage_not_configured(exc):
    return str(exc), 503


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
                    start_date DATE NOT NULL,
                    end_date DATE NOT NULL,
                    stage TEXT NOT NULL DEFAULT '협상중' CHECK (stage IN ('협상중','계약완료','진행중','완료')),
                    owner_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
                    created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
                    updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
                )
                """
            )
            # migrate older schema: contract_date -> start_date, add end_date
            cur.execute("SELECT column_name FROM information_schema.columns WHERE table_name = 'contracts'")
            cols = {row[0] for row in cur.fetchall()}
            if "contract_date" in cols and "start_date" not in cols:
                cur.execute("ALTER TABLE contracts RENAME COLUMN contract_date TO start_date")
            if "end_date" not in cols:
                cur.execute("ALTER TABLE contracts ADD COLUMN end_date DATE NOT NULL DEFAULT CURRENT_DATE")

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
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS contract_files (
                    id SERIAL PRIMARY KEY,
                    contract_id INTEGER NOT NULL REFERENCES contracts(id) ON DELETE CASCADE,
                    doc_type TEXT NOT NULL CHECK (doc_type IN ('계약서','세금계산서','기타')),
                    file_name TEXT NOT NULL,
                    storage_path TEXT NOT NULL,
                    mime_type TEXT,
                    uploaded_by INTEGER REFERENCES users(id),
                    uploaded_at TIMESTAMPTZ NOT NULL DEFAULT now()
                )
                """
            )
            # AI 문서대조 결과 (금액/기간을 문서에서 읽어 계약 정보와 비교)
            cur.execute("ALTER TABLE contract_files ADD COLUMN IF NOT EXISTS extracted_amount NUMERIC(14,0)")
            cur.execute("ALTER TABLE contract_files ADD COLUMN IF NOT EXISTS extracted_start_date DATE")
            cur.execute("ALTER TABLE contract_files ADD COLUMN IF NOT EXISTS extracted_end_date DATE")
            cur.execute("ALTER TABLE contract_files ADD COLUMN IF NOT EXISTS match_status TEXT")
            cur.execute("ALTER TABLE contract_files ADD COLUMN IF NOT EXISTS match_notes TEXT")
    finally:
        conn.close()


# ---------- Supabase Storage helpers ----------

def _require_storage_config():
    if not SUPABASE_URL or not SUPABASE_SERVICE_KEY:
        raise StorageNotConfigured(
            "SUPABASE_URL / SUPABASE_SERVICE_KEY 환경변수가 설정되지 않아 서류 업로드 기능을 쓸 수 없습니다."
        )


def _storage_headers(content_type=None):
    headers = {
        "Authorization": f"Bearer {SUPABASE_SERVICE_KEY}",
        "apikey": SUPABASE_SERVICE_KEY,
    }
    if content_type:
        headers["Content-Type"] = content_type
    return headers


def storage_upload(path, file_bytes, content_type):
    _require_storage_config()
    url = f"{SUPABASE_URL}/storage/v1/object/{STORAGE_BUCKET}/{path}"
    resp = requests.post(url, headers=_storage_headers(content_type or "application/octet-stream"), data=file_bytes)
    if not resp.ok:
        app.logger.error("storage_upload failed: %s %s -> %s", resp.status_code, url, resp.text[:500])
    resp.raise_for_status()


def storage_download(path):
    _require_storage_config()
    url = f"{SUPABASE_URL}/storage/v1/object/{STORAGE_BUCKET}/{path}"
    resp = requests.get(url, headers=_storage_headers())
    resp.raise_for_status()
    return resp.content


def storage_delete(path):
    _require_storage_config()
    url = f"{SUPABASE_URL}/storage/v1/object/{STORAGE_BUCKET}/{path}"
    requests.delete(url, headers=_storage_headers())


# ---------- AI 문서대조 (계약서/세금계산서 금액·기간 자동 인식) ----------

def extract_document_fields(file_bytes, mime_type):
    """Ask Claude to read a 계약서/세금계산서 and pull out amount + period.

    Returns None when extraction wasn't even attempted (no API key, or a file
    type vision can't read, e.g. .docx/.hwp) — that's different from having
    tried and found nothing, which returns a dict with the fields left null.
    """
    if not ANTHROPIC_API_KEY:
        return None
    block_type = EXTRACTABLE_MIME_TYPES.get(mime_type)
    if not block_type:
        return None

    content_block = {
        "type": block_type,
        "source": {
            "type": "base64",
            "media_type": mime_type,
            "data": base64.b64encode(file_bytes).decode("ascii"),
        },
    }
    prompt = (
        "이 문서는 회사 계약서 또는 세금계산서입니다. 아래 JSON 형식으로만 답하세요 (다른 설명 없이):\n"
        '{"amount": 숫자또는null, "start_date": "YYYY-MM-DD"또는null, '
        '"end_date": "YYYY-MM-DD"또는null, "note": "한 줄 설명"}\n'
        "amount는 부가세 포함 총 금액을 원화 숫자만으로 적으세요(콤마 없이). "
        "계약기간이 명시되어 있지 않고 발행일/작성일만 있다면 start_date에 그 날짜를 넣고 "
        "end_date는 null로 두세요. 찾을 수 없는 값은 null로 두세요."
    )
    try:
        resp = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "x-api-key": ANTHROPIC_API_KEY,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json",
            },
            json={
                "model": ANTHROPIC_MODEL,
                "max_tokens": 500,
                "messages": [{"role": "user", "content": [content_block, {"type": "text", "text": prompt}]}],
            },
            timeout=45,
        )
        resp.raise_for_status()
        text = resp.json()["content"][0]["text"]
        match = re.search(r"\{.*\}", text, re.DOTALL)
        data = json.loads(match.group(0)) if match else {}
        return {
            "amount": data.get("amount"),
            "start_date": data.get("start_date"),
            "end_date": data.get("end_date"),
            "note": data.get("note"),
        }
    except Exception as exc:
        app.logger.error("extract_document_fields failed: %s", exc)
        return {"amount": None, "start_date": None, "end_date": None, "note": f"인식 실패: {exc}"}


def compute_match_status(extracted, contract):
    """Compare AI-extracted fields against the contract record.

    Returns (match_status, match_notes). Deliberately conservative: only
    flags a mismatch when something extracted actively contradicts the
    contract, never when the document simply didn't state a field.
    """
    if extracted is None:
        return "미확인", "ANTHROPIC_API_KEY가 설정되지 않아 자동 인식을 건너뛰었습니다."

    notes = []
    amount = extracted.get("amount")
    if amount is not None:
        try:
            if round(float(amount)) != round(float(contract["amount"])):
                notes.append(f"인식된 금액 {int(round(float(amount))):,}원 ≠ 계약금액 {int(contract['amount']):,}원")
        except (TypeError, ValueError):
            pass

    contract_start = str(contract["start_date"])
    contract_end = str(contract["end_date"])
    for field, label in (("start_date", "인식된 날짜"), ("end_date", "인식된 종료일")):
        value = extracted.get(field)
        if value and not (contract_start <= value <= contract_end):
            notes.append(f"{label} {value}가 계약기간({contract_start}~{contract_end}) 밖입니다")

    if notes:
        return "불일치 의심", " / ".join(notes)
    if amount is None and not extracted.get("start_date") and not extracted.get("end_date"):
        return "인식불가", extracted.get("note") or "문서에서 금액·기간을 찾지 못했습니다."
    return "일치", extracted.get("note")


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
                SELECT to_char(start_date, 'YYYY-MM') AS ym, SUM(amount)::float AS total
                FROM contracts
                WHERE start_date >= (CURRENT_DATE - INTERVAL '6 months')
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
                "SELECT COALESCE(SUM(amount), 0)::float AS total FROM contracts WHERE to_char(start_date, 'YYYY-MM') = %s",
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
                SELECT to_char(start_date, 'YYYY-MM') AS ym, SUM(amount)::float AS total
                FROM contracts
                WHERE owner_id = %s AND start_date >= (CURRENT_DATE - INTERVAL '6 months')
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
                "SELECT COALESCE(SUM(amount), 0)::float AS total FROM contracts WHERE owner_id = %s AND to_char(start_date, 'YYYY-MM') = %s",
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
    client_filter = request.args.get("client", "")

    conditions = []
    params = []
    if session["role"] != "admin":
        conditions.append("c.owner_id = %s")
        params.append(session["user_id"])
    if client_filter:
        conditions.append("c.client = %s")
        params.append(client_filter)
    where_sql = ("WHERE " + " AND ".join(conditions)) if conditions else ""

    db = get_db()
    with db.cursor() as cur:
        cur.execute(
            f"""
            SELECT c.*, u.display_name AS owner_name, u.department AS owner_department
            FROM contracts c JOIN users u ON u.id = c.owner_id
            {where_sql}
            ORDER BY c.start_date DESC
            """,
            params,
        )
        contracts = cur.fetchall()
    return render_template(
        "contracts_list.html", contracts=contracts, user=current_user_dict(), client_filter=client_filter
    )


def _contract_form_context(error=None, contract=None, files=None):
    return {
        "user": current_user_dict(),
        "stages": STAGES,
        "doc_types": DOC_TYPES,
        "error": error,
        "contract": contract,
        "files": files or [],
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
    start_date = request.form.get("start_date", "").strip()
    end_date = request.form.get("end_date", "").strip()
    stage = request.form.get("stage", STAGES[0])

    if not title or not client or not amount or not start_date or not end_date:
        return render_template("contract_form.html", **_contract_form_context(error="필수 항목을 모두 입력하세요."))
    if end_date < start_date:
        return render_template("contract_form.html", **_contract_form_context(error="계약 종료일은 시작일보다 빠를 수 없습니다."))

    db = get_db()
    with db.cursor() as cur:
        cur.execute(
            """
            INSERT INTO contracts (title, client, amount, start_date, end_date, stage, owner_id)
            VALUES (%s, %s, %s, %s, %s, %s, %s) RETURNING id
            """,
            (title, client, amount, start_date, end_date, stage, session["user_id"]),
        )
        new_id = cur.fetchone()["id"]
    db.commit()
    return redirect(url_for("contract_edit", contract_id=new_id))


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
        db = get_db()
        with db.cursor() as cur:
            cur.execute(
                "SELECT * FROM contract_files WHERE contract_id = %s ORDER BY uploaded_at DESC", (contract_id,)
            )
            files = cur.fetchall()
        return render_template("contract_form.html", **_contract_form_context(contract=contract, files=files))

    title = request.form.get("title", "").strip()
    client = request.form.get("client", "").strip()
    amount = request.form.get("amount", "").strip()
    start_date = request.form.get("start_date", "").strip()
    end_date = request.form.get("end_date", "").strip()
    stage = request.form.get("stage", STAGES[0])

    if not title or not client or not amount or not start_date or not end_date:
        return render_template(
            "contract_form.html", **_contract_form_context(error="필수 항목을 모두 입력하세요.", contract=contract)
        )
    if end_date < start_date:
        return render_template(
            "contract_form.html",
            **_contract_form_context(error="계약 종료일은 시작일보다 빠를 수 없습니다.", contract=contract),
        )

    db = get_db()
    with db.cursor() as cur:
        cur.execute(
            """
            UPDATE contracts SET title=%s, client=%s, amount=%s, start_date=%s, end_date=%s, stage=%s, updated_at=now()
            WHERE id = %s
            """,
            (title, client, amount, start_date, end_date, stage, contract_id),
        )
    db.commit()
    return redirect(url_for("contracts_list"))


@app.route("/contracts/<int:contract_id>/delete", methods=["POST"])
@login_required
def contract_delete(contract_id):
    contract = _get_owned_contract(contract_id)
    db = get_db()
    with db.cursor() as cur:
        cur.execute("SELECT storage_path FROM contract_files WHERE contract_id = %s", (contract_id,))
        file_paths = [row["storage_path"] for row in cur.fetchall()]
        cur.execute("DELETE FROM contracts WHERE id = %s", (contract_id,))
    db.commit()
    for path in file_paths:
        storage_delete(path)
    return redirect(url_for("contracts_list"))


# ---------- contract files (수행실적 서류) ----------

@app.route("/contracts/<int:contract_id>/files", methods=["POST"])
@login_required
def contract_file_upload(contract_id):
    contract = _get_owned_contract(contract_id)

    doc_type = request.form.get("doc_type", "").strip()
    upload = request.files.get("file")

    if doc_type not in DOC_TYPES or not upload or not upload.filename:
        abort(400)

    original_name = upload.filename
    content_type = upload.mimetype or mimetypes.guess_type(original_name)[0] or "application/octet-stream"
    file_bytes = upload.read()
    # Keep the real (possibly Korean) filename only in the DB for display; the storage
    # object key must stay ASCII-safe. secure_filename() drops non-ASCII text entirely
    # (and can even swallow the dot for an all-Korean name), so pull the extension from
    # the original name directly instead of routing it through secure_filename().
    raw_ext = os.path.splitext(original_name)[1]
    safe_ext = "".join(ch for ch in raw_ext if ch.isalnum() or ch == ".")[:10]
    storage_path = f"{contract_id}/{uuid.uuid4().hex}{safe_ext}"

    storage_upload(storage_path, file_bytes, content_type)

    extracted = None
    match_status, match_notes = None, None
    if doc_type in ("계약서", "세금계산서"):
        extracted = extract_document_fields(file_bytes, content_type)
        match_status, match_notes = compute_match_status(extracted, contract)

    db = get_db()
    with db.cursor() as cur:
        cur.execute(
            """
            INSERT INTO contract_files (
                contract_id, doc_type, file_name, storage_path, mime_type, uploaded_by,
                extracted_amount, extracted_start_date, extracted_end_date, match_status, match_notes
            )
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """,
            (
                contract_id, doc_type, original_name, storage_path, content_type, session["user_id"],
                extracted.get("amount") if extracted else None,
                extracted.get("start_date") if extracted else None,
                extracted.get("end_date") if extracted else None,
                match_status, match_notes,
            ),
        )
    db.commit()
    return redirect(url_for("contract_edit", contract_id=contract_id))


def _get_owned_file(contract_id, file_id):
    contract = _get_owned_contract(contract_id)
    db = get_db()
    with db.cursor() as cur:
        cur.execute("SELECT * FROM contract_files WHERE id = %s AND contract_id = %s", (file_id, contract_id))
        file_row = cur.fetchone()
    if not file_row:
        abort(404)
    return contract, file_row


@app.route("/contracts/<int:contract_id>/files/<int:file_id>")
@login_required
def contract_file_download(contract_id, file_id):
    _, file_row = _get_owned_file(contract_id, file_id)
    data = storage_download(file_row["storage_path"])
    filename = file_row["file_name"]
    ascii_fallback = filename.encode("ascii", "ignore").decode("ascii") or "download"
    disposition = f"attachment; filename=\"{ascii_fallback}\"; filename*=UTF-8''{quote(filename)}"
    return Response(
        data,
        mimetype=file_row["mime_type"] or "application/octet-stream",
        headers={"Content-Disposition": disposition},
    )


@app.route("/contracts/<int:contract_id>/files/<int:file_id>/delete", methods=["POST"])
@login_required
def contract_file_delete(contract_id, file_id):
    _, file_row = _get_owned_file(contract_id, file_id)
    db = get_db()
    with db.cursor() as cur:
        cur.execute("DELETE FROM contract_files WHERE id = %s", (file_id,))
    db.commit()
    storage_delete(file_row["storage_path"])
    return redirect(url_for("contract_edit", contract_id=contract_id))


# ---------- 서류함 (전체 첨부 서류 모아보기) ----------

@app.route("/documents")
@login_required
def documents():
    doc_type = request.args.get("doc_type", "")
    owner_id = request.args.get("owner_id", "")
    date_from = request.args.get("date_from", "")
    date_to = request.args.get("date_to", "")

    conditions = []
    params = []
    if session["role"] != "admin":
        conditions.append("c.owner_id = %s")
        params.append(session["user_id"])
    elif owner_id:
        conditions.append("c.owner_id = %s")
        params.append(owner_id)
    if doc_type:
        conditions.append("f.doc_type = %s")
        params.append(doc_type)
    if date_from:
        conditions.append("c.start_date >= %s")
        params.append(date_from)
    if date_to:
        conditions.append("c.start_date <= %s")
        params.append(date_to)

    where_sql = ("WHERE " + " AND ".join(conditions)) if conditions else ""
    db = get_db()
    with db.cursor() as cur:
        cur.execute(
            f"""
            SELECT f.*, c.title AS contract_title, c.client, c.start_date, c.end_date,
                   u.display_name AS owner_name, u.department AS owner_department
            FROM contract_files f
            JOIN contracts c ON c.id = f.contract_id
            JOIN users u ON u.id = c.owner_id
            {where_sql}
            ORDER BY f.uploaded_at DESC
            """,
            params,
        )
        files = cur.fetchall()

        owners = []
        if session["role"] == "admin":
            cur.execute("SELECT id, display_name FROM users ORDER BY display_name")
            owners = cur.fetchall()

    return render_template(
        "documents.html",
        files=files,
        doc_types=DOC_TYPES,
        owners=owners,
        filters={"doc_type": doc_type, "owner_id": owner_id, "date_from": date_from, "date_to": date_to},
        user=current_user_dict(),
    )


# ---------- 거래처 관리 ----------

@app.route("/clients")
@login_required
def clients():
    db = get_db()
    with db.cursor() as cur:
        if session["role"] == "admin":
            cur.execute(
                """
                SELECT client, COUNT(*) AS cnt, SUM(amount)::float AS total, MAX(start_date) AS latest
                FROM contracts GROUP BY client ORDER BY total DESC
                """
            )
        else:
            cur.execute(
                """
                SELECT client, COUNT(*) AS cnt, SUM(amount)::float AS total, MAX(start_date) AS latest
                FROM contracts WHERE owner_id = %s GROUP BY client ORDER BY total DESC
                """,
                (session["user_id"],),
            )
        client_rows = cur.fetchall()
    return render_template("clients.html", clients=client_rows, user=current_user_dict())


# ---------- 리포트 / 엑셀 내보내기 ----------

@app.route("/reports")
@login_required
def reports():
    return render_template("reports.html", user=current_user_dict(), today=date.today().isoformat())


@app.route("/reports/export")
@login_required
def reports_export():
    start = request.args.get("start", "")
    end = request.args.get("end", "")

    conditions = []
    params = []
    if session["role"] != "admin":
        conditions.append("c.owner_id = %s")
        params.append(session["user_id"])
    if start:
        conditions.append("c.start_date >= %s")
        params.append(start)
    if end:
        conditions.append("c.start_date <= %s")
        params.append(end)
    where_sql = ("WHERE " + " AND ".join(conditions)) if conditions else ""

    db = get_db()
    with db.cursor() as cur:
        cur.execute(
            f"""
            SELECT c.title, c.client, c.amount, c.start_date, c.end_date, c.stage,
                   u.display_name AS owner_name, u.department AS owner_department
            FROM contracts c JOIN users u ON u.id = c.owner_id
            {where_sql}
            ORDER BY c.start_date
            """,
            params,
        )
        rows = cur.fetchall()

    wb = Workbook()
    ws = wb.active
    ws.title = "실적리포트"
    headers = ["계약명", "거래처", "금액", "시작일", "종료일", "단계", "담당자", "부서"]
    ws.append(headers)
    for r in rows:
        ws.append(
            [
                r["title"], r["client"], float(r["amount"]),
                r["start_date"].isoformat(), r["end_date"].isoformat(),
                r["stage"], r["owner_name"], r["owner_department"],
            ]
        )
    widths = [24, 16, 14, 12, 12, 10, 12, 14]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + i)].width = w

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    filename = f"실적리포트_{start or '전체'}_{end or '전체'}.xlsx"
    return Response(
        buf.read(),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(filename)}"},
    )


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
