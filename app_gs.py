import streamlit as st
import hmac
import hashlib
import pandas as pd
import calendar
from datetime import datetime
from urllib.parse import urlencode
import re
import csv
import io
import os
import uuid
import gspread
from google.oauth2.service_account import Credentials
import streamlit.components.v1 as components
from streamlit_autorefresh import st_autorefresh


st.set_page_config(layout="wide")

TOKEN_EXPIRE_SECONDS = 60 * 60 * 12

def get_auth_users():
    return st.secrets["auth"]["users"]

def get_token_secret() -> str:
    auth_section = st.secrets["auth"]
    if "token_secret" in auth_section:
        return str(auth_section["token_secret"])
    seed = "|".join(
        f"{user.get('username', '')}:{user.get('password', '')}"
        for user in get_auth_users()
    )
    return hashlib.sha256((seed + "|ir_schedule_token").encode()).hexdigest()

def make_login_token(username: str, expires_in: int = TOKEN_EXPIRE_SECONDS) -> str:
    expires_at = int(datetime.now().timestamp()) + int(expires_in)
    payload = f"{username}|{expires_at}"
    signature = hmac.new(
        get_token_secret().encode(),
        payload.encode(),
        hashlib.sha256,
    ).hexdigest()
    return f"{username}.{expires_at}.{signature}"

def find_saved_duty_image(year, month):
    base_dir = "duty_images"
    if not os.path.exists(base_dir):
        return None

    for file in os.listdir(base_dir):
        if file.startswith(f"{year}_{month}"):
            return os.path.join(base_dir, file)
    return None


def get_duty_image_path(year, month, ext):
    base_dir = "duty_images"
    os.makedirs(base_dir, exist_ok=True)
    return os.path.join(base_dir, f"{year}_{month}.{ext}")

def verify_login_token(token: str):
    token = str(token or "").strip()
    if not token:
        return None
    try:
        username, expires_at_text, signature = token.split(".", 2)
        expires_at = int(expires_at_text)
    except Exception:
        return None

    if expires_at < int(datetime.now().timestamp()):
        return None

    payload = f"{username}|{expires_at}"
    expected_signature = hmac.new(
        get_token_secret().encode(),
        payload.encode(),
        hashlib.sha256,
    ).hexdigest()

    if not hmac.compare_digest(expected_signature, signature):
        return None

    for user in get_auth_users():
        if hmac.compare_digest(str(user.get("username", "")), str(username)):
            return str(username)
    return None

def set_login_state(username: str, token: str | None = None):
    st.session_state["logged_in"] = True
    st.session_state["username"] = str(username)
    st.session_state["auth_token"] = token or make_login_token(str(username))

def current_auth_token() -> str:
    username = str(st.session_state.get("username", "")).strip()
    token = str(st.session_state.get("auth_token", "")).strip()
    verified_username = verify_login_token(token)
    if username and verified_username == username:
        return token
    token = make_login_token(username)
    st.session_state["auth_token"] = token
    return token

def app_query_string(**params) -> str:
    query_params = {}
    token = str(st.session_state.get("auth_token", "")).strip()
    if token:
        query_params["token"] = token
    for key, value in params.items():
        if value is None or value == "":
            continue
        query_params[key] = value
    return "?" + urlencode(query_params, doseq=True) if query_params else ""

def replace_query_params(**params):
    st.query_params.clear()
    token = str(st.session_state.get("auth_token", "")).strip()
    if token:
        st.query_params["token"] = token
    for key, value in params.items():
        if value is None or value == "":
            continue
        st.query_params[key] = value

def clear_query_param(key: str):
    if key in st.query_params:
        del st.query_params[key]

def check_login():
    if "logged_in" not in st.session_state:
        st.session_state["logged_in"] = False
    if "username" not in st.session_state:
        st.session_state["username"] = ""
    if "auth_token" not in st.session_state:
        st.session_state["auth_token"] = ""

    token_user = verify_login_token(st.query_params.get("token"))
    if token_user:
        set_login_state(token_user, st.query_params.get("token"))
        return

    if st.session_state["logged_in"]:
        if not st.session_state.get("auth_token"):
            st.session_state["auth_token"] = make_login_token(st.session_state["username"])
        return

    st.title("로그인")

    username = st.text_input("ID", key="login_username")
    password = st.text_input("PW", type="password", key="login_password")

    if st.button("로그인", key="login_button"):
        matched_user = next(
            (
                user for user in get_auth_users()
                if hmac.compare_digest(str(user["username"]), str(username))
                and hmac.compare_digest(str(user["password"]), str(password))
            ),
            None
        )

        if matched_user:
            set_login_state(str(username))
            replace_query_params()
            st.rerun()
        else:
            st.error("ID 또는 PW가 올바르지 않습니다.")

    st.stop()

def render_logout():
    right1, right2 = st.columns([8.5, 1.5])
    with right2:
        if st.button(f"로그아웃", use_container_width=True, key="logout_button"):
            st.session_state["logged_in"] = False
            st.session_state["username"] = ""
            st.session_state["auth_token"] = ""
            st.query_params.clear()
            st.rerun()

check_login()
SPREADSHEET_NAME = "IR_schedule"
PROCEDURES_SHEET = "procedures"
VACATION_SHEET = "vacation_notes"

DELETE_PASSWORD = "0000"
DATA_FILE = "procedures.csv"
VACATION_FILE = "vacation_notes.csv"
DUTY_IMAGE_DIR = "duty_images"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
SPREADSHEET_NAME = "IR_schedule"
PROCEDURES_SHEET = "procedures"
VACATION_SHEET = "vacation_notes"

STATUS_PLANNED = "예정"
STATUS_CALLED = "호출"
STATUS_ARRIVED = "도착"
STATUS_INROOM = "입실"
STATUS_DONE = "시술완료"

VALID_STATUSES = [STATUS_PLANNED, STATUS_CALLED, STATUS_ARRIVED, STATUS_INROOM, STATUS_DONE]

STATUS_COLORS = {
    STATUS_PLANNED: "#9aa0a6",
    STATUS_CALLED: "#f59e0b",
    STATUS_ARRIVED: "#f15628",   
    STATUS_INROOM: "#ef4444",
    STATUS_DONE: "#38bdf8",
}

COLUMNS = [
    "id", "updated_at",
    "날짜", "순서", "등록번호", "이름", "성별", "나이",
    "병실", "시술과", "시술명", "의뢰과", "의뢰의",
    "Room", "교수", "응급", "진행상황",
    "동의서",
    "감염", "감염메모",
    "ADR", "ADR메모",
    "신기능", "Cr",
    "출혈", "PT_INR", "PLT",
    "메모"
]

VACATION_COLUMNS = ["id", "updated_at", "날짜", "메모", "잠금"]

def get_prof_options(proc_dept: str):
    mapping = {
        "IR": ["정선화", "박상영"],
        "NS": ["강정한", "허원", "심환석"],
        "NU": ["안상준", "장성화"],
    }
    return mapping.get(proc_dept, [])

def get_prof_select_options(proc_dept: str):
    return [""] + get_prof_options(proc_dept)

def empty_procedures_df():
    return pd.DataFrame(columns=COLUMNS)

def empty_vacation_df():
    return pd.DataFrame(columns=VACATION_COLUMNS)

def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")

def truthy(value) -> bool:
    return str(value).strip().lower() in ["true", "1", "y", "yes"]

def normalize_room_value(value) -> str:
    room_text = "" if value is None or pd.isna(value) else str(value).strip()
    room_map = {
        "1번방": "1",
        "2번방": "2",
        "1.0": "1",
        "2.0": "2",
        "ROOM": "",
        "Room": "",
        "nan": "",
        "None": "",
    }
    room_text = room_map.get(room_text, room_text)
    return room_text if room_text in ["", "1", "2", "H"] else ""

def sheet_enabled() -> bool:
    return "gcp_service_account" in st.secrets

@st.cache_resource
def get_gspread_client():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

@st.cache_resource
def get_workbook():
    client = get_gspread_client()
    return client.open(SPREADSHEET_NAME)

@st.cache_resource
def get_worksheet(name: str):
    wb = get_workbook()
    return wb.worksheet(name)

def get_or_create_worksheet(name: str, header: list[str]):
    wb = get_workbook()

    # 1) 있으면 바로 반환
    try:
        return wb.worksheet(name)
    except gspread.exceptions.WorksheetNotFound:
        pass

    # 2) 없으면 생성
    try:
        ws = wb.add_worksheet(title=name, rows=2000, cols=max(len(header), 20))
        ws.update("A1", [header])
        return ws
    except gspread.exceptions.APIError as e:
        # 동시에 다른 쪽에서 만든 경우 대비
        if "already exists" in str(e):
            return wb.worksheet(name)
        raise

@st.cache_resource
def get_worksheet_cached(name: str):
    wb = get_workbook()
    try:
        return wb.worksheet(name)
    except gspread.exceptions.WorksheetNotFound:
        return None
    

def record_to_row(record: dict, header: list[str]):
    row = []
    for col in header:
        value = record.get(col, "")
        if pd.isna(value):
            value = ""
        if isinstance(value, bool):
            row.append("TRUE" if value else "FALSE")
        else:
            row.append(str(value))
    return row

def sheet_records_to_df(rows: list[dict], columns: list[str]) -> pd.DataFrame:
    if not rows:
        df = pd.DataFrame(columns=columns)
    else:
        df = pd.DataFrame(rows)

    for col in columns:
        if col not in df.columns:
            df[col] = ""

    return df[columns].copy()

def normalize_procedures_df(df: pd.DataFrame) -> pd.DataFrame:
    for col in COLUMNS:
        if col not in df.columns:
            df[col] = ""

    if "id" not in df.columns:
        df["id"] = ""
    if "updated_at" not in df.columns:
        df["updated_at"] = ""

    df = df[COLUMNS].copy()
    df["id"] = df["id"].astype(str).replace(["", "nan", "None"], "")
    missing_mask = df["id"].str.strip() == ""
    if missing_mask.any():
        df.loc[missing_mask, "id"] = [str(uuid.uuid4()) for _ in range(missing_mask.sum())]

    updated_missing = df["updated_at"].astype(str).str.strip() == ""
    if updated_missing.any():
        df.loc[updated_missing, "updated_at"] = now_iso()

    df["순서"] = pd.to_numeric(df["순서"], errors="coerce").fillna(0).astype(int)
    df["나이"] = pd.to_numeric(df["나이"], errors="coerce").fillna(0).astype(int)
    for col in ["응급", "감염", "동의서", "ADR", "신기능", "출혈"]:
        df[col] = df[col].apply(truthy)

    df["진행상황"] = df["진행상황"].replace({"완료": STATUS_DONE})
    df["진행상황"] = df["진행상황"].apply(lambda x: x if x in VALID_STATUSES else STATUS_PLANNED)
    df["Room"] = df["Room"].apply(normalize_room_value)
    if "시술과" not in df.columns:
        df["시술과"] = "IR"
    return df

def normalize_vacation_df(df: pd.DataFrame) -> pd.DataFrame:
    for col in VACATION_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    df = df[VACATION_COLUMNS].copy()
    df["id"] = df["id"].astype(str).replace(["", "nan", "None"], "")
    missing_mask = df["id"].str.strip() == ""
    if missing_mask.any():
        df.loc[missing_mask, "id"] = [str(uuid.uuid4()) for _ in range(missing_mask.sum())]

    updated_missing = df["updated_at"].astype(str).str.strip() == ""
    if updated_missing.any():
        df.loc[updated_missing, "updated_at"] = now_iso()

    df["잠금"] = df["잠금"].apply(truthy)
    return df

def load_data():
    if sheet_enabled():
        ws = get_worksheet(PROCEDURES_SHEET)
        rows = ws.get_all_records()
        return normalize_procedures_df(sheet_records_to_df(rows, COLUMNS))

    if not os.path.exists(DATA_FILE):
        return empty_procedures_df()

    try:
        df = pd.read_csv(DATA_FILE, dtype=str).fillna("")
    except Exception:
        return empty_procedures_df()

    if "진단명" in df.columns:
        df = df.drop(columns=["진단명"])

    return normalize_procedures_df(df)

def save_data(df: pd.DataFrame):
    save_df = normalize_procedures_df(df.copy())
    bool_cols = ["응급", "동의서", "감염", "ADR", "신기능", "출혈"]

    for col in bool_cols:
        if col in save_df.columns:
            save_df[col] = save_df[col].apply(lambda x: True if bool(x) else False)

    if sheet_enabled():
        ws = get_worksheet(PROCEDURES_SHEET)
        ws.clear()
        ws.update("A1", [COLUMNS] + save_df[COLUMNS].astype(str).values.tolist())
        return

    save_df.to_csv(DATA_FILE, index=False, encoding="utf-8-sig")

def load_vacation_data():
    ws = get_worksheet(VACATION_SHEET)

    values = ws.get_all_values()

    # 빈 시트 대응
    if not values:
        return pd.DataFrame(columns=VACATION_COLUMNS + ["_sheet_row"])

    header = values[0]
    rows = values[1:]

    records = []
    for sheet_row_num, row in enumerate(rows, start=2):
        record = {}

        for i, col in enumerate(header):
            record[col] = row[i] if i < len(row) else ""

        record["_sheet_row"] = sheet_row_num
        records.append(record)

    df = pd.DataFrame(records)

    # 컬럼 보정
    for col in VACATION_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    # boolean 처리
    if "잠금" in df.columns:
        df["잠금"] = df["잠금"].apply(
            lambda x: str(x).strip().lower() in ["true", "1", "yes", "y"]
        )

    # 문자열 컬럼 정리
    for col in ["날짜", "메모"]:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str)

    # id 보정 (없는 경우 대비)
    if "id" not in df.columns:
        df["id"] = [str(uuid.uuid4()) for _ in range(len(df))]
    else:
        df["id"] = df["id"].apply(
            lambda x: str(x) if str(x).strip() else str(uuid.uuid4())
        )

    return df[VACATION_COLUMNS + ["_sheet_row"]]

def save_vacation_data(df: pd.DataFrame):
    save_df = normalize_vacation_df(df.copy())
    save_df["잠금"] = save_df["잠금"].apply(lambda x: True if bool(x) else False)

    if sheet_enabled():
        ws = get_worksheet(PROCEDURES_SHEET)
        ws.clear()
        ws.update("A1", [VACATION_COLUMNS] + save_df[VACATION_COLUMNS].astype(str).values.tolist())
        return

    save_df.to_csv(VACATION_FILE, index=False, encoding="utf-8-sig")

def refresh_procedures():
    st.session_state["procedures"] = load_data()

def refresh_vacation_notes():
    st.session_state["vacation_notes"] = load_vacation_data()

def get_df_index_by_id(df: pd.DataFrame, row_id: str):
    matches = df.index[df["id"].astype(str) == str(row_id)].tolist()
    return matches[0] if matches else None

def find_sheet_row_by_id(ws, row_id: str):
    values = ws.get_all_values()
    if not values:
        return None, None

    header = values[0]
    if "id" not in header:
        return None, header

    id_idx = header.index("id")
    for row_num, row in enumerate(values[1:], start=2):
        if len(row) > id_idx and str(row[id_idx]) == str(row_id):
            return row_num, header
    return None, header

def update_procedure_record(row_id: str, updates: dict):
    if not sheet_enabled():
        return

    ws = get_worksheet(PROCEDURES_SHEET)
    row_num, header = find_sheet_row_by_id(ws, row_id)
    if row_num is None:
        return

    current_row = ws.row_values(row_num)
    record = {header[i]: current_row[i] if i < len(current_row) else "" for i in range(len(header))}
    record.update(updates)
    record["updated_at"] = now_iso()
    ws.update(f"A{row_num}:{gspread.utils.rowcol_to_a1(row_num, len(header))}", [record_to_row(record, header)])

def append_procedure_record(record: dict):
    record = dict(record)
    record["id"] = record.get("id") or str(uuid.uuid4())
    record["updated_at"] = now_iso()

    if sheet_enabled():
        ws = get_worksheet(PROCEDURES_SHEET)
        header = ws.row_values(1) or COLUMNS
        ws.append_row(record_to_row(record, header))
        return

    df = load_data()
    save_data(pd.concat([df, pd.DataFrame([record])], ignore_index=True))

def delete_procedure_record(row_id: str):
    if not sheet_enabled():
        return

    ws = get_worksheet(PROCEDURES_SHEET)
    row_num, _ = find_sheet_row_by_id(ws, row_id)
    if row_num is not None:
        ws.delete_rows(row_num)

def upsert_vacation_note_record(date_str: str, memo: str, locked: bool):
    ws = get_worksheet(VACATION_SHEET)
    vac_df = st.session_state.get("vacation_notes", pd.DataFrame()).copy()

    memo = "" if memo is None else str(memo).strip()
    now_str = datetime.now().isoformat(timespec="seconds")

    matches = vac_df[vac_df["날짜"].astype(str) == str(date_str)]

    # 기존 행 수정
    if not matches.empty:
        idx = matches.index[0]
        sheet_row = vac_df.at[idx, "_sheet_row"] if "_sheet_row" in vac_df.columns else None

        # 세션에 sheet row가 있으면 읽기 없이 바로 update
        if pd.notna(sheet_row):
            record = {
                "id": str(vac_df.at[idx, "id"]) if "id" in vac_df.columns else str(uuid.uuid4()),
                "날짜": str(date_str),
                "메모": memo,
                "잠금": bool(locked),
                "updated_at": now_str,
            }

            row_data = []
            for col in VACATION_COLUMNS:
                val = record.get(col, "")
                if isinstance(val, bool):
                    row_data.append("TRUE" if val else "FALSE")
                else:
                    row_data.append("" if pd.isna(val) else str(val))

            sheet_row = int(sheet_row)
            end_a1 = gspread.utils.rowcol_to_a1(sheet_row, len(VACATION_COLUMNS))
            end_col = ''.join([c for c in end_a1 if c.isalpha()])
            ws.update(f"A{sheet_row}:{end_col}{sheet_row}", [row_data])
            return

    # 새 행 append
    new_record = {
        "id": str(uuid.uuid4()) if matches.empty else str(matches.iloc[0].get("id", uuid.uuid4())),
        "날짜": str(date_str),
        "메모": memo,
        "잠금": bool(locked),
        "updated_at": now_str,
    }

    row_data = []
    for col in VACATION_COLUMNS:
        val = new_record.get(col, "")
        if isinstance(val, bool):
            row_data.append("TRUE" if val else "FALSE")
        else:
            row_data.append("" if pd.isna(val) else str(val))

    ws.append_row(row_data)

def ensure_vacation_row(date_str):
    return

def get_vacation_note(date_str: str):
    vac_df = st.session_state.get("vacation_notes", pd.DataFrame())

    if vac_df is None or vac_df.empty:
        return "", False

    matches = vac_df[vac_df["날짜"].astype(str) == str(date_str)]
    if matches.empty:
        return "", False

    row = matches.iloc[0]
    memo = "" if pd.isna(row.get("메모", "")) else str(row.get("메모", ""))
    locked = bool(row.get("잠금", False))
    return memo, locked

def save_vacation_note(date_str: str, memo: str, locked: bool):
    memo = "" if memo is None else str(memo).strip()
    now_str = datetime.now().isoformat(timespec="seconds")

    # 1) 시트 저장
    upsert_vacation_note_record(date_str, memo, locked)

    # 2) 세션 갱신
    vac_df = st.session_state.get("vacation_notes", pd.DataFrame()).copy()

    if vac_df is None or vac_df.empty:
        vac_df = pd.DataFrame(columns=VACATION_COLUMNS + ["_sheet_row"])

    matches = vac_df[vac_df["날짜"].astype(str) == str(date_str)]

    if not matches.empty:
        idx = matches.index[0]
        vac_df.at[idx, "메모"] = memo
        vac_df.at[idx, "잠금"] = bool(locked)
        if "updated_at" in vac_df.columns:
            vac_df.at[idx, "updated_at"] = now_str
    else:
        new_row = {
            "id": str(uuid.uuid4()),
            "날짜": str(date_str),
            "메모": memo,
            "잠금": bool(locked),
            "updated_at": now_str,
            "_sheet_row": None,
        }
        vac_df = pd.concat([vac_df, pd.DataFrame([new_row])], ignore_index=True)

    st.session_state["vacation_notes"] = vac_df

@st.dialog("휴가/부재 메모")
def vacation_note_dialog(date_str: str):
    memo_text, _ = get_vacation_note(date_str)

    input_key = f"vac_popup_text_{date_str}"
    if input_key not in st.session_state:
        st.session_state[input_key] = memo_text

    st.markdown(f"**{date_str}**")
    st.text_input(
        "메모",
        key=input_key,
        placeholder="예: 홍길동 휴가",
        label_visibility="visible"
    )

    b1, b2 = st.columns(2)

    with b1:
        if st.button("확인", key=f"vac_popup_save_{date_str}", use_container_width=True):
            save_vacation_note(date_str, st.session_state.get(input_key, ""), True)
            st.rerun()

    with b2:
        if st.button("삭제", key=f"vac_popup_delete_{date_str}", use_container_width=True):
            save_vacation_note(date_str, "", False)
            st.rerun()

if "procedures" not in st.session_state:
    st.session_state["procedures"] = load_data()

if "calendar_year" not in st.session_state:
    today0 = datetime.today()
    st.session_state["calendar_year"] = today0.year

if "vacation_notes" not in st.session_state:
    st.session_state["vacation_notes"] = load_vacation_data()

if "calendar_month" not in st.session_state:
    today0 = datetime.today()
    st.session_state["calendar_month"] = today0.month

def normalize_gender(value):
    if pd.isna(value):
        return ""
    s = str(value).strip().upper()
    if s in ["M", "남", "남자", "MALE"]:
        return "M"
    if s in ["F", "여", "여자", "FEMALE"]:
        return "F"
    return str(value).strip()

def normalize_age(value):
    if pd.isna(value):
        return 0
    s = str(value).strip()
    match = re.search(r"\d+", s)
    return int(match.group()) if match else 0

def reindex_day_orders(target_date):
    df = st.session_state["procedures"].copy()
    mask = df["날짜"] == target_date
    day_df = df[mask].sort_values("순서").copy()

    for new_order, (idx, row) in enumerate(day_df.iterrows(), start=1):
        df.at[idx, "순서"] = new_order
        df.at[idx, "updated_at"] = now_iso()
        if sheet_enabled():
            update_procedure_record(row["id"], {"순서": new_order})

    st.session_state["procedures"] = df
    if not sheet_enabled():         
        save_data(df)

def extract_ward_text(admission_text: str) -> str:
    if admission_text is None:
        return "외래"

    s = str(admission_text).strip()
    if s == "":
        return "외래"

    if "외래" in s:
        return "외래"

    m = re.search(r"\((.*?)\)", s)
    if m:
        inside = m.group(1)
        parts = [p.strip() for p in inside.split(",") if p.strip()]
        if len(parts) >= 1:
            return ",".join(parts)  

    m2 = re.search(r"(\d{2,4}-\d{1,2})", s)
    if m2:
        return m2.group(1)

    if "입원" in s:
        return "입원"

    return "외래"

def extract_emr_n_ward_text(ward_text: str) -> str:
    if ward_text is None:
        return ""

    s = str(ward_text).strip()
    if s == "" or s.lower() in ["nan", "none"]:
        return ""

    # 10BW(1034) -> 그대로 유지
    m = re.match(r"^([A-Za-z0-9\-]+)\(([^()]+)\)$", s)
    if m:
        ward = m.group(1).strip()
        room = m.group(2).strip()
        return f"{ward}({room})"

    # 혹시 괄호 앞뒤 공백이 섞여 있어도 정리
    m2 = re.match(r"^(.+?)\s*\(\s*([^()]+)\s*\)$", s)
    if m2:
        ward = m2.group(1).strip()
        room = m2.group(2).strip()
        return f"{ward}({room})"

    # 그 외는 원문 유지
    return s

def infer_procedure_text(text: str) -> str:
    if text is None or pd.isna(text):
        return ""

    lines = re.split(r'[\r\n]+', str(text))

    for line in lines:
        line = line.strip()

        if "시행하겠습니다" in line:
            before = line.split("시행하겠습니다")[0].strip()
            return before.rstrip(" ,.:;/-")

    return ""

    if text is None or pd.isna(text):
        return ""

    t = str(text).strip()
    if not t:
        return ""

    t = re.sub(r"\s+", " ", t).strip()

    m = re.search(r"(.+?)\s*시행하겠습니다", t)
    if not m:
        return ""

    return m.group(1).strip(" ,.:;/-")

def parse_emr_text_to_dataframe(raw_text: str) -> pd.DataFrame:
    if not raw_text or raw_text.strip() == "":
        return pd.DataFrame()

    sio = io.StringIO(raw_text)
    reader = csv.reader(sio, delimiter="\t", quotechar='"')
    rows = list(reader)

    if not rows:
        return pd.DataFrame()

    cleaned_rows = []
    for row in rows:
        if row and any(str(cell).strip() != "" for cell in row):
            cleaned_rows.append(row)

    if len(cleaned_rows) < 2:
        return pd.DataFrame()

    header = cleaned_rows[0]
    data_rows = cleaned_rows[1:]

    fixed_rows = []
    ncol = len(header)
    for row in data_rows:
        if len(row) < ncol:
            row = row + [""] * (ncol - len(row))
        elif len(row) > ncol:
            row = row[:ncol]
        fixed_rows.append(row)

    return pd.DataFrame(fixed_rows, columns=header)

def map_emr_n_proc_dept(dept_text: str) -> str:
    dept = "" if dept_text is None or pd.isna(dept_text) else str(dept_text).strip()
    mapping = {
        "신경외과": "NS",
        "신경과": "NU",
    }
    return mapping.get(dept, "")

def parse_emr_n_text_to_dataframe(raw_text: str) -> pd.DataFrame:
    if not raw_text or raw_text.strip() == "":
        return pd.DataFrame()

    sio = io.StringIO(raw_text)
    reader = csv.reader(sio, delimiter="\t", quotechar='"')
    rows = list(reader)

    if not rows:
        return pd.DataFrame()

    cleaned_rows = []
    for row in rows:
        if row and any(str(cell).strip() != "" for cell in row):
            cleaned_rows.append(row)

    if len(cleaned_rows) < 2:
        return pd.DataFrame()

    header = cleaned_rows[0]
    data_rows = cleaned_rows[1:]

    fixed_rows = []
    ncol = len(header)
    for row in data_rows:
        if len(row) < ncol:
            row = row + [""] * (ncol - len(row))
        elif len(row) > ncol:
            row = row[:ncol]
        fixed_rows.append(row)

    return pd.DataFrame(fixed_rows, columns=header)

def set_status(row_id, status):
    if status not in VALID_STATUSES:
        status = STATUS_PLANNED

    df = st.session_state["procedures"].copy()
    idx = get_df_index_by_id(df, row_id)
    if idx is None:
        refresh_procedures()
        return

    df.at[idx, "진행상황"] = status
    df.at[idx, "updated_at"] = now_iso()
    st.session_state["procedures"] = df

    if sheet_enabled():
        update_procedure_record(row_id, {"진행상황": status})
    else:
        save_data(df)

def set_emergency(row_id, value):
    df = st.session_state["procedures"].copy()
    idx = get_df_index_by_id(df, row_id)
    if idx is None:
        refresh_procedures()
        return

    emergency_value = (value == "🚑")
    df.at[idx, "응급"] = emergency_value
    df.at[idx, "updated_at"] = now_iso()
    st.session_state["procedures"] = df

    if sheet_enabled():
        update_procedure_record(row_id, {"응급": emergency_value})
    else:
        save_data(df)

def toggle_consent(row_id):
    df = st.session_state["procedures"].copy()
    idx = get_df_index_by_id(df, row_id)
    if idx is None:
        refresh_procedures()
        return

    current = False
    if "동의서" in df.columns:
        current = bool(df.at[idx, "동의서"])

    new_value = not current
    df.at[idx, "동의서"] = new_value
    df.at[idx, "updated_at"] = now_iso()
    st.session_state["procedures"] = df

    if sheet_enabled():
        update_procedure_record(row_id, {"동의서": new_value})
    else:
        save_data(df)

def save_infection_info(row_id, infection_text):
    df = st.session_state["procedures"].copy()
    idx = get_df_index_by_id(df, row_id)
    if idx is None:
        refresh_procedures()
        return

    text = "" if infection_text is None else str(infection_text).strip()
    infection_on = bool(text)

    df.at[idx, "감염"] = infection_on
    df.at[idx, "감염메모"] = text if infection_on else ""
    df.at[idx, "updated_at"] = now_iso()
    st.session_state["procedures"] = df

    if sheet_enabled():
        update_procedure_record(row_id, {"감염": infection_on, "감염메모": text if infection_on else ""})
    else:
        save_data(df)

def save_adr_info(row_id, adr_text):
    df = st.session_state["procedures"].copy()
    idx = get_df_index_by_id(df, row_id)
    if idx is None:
        refresh_procedures()
        return

    text = "" if adr_text is None else str(adr_text).strip()
    adr_on = bool(text)

    df.at[idx, "ADR"] = adr_on
    df.at[idx, "ADR메모"] = text if adr_on else ""
    df.at[idx, "updated_at"] = now_iso()
    st.session_state["procedures"] = df

    if sheet_enabled():
        update_procedure_record(row_id, {"ADR": adr_on, "ADR메모": text if adr_on else ""})
    else:
        save_data(df)

def save_renal_info(row_id, cr_text):
    df = st.session_state["procedures"].copy()
    idx = get_df_index_by_id(df, row_id)
    if idx is None:
        refresh_procedures()
        return

    text = "" if cr_text is None else str(cr_text).strip()
    renal_on = bool(text)

    df.at[idx, "신기능"] = renal_on
    df.at[idx, "Cr"] = text if renal_on else ""
    df.at[idx, "updated_at"] = now_iso()
    st.session_state["procedures"] = df

    if sheet_enabled():
        update_procedure_record(row_id, {"신기능": renal_on, "Cr": text if renal_on else ""})
    else:
        save_data(df)

def save_bleeding_info(row_id, pt_inr_text, plt_text):
    df = st.session_state["procedures"].copy()
    idx = get_df_index_by_id(df, row_id)
    if idx is None:
        refresh_procedures()
        return

    pt_text = "" if pt_inr_text is None else str(pt_inr_text).strip()
    plt_val = "" if plt_text is None else str(plt_text).strip()
    bleeding_on = bool(pt_text or plt_val)

    df.at[idx, "출혈"] = bleeding_on
    df.at[idx, "PT_INR"] = pt_text if bleeding_on else ""
    df.at[idx, "PLT"] = plt_val if bleeding_on else ""
    df.at[idx, "updated_at"] = now_iso()
    st.session_state["procedures"] = df

    if sheet_enabled():
        update_procedure_record(row_id, {
            "출혈": bleeding_on,
            "PT_INR": pt_text if bleeding_on else "",
            "PLT": plt_val if bleeding_on else ""
        })
    else:
        save_data(df)

def status_badge(label: str, status_type: str):
    bg = STATUS_COLORS.get(status_type, "#9aa0a6")
    st.markdown(
        f"""
        <div style="
            width: 100%;
            text-align: center;
            padding: 0.38rem 0.2rem;
            border-radius: 0.5rem;
            font-size: 0.9rem;
            font-weight: 600;
            color: white;
            background: {bg};
            border: 1px solid {bg};
            box-sizing: border-box;
        ">{label}</div>
        """,
        unsafe_allow_html=True,
    )

def get_display_day_df(df: pd.DataFrame, selected_date: str) -> pd.DataFrame:
    day_df = df[df["날짜"] == selected_date].copy()
    status_priority = {
        STATUS_INROOM: 0,
        STATUS_CALLED: 1,
        STATUS_PLANNED: 2,
        STATUS_DONE: 3
    }
    day_df["status_priority"] = day_df["진행상황"].map(status_priority).fillna(2)
    return day_df.sort_values(["status_priority", "순서"], ascending=[True, True])

def prev_month():
    y = st.session_state["calendar_year"]
    m = st.session_state["calendar_month"]
    if m == 1:
        st.session_state["calendar_year"] = y - 1
        st.session_state["calendar_month"] = 12
    else:
        st.session_state["calendar_month"] = m - 1

def next_month():
    y = st.session_state["calendar_year"]
    m = st.session_state["calendar_month"]
    if m == 12:
        st.session_state["calendar_year"] = y + 1
        st.session_state["calendar_month"] = 1
    else:
        st.session_state["calendar_month"] = m + 1

def update_memo(row_id, memo_value):
    df = st.session_state["procedures"].copy()
    idx = get_df_index_by_id(df, row_id)
    if idx is None:
        refresh_procedures()
        return

    old_memo = "" if pd.isna(df.at[idx, "메모"]) else str(df.at[idx, "메모"])
    new_memo = "" if memo_value is None else str(memo_value)

    if old_memo == new_memo:
        return

    df.at[idx, "메모"] = new_memo
    df.at[idx, "updated_at"] = now_iso()
    st.session_state["procedures"] = df

    if sheet_enabled():
        update_procedure_record(row_id, {"메모": new_memo})
    else:
        save_data(df)

def update_procedure_edit_fields(
    row_id,
    reg_id,
    name,
    ward,
    proc_dept,
    procedure_value,
    room,
    prof,
    req_dept,
    req_doctor,
    emergency_value,
    memo_value,
):
    df = st.session_state["procedures"].copy()
    idx = get_df_index_by_id(df, row_id)
    if idx is None:
        refresh_procedures()
        return

    normalized_updates = {
        "등록번호": "" if reg_id is None else str(reg_id).strip(),
        "이름": "" if name is None else str(name).strip(),
        "병실": "" if ward is None else str(ward).strip(),
        "시술과": "IR" if proc_dept not in ["IR", "NS", "NU"] else proc_dept,
        "시술명": "" if procedure_value is None else str(procedure_value).strip(),
        "Room": normalize_room_value(room),
        "교수": "" if prof is None else str(prof).strip(),
        "의뢰과": "" if req_dept is None else str(req_dept).strip(),
        "의뢰의": "" if req_doctor is None else str(req_doctor).strip(),
        "응급": str(emergency_value).strip() == "🚑",
        "메모": "" if memo_value is None else str(memo_value).strip(),
    }

    changed = False
    for col, new_value in normalized_updates.items():
        current_value = df.at[idx, col]
        if isinstance(new_value, bool):
            if bool(current_value) != new_value:
                changed = True
                break
        else:
            current_text = "" if pd.isna(current_value) else str(current_value).strip()
            if current_text != str(new_value).strip():
                changed = True
                break

    if not changed:
        return
        
    for col, new_value in normalized_updates.items():
        if col in df.columns:
            df[col] = df[col].astype(object)

    for col, new_value in normalized_updates.items():
        df.at[idx, col] = new_value

    df.at[idx, "updated_at"] = now_iso()
    st.session_state["procedures"] = df

    if sheet_enabled():
        update_procedure_record(row_id, normalized_updates)
    else:
        save_data(df)

def get_day_case_summary(df: pd.DataFrame, date_str: str):
    day_df = df[df["날짜"] == date_str].copy()
    total_count = len(day_df)
    ir_count = len(day_df[day_df["시술과"] == "IR"])
    ns_count = len(day_df[day_df["시술과"] == "NS"])
    nu_count = len(day_df[day_df["시술과"] == "NU"])
    neuro_total = ns_count + nu_count
    return total_count, ir_count, ns_count, nu_count, neuro_total

def format_date_with_weekday(date_str: str) -> str:
    try:
        dt = datetime.strptime(date_str, "%Y-%m-%d")
        weekday_kr = ["월", "화", "수", "목", "금", "토", "일"][dt.weekday()]
        return f"{date_str}({weekday_kr})"
    except Exception:
        return str(date_str)

def format_term_from_today(date_str: str) -> str:
    try:
        target_date = datetime.strptime(date_str, "%Y-%m-%d").date()
        today = datetime.today().date()

        delta_days = (today - target_date).days
        if delta_days < 0:
            return "0d"

        years = delta_days // 365
        remain = delta_days % 365
        months = remain // 30
        days = remain % 30

        parts = []
        if years > 0:
            parts.append(f"{years}y")
        if months > 0:
            parts.append(f"{months}m")
        if days > 0 or not parts:
            parts.append(f"{days}d")

        return " ".join(parts)
    except Exception:
        return ""

def get_month_case_total(df: pd.DataFrame, year: int, month: int) -> int:
    prefix = f"{year}-{month:02d}-"
    month_df = df[df["날짜"].astype(str).str.startswith(prefix)].copy()
    return len(month_df)

def render_rank_table(df: pd.DataFrame):
    st.dataframe(df, use_container_width=True, hide_index=True)

@st.dialog("시술 삭제")
def delete_dialog(row_id):
    pw = st.text_input("비밀번호", type="password", key=f"delete_pw_{row_id}")

    if st.button("삭제 확인", key=f"delete_confirm_{row_id}"):
        if pw == DELETE_PASSWORD:
            df = st.session_state["procedures"].copy()
            idx = get_df_index_by_id(df, row_id)
            if idx is None:
                refresh_procedures()
                st.error("대상 행을 찾지 못했습니다.")
                return

            target_date = df.loc[idx, "날짜"]
            df = df.drop(idx).reset_index(drop=True)
            st.session_state["procedures"] = df

            if sheet_enabled():
                delete_procedure_record(row_id)
            else:
                save_data(df)

            reindex_day_orders(target_date)
            refresh_procedures()
            st.success("삭제되었습니다")
            st.rerun()
        else:
            st.error("비밀번호가 틀렸습니다")

@st.dialog("시술 정보 수정")
def edit_procedure_dialog(row_id):
    df = st.session_state["procedures"].copy()
    idx = get_df_index_by_id(df, row_id)
    if idx is None:
        st.error("대상 행을 찾지 못했습니다.")
        return

    row = df.loc[idx]
    proc_dept_options = ["IR", "NS", "NU"]
    room_options = ["", "1", "2", "H"]
    current_proc_dept = row["시술과"] if row["시술과"] in proc_dept_options else "IR"
    current_room = normalize_room_value(row.get("Room", ""))
    current_emergency = "🚑" if bool(row.get("응급", False)) else "N"

    reg_id = st.text_input("등록번호", value="" if pd.isna(row["등록번호"]) else str(row["등록번호"]))
    name = st.text_input("이름", value="" if pd.isna(row["이름"]) else str(row["이름"]))
    ward = st.text_input("병실", value="" if pd.isna(row["병실"]) else str(row["병실"]))

    c1, c2 = st.columns(2)
    proc_dept = c1.selectbox("시술과", proc_dept_options, index=proc_dept_options.index(current_proc_dept))
    room = c2.selectbox("Room", room_options, index=room_options.index(current_room))

    procedure_value = st.text_input("시술명", value="" if pd.isna(row["시술명"]) else str(row["시술명"]))
    prof_options = get_prof_select_options(proc_dept)
    current_prof = row["교수"] if row["교수"] in prof_options else ""
    prof = st.selectbox("시술의", prof_options, index=prof_options.index(current_prof))

    c3, c4 = st.columns(2)
    req_dept = c3.text_input("의뢰과", value="" if pd.isna(row["의뢰과"]) else str(row["의뢰과"]))
    req_doctor = c4.text_input("의뢰의", value="" if pd.isna(row["의뢰의"]) else str(row["의뢰의"]))

    memo_value = st.text_area("메모", value="" if pd.isna(row.get("메모", "")) else str(row.get("메모", "")), height=100)
    emergency_value = st.selectbox("응급", ["N", "🚑"], index=0 if current_emergency == "N" else 1)

    if st.button("확인", key=f"edit_confirm_{row_id}", use_container_width=True):
        update_procedure_edit_fields(
            row_id,
            reg_id,
            name,
            ward,
            proc_dept,
            procedure_value,
            room,
            prof,
            req_dept,
            req_doctor,
            emergency_value,
            memo_value,
        )
        st.rerun()

@st.dialog("시술의 Ranking")
def ranking_dialog(selected_date):
    df = st.session_state["procedures"].copy()
    day_df = df[df["날짜"] == selected_date].copy()
    day_df = day_df[day_df["교수"].astype(str).str.strip() != ""].copy()

    if day_df.empty:
        st.info("표시할 시술의 데이터가 없습니다.")
        return

    rank_df = (
        day_df.groupby("교수")
        .size()
        .reset_index(name="건수")
        .sort_values(["건수", "교수"], ascending=[False, True])
        .reset_index(drop=True)
    )

    rank_df["시술의"] = rank_df["교수"]
    if len(rank_df) > 0:
        rank_df.loc[0, "시술의"] = f"{rank_df.loc[0, '교수']} 🥇"

    show_df = rank_df[["시술의", "건수"]].copy()
    render_rank_table(show_df)

@st.dialog("월간 Ranking")
def monthly_ranking_dialog(year: int, month: int):
    df = st.session_state["procedures"].copy()
    prefix = f"{year}-{month:02d}-"
    month_df = df[df["날짜"].astype(str).str.startswith(prefix)].copy()

    if month_df.empty:
        st.info("해당 월의 시술 데이터가 없습니다.")
        return

    st.markdown(f"### {year}년 {month}월")

    dept_rank_df = (
        month_df.groupby("시술과")
        .size()
        .reset_index(name="건수")
        .sort_values(["건수", "시술과"], ascending=[False, True])
        .reset_index(drop=True)
    )
    dept_rank_df["순위"] = dept_rank_df.index + 1
    if len(dept_rank_df) > 0:
        dept_rank_df.loc[0, "시술과"] = f"{dept_rank_df.loc[0, '시술과']} 🥇"

    st.markdown("#### 시술과별")
    render_rank_table(dept_rank_df[["순위", "시술과", "건수"]])

    prof_df = month_df[month_df["교수"].astype(str).str.strip() != ""].copy()
    if prof_df.empty:
        st.markdown("#### 시술의별")
        st.info("해당 월의 시술의 데이터가 없습니다.")
    else:
        prof_rank_df = (
            prof_df.groupby("교수")
            .size()
            .reset_index(name="건수")
            .sort_values(["건수", "교수"], ascending=[False, True])
            .reset_index(drop=True)
        )
        prof_rank_df["순위"] = prof_rank_df.index + 1
        if len(prof_rank_df) > 0:
            prof_rank_df.loc[0, "교수"] = f"{prof_rank_df.loc[0, '교수']} 🥇"

        st.markdown("#### 시술의별")
        render_rank_table(prof_rank_df[["순위", "교수", "건수"]])

@st.dialog("직접 입력")
def add_procedure(selected_date):
    reg_id = st.text_input("등록번호")
    name = st.text_input("이름")

    col1, col2 = st.columns(2)
    gender = col1.selectbox("성별", ["M", "F"])
    age = col2.number_input("나이", 0, 120, 60)

    ward = st.text_input("병실")
    department_proc = st.selectbox("시술과", ["IR", "NS", "NU"], index=0)
    procedure = st.text_input("시술명")

    col3, col4 = st.columns(2)
    dept = col3.text_input("의뢰과")
    doctor = col4.text_input("의뢰의")

    emergency_value = st.selectbox("응급", ["N", "🚑"], index=0)
    emergency = (emergency_value == "🚑")

    room_options = ["", "1", "2", "H"]
    room = st.selectbox("Room", room_options, index=0)

    prof_options = get_prof_select_options(department_proc)
    prof = st.selectbox("시술의", prof_options, index=0)
    memo = st.text_area("메모", height=100)

    if st.button("등록"):
        df = st.session_state["procedures"].copy()
        order = len(df[df["날짜"] == selected_date]) + 1
        
        record = {
            "id": str(uuid.uuid4()),
            "updated_at": now_iso(),
            "날짜": selected_date,
            "순서": order,
            "등록번호": reg_id,
            "이름": name,
            "성별": gender,
            "나이": int(age),
            "병실": ward,
            "시술과": department_proc,
            "시술명": procedure,
            "의뢰과": dept,
            "의뢰의": doctor,
            "Room": room,
            "교수": prof,
            "응급": emergency,
            "진행상황": STATUS_PLANNED,
            "동의서": False,
            "감염": False,
            "감염메모": "",
            "ADR": False,
            "ADR메모": "",
            "신기능": False,
            "Cr": "",
            "출혈": False,
            "PT_INR": "",
            "PLT": "",
            "메모": "" if memo is None else str(memo).strip(),
        }

        st.session_state["procedures"] = pd.concat([df, pd.DataFrame([record])], ignore_index=True)
        if sheet_enabled():
            append_procedure_record(record)
        else:
            save_data(st.session_state["procedures"])
        refresh_procedures()
        st.rerun()

@st.dialog("EMR 텍스트 붙여넣기")
def paste_emr_dialog(selected_date):
    st.write(f"선택 날짜: {selected_date}")
    st.caption("EMR 협진창의 의뢰 정보를 복사해서 아래에 그대로 붙여넣으세요.")

    raw_text = st.text_area(
        "EMR 복사 텍스트",
        height=300,
        placeholder="EMR에서 복사한 내용을 여기에 붙여넣기"
    )

    # -------------------------
    # 시술명 선택 UI 추가
    # -------------------------
    procedure_options = [
    "직접입력",
    "PICC",
    "PICC change",
    "Chemoport insertion",
    "Perm HD insertion",
    "PTBD",
    "PTGBD",
    "Ascites PCD",
    "Chest PCD, Rt.",
    "Chest PCD, Lt.",
    "Chest PCD, Both",
    "Abscess PCD, liver",
    "Abscess PCD, abdomen",
    "PCD removal",
    "PTGBD removal",
    "PTGBD tubogram",
    "PRG",
    "PCN, Both",
    "PCN, Rt.",
    "PCN, Lt.",
    "TACE",
    "TAE(Transarterial embolization)",
    "BAE",
    "UAE",
    "PTA, L/E",
    "PTA, AVF(G)",
    "IVC filter insertion",
    "IVC filter removal",
]

    procedure_choice = st.selectbox(
        "시술명 선택",
        procedure_options,
        index=0,
        key="emr_procedure_choice"
    )

    custom_procedure_name = ""
    if procedure_choice == "직접입력":
        custom_procedure_name = st.text_input(
            "직접 입력 시술명",
            value="",
            placeholder="시술명을 입력하세요",
            key="emr_custom_procedure_name"
        )

    doctor_options = ["정선화", "박상영"]
    selected_doctor = st.selectbox(
        "시술의 선택",
        doctor_options,
        index=0,
        key="emr_doctor_choice"
    )

    if st.button("불러오기"):
        if raw_text.strip() == "":
            st.error("붙여넣은 텍스트가 없습니다.")
            return

        # 최종 시술명 결정
        if procedure_choice == "직접입력":
            selected_procedure_name = custom_procedure_name.strip()
            if selected_procedure_name == "":
                st.error("직접입력을 선택한 경우 시술명을 입력해야 합니다.")
                return
        else:
            selected_procedure_name = procedure_choice.strip()

        try:
            emr_df = parse_emr_text_to_dataframe(raw_text)
        except Exception as e:
            st.error(f"텍스트를 읽는 중 오류가 발생했습니다: {e}")
            return

        if emr_df.empty:
            st.error("텍스트에서 표 형식을 읽지 못했습니다.")
            return

        required_cols = ["등록번호", "환자명", "성별", "나이", "의뢰과", "의뢰의", "입/외", "회신내용"]
        missing_cols = [col for col in required_cols if col not in emr_df.columns]

        if missing_cols:
            st.error(f"다음 컬럼이 없습니다: {', '.join(missing_cols)}")
            return

        current_df = st.session_state["procedures"].copy()
        current_day_df = current_df[current_df["날짜"] == selected_date]
        next_order = len(current_day_df) + 1

        new_rows = []

        for _, row in emr_df.iterrows():
            reg_id = "" if pd.isna(row["등록번호"]) else str(row["등록번호"]).strip()
            name = "" if pd.isna(row["환자명"]) else str(row["환자명"]).strip()

            if reg_id in ["", "nan", "None"] and name in ["", "nan", "None"]:
                continue

            admission_text = "" if pd.isna(row["입/외"]) else str(row["입/외"]).strip()
            ward = extract_ward_text(admission_text)

            new_rows.append({
                "id": str(uuid.uuid4()),
                "updated_at": now_iso(),
                "날짜": selected_date,
                "순서": next_order,
                "등록번호": reg_id,
                "이름": name,
                "성별": normalize_gender(row["성별"]),
                "나이": normalize_age(row["나이"]),
                "병실": ward,
                "시술과": "IR",
                "시술명": selected_procedure_name,
                "의뢰과": "" if pd.isna(row["의뢰과"]) else str(row["의뢰과"]).strip(),
                "의뢰의": "" if pd.isna(row["의뢰의"]) else str(row["의뢰의"]).strip(),
                "Room": "2",
                "교수": selected_doctor,
                "응급": False if str(row.get("응급", "N")).strip().upper() != "Y" else True,
                "진행상황": STATUS_PLANNED,
                "동의서": False,
                "감염": False,
                "감염메모": "",
                "ADR": False,
                "ADR메모": "",
                "신기능": False,
                "Cr": "",
                "출혈": False,
                "PT_INR": "",
                "PLT": ""
            })
            next_order += 1

        if len(new_rows) == 0:
            st.warning("가져올 환자 목록이 없습니다.")
            return

        new_df = pd.DataFrame(new_rows)
        st.session_state["procedures"] = pd.concat([current_df, new_df], ignore_index=True)

        if sheet_enabled():
            for record in new_rows:
                append_procedure_record(record)
        else:
            save_data(st.session_state["procedures"])

        refresh_procedures()
        st.success(f"{len(new_rows)}명의 환자 목록을 불러왔습니다.")
        st.rerun()

@st.dialog("EMR-N 텍스트 붙여넣기")
def paste_emr_n_dialog(selected_date):
    st.write(f"선택 날짜: {selected_date}")
    st.caption("EMR-N 목록을 그대로 복사해서 아래에 붙여넣으세요.")

    raw_text = st.text_area(
        "EMR-N 복사 텍스트",
        height=300,
        placeholder="EMR-N에서 복사한 내용을 여기에 붙여넣기",
        key="emr_n_text_area"
    )

    emr_n_procedure_options = [
        "직접입력",
        "TFCA",
        "Coil embolization",
        "Thrombectomy",
    ]

    procedure_choice = st.selectbox(
        "시술명 선택",
        emr_n_procedure_options,
        index=0,
        key="emr_n_procedure_choice"
    )

    custom_procedure_name = ""
    if procedure_choice == "직접입력":
        custom_procedure_name = st.text_input(
            "직접 입력 시술명",
            value="",
            placeholder="시술명을 입력하세요",
            key="emr_n_custom_procedure_name"
        )

    if st.button("불러오기", key="load_emr_n"):
        if raw_text.strip() == "":
            st.error("붙여넣은 텍스트가 없습니다.")
            return

        if procedure_choice == "직접입력":
            selected_procedure_name = custom_procedure_name.strip()
            if selected_procedure_name == "":
                st.error("직접입력을 선택한 경우 시술명을 입력해야 합니다.")
                return
        else:
            selected_procedure_name = procedure_choice.strip()

        try:
            emr_df = parse_emr_n_text_to_dataframe(raw_text)
        except Exception as e:
            st.error(f"텍스트를 읽는 중 오류가 발생했습니다: {e}")
            return

        if emr_df.empty:
            st.error("텍스트에서 표 형식을 읽지 못했습니다.")
            return

        required_cols = ["등록번호", "환자명", "성별", "나이", "진료과", "시술의", "병동(병실)"]
        missing_cols = [col for col in required_cols if col not in emr_df.columns]

        if missing_cols:
            st.error(f"다음 컬럼이 없습니다: {', '.join(missing_cols)}")
            return

        current_df = st.session_state["procedures"].copy()
        current_day_df = current_df[current_df["날짜"] == selected_date]
        next_order = len(current_day_df) + 1

        new_rows = []

        for _, row in emr_df.iterrows():
            reg_id = "" if pd.isna(row["등록번호"]) else str(row["등록번호"]).strip()
            name = "" if pd.isna(row["환자명"]) else str(row["환자명"]).strip()

            if reg_id in ["", "nan", "None"] and name in ["", "nan", "None"]:
                continue

            ward_text = "" if pd.isna(row["병동(병실)"]) else str(row["병동(병실)"]).strip()
            proc_dept_src = "" if pd.isna(row["진료과"]) else str(row["진료과"]).strip()
            proc_dept = map_emr_n_proc_dept(proc_dept_src)

            new_rows.append({
                "id": str(uuid.uuid4()),
                "updated_at": now_iso(),
                "날짜": selected_date,
                "순서": next_order,
                "등록번호": reg_id,
                "이름": name,
                "성별": normalize_gender(row["성별"]),
                "나이": normalize_age(row["나이"]),
                "병실": extract_emr_n_ward_text(ward_text),
                "시술과": proc_dept,
                "시술명": selected_procedure_name,
                "의뢰과": proc_dept_src,
                "의뢰의": "" if pd.isna(row.get("진료의", "")) else str(row.get("진료의", "")).strip(),
                "Room": "1",
                "교수": "" if pd.isna(row["시술의"]) else str(row["시술의"]).strip(),
                "응급": False,
                "진행상황": STATUS_PLANNED,
                "동의서": False,
                "감염": False,
                "감염메모": "",
                "ADR": False,
                "ADR메모": "",
                "신기능": False,
                "Cr": "",
                "출혈": False,
                "PT_INR": "",
                "PLT": "",
                "메모": ""
            })
            next_order += 1

        if len(new_rows) == 0:
            st.warning("가져올 환자 목록이 없습니다.")
            return

        new_df = pd.DataFrame(new_rows)
        st.session_state["procedures"] = pd.concat([current_df, new_df], ignore_index=True)

        if sheet_enabled():
            for record in new_rows:
                append_procedure_record(record)
        else:
            save_data(st.session_state["procedures"])

        refresh_procedures()
        st.success(f"{len(new_rows)}명의 환자 목록을 불러왔습니다.")
        st.rerun()

def move_up(row_id):
    df = st.session_state["procedures"].copy()
    idx = get_df_index_by_id(df, row_id)
    if idx is None:
        refresh_procedures()
        return

    date_value = df.loc[idx, "날짜"]
    same_status = df.loc[idx, "진행상황"]

    day_df = df[(df["날짜"] == date_value) & (df["진행상황"] == same_status)].sort_values("순서")
    position_list = day_df["id"].astype(str).tolist()
    if str(row_id) not in position_list:
        return

    pos = position_list.index(str(row_id))
    if pos == 0:
        return

    prev_row_id = position_list[pos - 1]
    prev_index = get_df_index_by_id(df, prev_row_id)

    a = df.loc[idx, "순서"]
    b = df.loc[prev_index, "순서"]

    df.loc[idx, "순서"] = b
    df.loc[prev_index, "순서"] = a
    df.loc[idx, "updated_at"] = now_iso()
    df.loc[prev_index, "updated_at"] = now_iso()
    st.session_state["procedures"] = df

    if sheet_enabled():
        update_procedure_record(row_id, {"순서": b})
        update_procedure_record(prev_row_id, {"순서": a})
    else:
        save_data(df)

def move_down(row_id):
    df = st.session_state["procedures"].copy()
    idx = get_df_index_by_id(df, row_id)
    if idx is None:
        refresh_procedures()
        return

    date_value = df.loc[idx, "날짜"]
    same_status = df.loc[idx, "진행상황"]

    day_df = df[(df["날짜"] == date_value) & (df["진행상황"] == same_status)].sort_values("순서")
    position_list = day_df["id"].astype(str).tolist()
    if str(row_id) not in position_list:
        return

    pos = position_list.index(str(row_id))
    if pos == len(position_list) - 1:
        return

    next_row_id = position_list[pos + 1]
    next_index = get_df_index_by_id(df, next_row_id)

    a = df.loc[idx, "순서"]
    b = df.loc[next_index, "순서"]

    df.loc[idx, "순서"] = b
    df.loc[next_index, "순서"] = a
    df.loc[idx, "updated_at"] = now_iso()
    df.loc[next_index, "updated_at"] = now_iso()
    st.session_state["procedures"] = df

    if sheet_enabled():
        update_procedure_record(row_id, {"순서": b})
        update_procedure_record(next_row_id, {"순서": a})
    else:
        save_data(df)

st.markdown("""

<style>
input:disabled, textarea:disabled {
    -webkit-text-fill-color: #000000 !important;
    color: #000000 !important;
    opacity: 1 !important;
}
[class*="st-key-row_inroom_"] {
    background-color: #fff1f5;
    border: 1px solid #fdf2f8;
    border-radius: 10px;
    padding: 0.9rem 0.35rem 0.9rem 0.45rem;
    margin-bottom: 0.10rem;
}
[class*="st-key-row_done_"] {
    background-color: #eef9ff;
    border: 1px solid #d6eefc;
    border-radius: 10px;
    padding: 0.9rem 0.45rem 0.9rem 0.45rem;
    margin-bottom: 0.10rem;
}
[class*="st-key-cal_"] button {
    height: 160px !important;
    padding: 5px 5px !important;
    white-space: pre-line !important;
}

[class*="st-key-cal_"] button p {
    font-size: 18px !important;
    line-height: 1.5 !important;
    font-weight: 400 !important;
    margin: 0 !important;
}

[class*="st-key-cal_"] button:hover {
    background-color: #f5f5f5;
    border: 3px solid #e6e6e6;
}

[class*="st-key-cal_today"] button {
    background-color: #fff7cc !important;
    border: 3px solid #f2dc7d !important;
}
[class*="st-key-cal_today"] button:hover {
    background-color: #fff2a8 !important;
}

.nowrap-header {
    white-space: nowrap;
    word-break: keep-all;
    overflow-wrap: normal;
    font-weight: 700;
    font-size: 0.9rem;
}

[class*="st-key-cal_sat_"] button {
    color: #2563eb !important;
}

[class*="st-key-cal_sun_"] button {
    color: #dc2626 !important;
}

[class*="st-key-cal_today_sat_"] button {
    color: #2563eb !important;
}

[class*="st-key-cal_today_sun_"] button {
    color: #dc2626 !important;
}

button[kind="primary"][data-testid="baseButton-secondary"],
button[kind="secondary"][data-testid="baseButton-secondary"] {
    border-radius: 0.45rem !important;
}

/* 돋보기 history 버튼 스타일 제거 */
[class*="st-key-history_btn_"] button {
    background: transparent !important;
    border: none !important;
    box-shadow: none !important;
    padding: 0 !important;
    font-size: 1rem !important;
    min-height: auto !important;
    height: auto !important;
}

[class*="st-key-history_btn_"] button:hover {
    background: transparent !important;
}

[class*="st-key-precheck_"] button,
[class*="st-key-infection_"] button {
    min-height: 1.45rem !important;
    height: 1.45rem !important;
    padding: 0.10rem 0.10rem !important;
    font-size: 0.72rem !important;
    line-height: 1.1 !important;
    border-radius: 0.8rem !important;
}

[class*="st-key-precheck_infect_btn_"] button[kind="primary"],
[class*="st-key-infection_save_"] button[kind="primary"] {
    background-color: #16a34a !important;
    border: 1px solid #16a34a !important;
    color: white !important;
}
[class*="st-key-precheck_consent_"] button[kind="primary"] {
    background-color: #38bdf8 !important;
    border: 1px solid #38bdf8 !important;
    color: white !important;
}

[class*="st-key-precheck_infect_btn_"] button[kind="primary"] {
    background-color: #8b5e3c !important;
    border: 1px solid #8b5e3c !important;
    color: white !important;
}

[class*="st-key-precheck_adr_btn_"] button[kind="primary"] {
    background-color: #f59e0b !important;
    border: 1px solid #f59e0b !important;
    color: white !important;
}

[class*="st-key-precheck_renal_btn_"] button[kind="primary"] {
    background-color: #8b5cf6 !important;
    border: 1px solid #8b5cf6 !important;
    color: white !important;
}

[class*="st-key-precheck_bleeding_btn_"] button[kind="primary"] {
    background-color: #fca5a5 !important;
    border: 1px solid #fca5a5 !important;
    color: #7f1d1d !important;
}
           
[class*="st-key-infection_save_"] button,
[class*="st-key-adr_save_"] button,
[class*="st-key-renal_save_"] button,
[class*="st-key-bleeding_save_"] button {
    min-height: 1.00rem !important;
    height: 1.00rem !important;
    padding: 0.02rem 0.35rem !important;
    font-size: 0.68rem !important;
    line-height: 1 !important;
    border-radius: 999px !important;
    background-color: #111827 !important;
    border: 1px solid #111827 !important;
    color: white !important;
}
            
[class*="st-key-vac_plus_"] button {
    min-height: 1.8rem !important;
    height: 1.8rem !important;
    padding: 0.05rem 0 !important;
    font-size: 0.95rem !important;
    border-radius: 0.4rem !important;
} 
</style>
""", unsafe_allow_html=True)

query = st.query_params
selected_date = query.get("date")
duty_month = query.get("duty")
board_date = query.get("board")
history_reg = query.get("history")
history_back_date = query.get("history_date")

# -------------------------
# 당직 페이지
# -------------------------
if duty_month:
    try:
        duty_year = int(duty_month.split("-")[0])
        duty_mon = int(duty_month.split("-")[1])
    except Exception:
        duty_year = st.session_state["calendar_year"]
        duty_mon = st.session_state["calendar_month"]

    top1, top2 = st.columns([1.4, 8.6])
    with top1:
        if st.button("⬅ 메인 캘린더로", use_container_width=True):
            if "duty" in st.query_params:
                del st.query_params["duty"]
            st.rerun()

    st.markdown(
        f"""
        <div style="text-align:center; font-size:2rem; font-weight:700; margin-bottom:0.8rem;">
            {duty_year}년 {duty_mon}월 당직표
        </div>
        """,
        unsafe_allow_html=True
    )

    uploaded_file = st.file_uploader(
        "당직표 이미지 업로드",
        type=["png", "jpg", "jpeg", "webp"],
        key=f"duty_uploader_{duty_year}_{duty_mon:02d}"
    )

    if uploaded_file is not None:
        ext = uploaded_file.name.split(".")[-1].lower()
        if ext not in ["png", "jpg", "jpeg", "webp"]:
            ext = "png"

        # 기존 월 이미지 삭제
        old_path = find_saved_duty_image(duty_year, duty_mon)
        if old_path and os.path.exists(old_path):
            os.remove(old_path)

        save_path = get_duty_image_path(duty_year, duty_mon, ext)
        with open(save_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        st.success("당직표 이미지가 저장되었습니다.")
        st.rerun()

    saved_image_path = find_saved_duty_image(duty_year, duty_mon)

    if saved_image_path:
        st.image(saved_image_path, caption=f"{duty_year}년 {duty_mon}월 당직표", use_container_width=True)

        col_del1, col_del2 = st.columns([1.2, 8.8])
        with col_del1:
            if st.button("이미지 삭제", key=f"delete_duty_img_{duty_year}_{duty_mon:02d}", use_container_width=True):
                os.remove(saved_image_path)
                st.rerun()
    else:
        st.info("업로드된 당직표 이미지가 없습니다.")

    st.stop()

# -------------------------
# 현황판 페이지
# -------------------------
from streamlit_autorefresh import st_autorefresh

if board_date:
    refresh_count = st_autorefresh(interval=30000, key=f"board_autorefresh_{board_date}")

    if "procedures" not in st.session_state or refresh_count > 0:
        st.session_state["procedures"] = load_data()

    top_back, top_refresh, top_space = st.columns([1.4, 1.4, 7.2])

    with top_back:
        if st.button("⬅ 시술목록", use_container_width=True, key=f"back_from_board_{board_date}"):
            replace_query_params(date=board_date)
            st.rerun()

    with top_refresh:
        if st.button("🔄 새로고침", use_container_width=True, key=f"refresh_board_{board_date}"):
            st.session_state["procedures"] = load_data()
            st.rerun()

    df = st.session_state["procedures"].copy()
    day_df = df[df["날짜"] == board_date].copy()

    total_count = len(day_df)
    done_count = len(day_df[day_df["진행상황"] == STATUS_DONE])
    progress_pct = int((done_count / total_count) * 100) if total_count > 0 else 0

    board_dt = datetime.strptime(board_date, "%Y-%m-%d")
    weekday_kr = ["월", "화", "수", "목", "금", "토", "일"][board_dt.weekday()]
    board_date_text = f"{board_date} ({weekday_kr})"

    room2_in_df = day_df[
        (day_df["Room"].astype(str) == "2") &
        (day_df["진행상황"] == STATUS_INROOM)
    ].copy().sort_values("순서")

    room1_in_df = day_df[
        (day_df["Room"].astype(str) == "1") &
        (day_df["진행상황"] == STATUS_INROOM)
    ].copy().sort_values("순서")

    room2_call_df = day_df[
        (day_df["Room"].astype(str) == "2") &
        (day_df["진행상황"].isin([STATUS_CALLED, STATUS_ARRIVED]))
    ].copy()

    room2_call_df["arrived_priority"] = room2_call_df["진행상황"].apply(
        lambda x: 0 if x == STATUS_ARRIVED else 1
    )
    room2_call_df = room2_call_df.sort_values(["arrived_priority", "순서"]).drop(columns=["arrived_priority"])

    room1_call_df = day_df[
        (day_df["Room"].astype(str) == "1") &
        (day_df["진행상황"].isin([STATUS_CALLED, STATUS_ARRIVED]))
    ].copy()

    room1_call_df["arrived_priority"] = room1_call_df["진행상황"].apply(
        lambda x: 0 if x == STATUS_ARRIVED else 1
    )
    room1_call_df = room1_call_df.sort_values(["arrived_priority", "순서"]).drop(columns=["arrived_priority"])

    def patient_text(r):
        reg = "" if pd.isna(r["등록번호"]) else str(r["등록번호"])
        name = "" if pd.isna(r["이름"]) else str(r["이름"])
        proc = "" if pd.isna(r["시술명"]) else str(r["시술명"])
        return f"{reg} {name}<br/>{proc}"

    room2_in_text = "<br/><br/>".join(
        patient_text(r) for _, r in room2_in_df.iterrows()
    ) if not room2_in_df.empty else "-"

    room1_in_text = "<br/><br/>".join(
        patient_text(r) for _, r in room1_in_df.iterrows()
    ) if not room1_in_df.empty else "-"

    room2_call_rows = [patient_text(r) for _, r in room2_call_df.iterrows()]
    room1_call_rows = [patient_text(r) for _, r in room1_call_df.iterrows()]

    max_call_len = max(len(room2_call_rows), len(room1_call_rows), 1)

    call_rows_html = ""
    for idx in range(max_call_len):
        room2_val = room2_call_rows[idx] if idx < len(room2_call_rows) else "&nbsp;"
        room1_val = room1_call_rows[idx] if idx < len(room1_call_rows) else "&nbsp;"
        label_text = "호출/도착" if idx == 0 else "&nbsp;"
        label_class = "section-label" if idx == 0 else "blank-label"

        call_rows_html += f"""
        <tr>
            <td class="{label_class}">{label_text}</td>
            <td class="patient-cell">{room2_val}</td>
            <td class="patient-cell">{room1_val}</td>
        </tr>
        """

    board_html = f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8"/>
    <style>
        body {{
            margin: 0;
            font-family: Arial, sans-serif;
            background: white;
            color: #111827;
        }}

        .wrapper {{
            padding: 10px 12px 20px 12px;
            position: relative;
            min-height: 860px;
            box-sizing: border-box;
        }}

        .title {{
            font-size: 36px;
            font-weight: 800;
            text-align: center;
            margin-bottom: 10px;
        }}

        .topbar {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            align-items: center;
            margin-bottom: 18px;
        }}

        .progress {{
            font-size: 24px;
            font-weight: 800;
            text-align: left;
        }}

        .datetime-wrap {{
            display: flex;
            justify-content: flex-end;
            align-items: baseline;
            gap: 14px;
        }}

        .board-date {{
            font-size: 34px;
            font-weight: 800;
            white-space: nowrap;
        }}

        .clock {{
            font-size: 30px;
            font-weight: 800;
            letter-spacing: 1px;
            font-variant-numeric: tabular-nums;
            white-space: nowrap;
        }}

        table {{
            width: 100%;
            border-collapse: collapse;
            table-layout: fixed;
        }}

        th, td {{
            border: 2px solid #cbd5e1;
            padding: 14px 10px;
            text-align: center;
            vertical-align: middle;
        }}

        th {{
            background: #cbd5e1;
            font-size: 28px;
            font-weight: 800;
        }}

        .section-label {{
            background: #e5eef9;
            color: #006400;
            font-size: 35px;
            font-weight: 800;
            width: 20%;
        }}

        .blank-label {{
            background: white;
            color: white;
            width: 20%;
        }}

        .patient-cell {{
            background: #f8fafc;
            font-size: 35px;
            font-weight: 700;
            line-height: 1.5;
            word-break: keep-all;
            white-space: normal;
        }}

        .spacer-row td {{
            height: 50px;
            background: white;
            padding: 0;
            border-left: 2px solid #cbd5e1;
            border-right: 2px solid #cbd5e1;
            border-top: 0;
            border-bottom: 0;
        }}

    </style>
</head>
<body>
    <div class="wrapper">
        <div class="title">실시간 시술현황</div>

        <div class="topbar">
            <div class="progress">진행도: {done_count}/{total_count} ({progress_pct}%)</div>
            <div class="datetime-wrap">
                <div class="board-date">{board_date_text}</div>
                <div class="clock" id="clock">--:--:--</div>
            </div>
        </div>

        <table>
            <colgroup>
                <col style="width:18%;">
                <col style="width:41%;">
                <col style="width:41%;">
            </colgroup>
            <tr>
                <th> </th>
                <th>2번방 (Single plane)</th>
                <th>1번방 (Bi plane)</th>
            </tr>

            <tr>
                <td class="section-label">입실/시술중</td>
                <td class="patient-cell">{room2_in_text}</td>
                <td class="patient-cell">{room1_in_text}</td>
            </tr>

            <tr class="spacer-row">
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
            </tr>

            {call_rows_html}
        </table>
    </div>

    <script>
        function updateClock() {{
            const now = new Date();
            const hh = String(now.getHours()).padStart(2, '0');
            const mm = String(now.getMinutes()).padStart(2, '0');
            const ss = String(now.getSeconds()).padStart(2, '0');
            document.getElementById('clock').textContent = hh + ':' + mm + ':' + ss;
        }}
        updateClock();
        setInterval(updateClock, 1000);
    </script>
</body>
</html>
"""

    components.html(board_html, height=900, scrolling=True)
    st.stop()
# -------------------------
# 환자 history 페이지
# -------------------------
if history_reg:
    df = st.session_state["procedures"].copy()

    patient_df = df[df["등록번호"].astype(str).str.strip() == str(history_reg).strip()].copy()

    if patient_df.empty:
        st.title("환자 시술 History")
        st.warning("해당 등록번호의 시술 기록이 없습니다.")

        if st.button("⬅ 뒤로가기"):
            if history_back_date:
                st.query_params["date"] = history_back_date
            clear_query_param("history")
            clear_query_param("history_date")
            st.rerun()

        st.stop()

    patient_df["날짜_dt"] = pd.to_datetime(patient_df["날짜"], errors="coerce")
    patient_df = patient_df.sort_values(["날짜_dt", "순서"], ascending=[False, False]).copy()

    latest_row = patient_df.iloc[0]
    patient_title = f"{latest_row['등록번호']} | {latest_row['이름']} | {latest_row['병실']}"

    top1, top2 = st.columns([1.2, 8.8])
    with top1:
        if st.button("⬅ 뒤로가기", use_container_width=True):
            replace_query_params(date=history_back_date if history_back_date else None)
            st.rerun()

    st.title("환자 시술 History")
    st.markdown(
        f"""
        <div style="font-size:1.2rem; font-weight:700; margin-bottom:1rem;">
            {patient_title}
        </div>
        """,
        unsafe_allow_html=True
    )

    history_show_cols = [
        "시술과", "시술명", "의뢰과", "의뢰의",
        "Room", "교수", "응급", "감염메모", "ADR메모", "Cr", "PT_INR", "PLT", "메모"
    ]

    for _, row in patient_df.iterrows():
        date_text = format_date_with_weekday(str(row["날짜"]))
        term_text = format_term_from_today(str(row["날짜"]))
        patient_info = f"{row['등록번호']} | {row['이름']} | {row['병실']}"

        st.markdown("---")
        st.markdown(
            f"""
            <div style="font-size:1.05rem; font-weight:700; margin-bottom:0.15rem;">
                {patient_info}
            </div>
            <div style="font-size:0.95rem; color:#374151; margin-bottom:0.1rem;">
                {date_text}
            </div>
            <div style="font-size:0.9rem; color:#6b7280; margin-bottom:0.5rem;">
                {term_text}
            </div>
            """,
            unsafe_allow_html=True
        )

        one_row = {}
        for col in history_show_cols:
            value = row[col] if col in row.index else ""
            one_row[col] = [value]

        show_df = pd.DataFrame(one_row)
        st.dataframe(show_df, use_container_width=True, hide_index=True)

    st.stop()

# -------------------------
# 시술 목록 페이지
# -------------------------
if selected_date and not history_reg:
    st.title(f"📅 {selected_date} 시술 목록")

    df = st.session_state["procedures"]
    day_all_df = df[df["날짜"] == selected_date].copy()

    display_df = day_all_df.copy()

    status_priority_map = {
        STATUS_INROOM: 0,   # 입실 최상단
        STATUS_ARRIVED: 1,  # 그다음 도착
        STATUS_CALLED: 2,   # 그다음 호출
        STATUS_PLANNED: 3,  # 그다음 예정
        STATUS_DONE: 4,     # 완료 최하단
    }

    display_df["status_priority"] = display_df["진행상황"].map(status_priority_map).fillna(99)

    display_df["IR_priority"] = display_df["시술과"].apply(
        lambda x: 0 if str(x).strip().upper() == "IR" else 1
    )

    display_df = display_df.sort_values(["status_priority", "IR_priority", "순서"])
    display_df = display_df.drop(columns=["status_priority", "IR_priority"])

    total_count = len(day_all_df)
    inroom_count = len(day_all_df[day_all_df["진행상황"] == STATUS_INROOM])
    called_count = len(day_all_df[day_all_df["진행상황"] == STATUS_CALLED])
    planned_count = len(day_all_df[day_all_df["진행상황"] == STATUS_PLANNED])
    done_count = len(day_all_df[day_all_df["진행상황"] == STATUS_DONE])

    top_left, top_board, top_spacer, top_btn1, top_btn2, top_btn2n, top_btn3 = st.columns(
    [6.0, 1.5, 0.2, 1.8, 1.5, 1.6, 2.0])

    with top_left:
        if st.button("⬅ 캘린더로 돌아가기", use_container_width=False):
            clear_query_param("date")
            clear_query_param("board")
            clear_query_param("duty")
            st.rerun()

    with top_board:
        board_url = app_query_string(board=selected_date)
        st.markdown(
            f"""
            <a href="{board_url}" target="_blank" style="text-decoration:none;">
                <div style="display:flex; align-items:center; justify-content:center; width:100%; min-height:2.5rem; padding:0.25rem 0.75rem; border:1px solid rgba(49, 51, 63, 0.2); border-radius:0.5rem; background:#ffffff; color:#111827; font-weight:600; cursor:pointer;">📺 현황판</div>
            </a>
            """,
            unsafe_allow_html=True,
        )

    with top_btn1:
        if st.button("➕ 직접 입력", use_container_width=False):
            add_procedure(selected_date)

    with top_btn2:
        if st.button("📋 EMR-IR", use_container_width=False):
            paste_emr_dialog(selected_date)

    with top_btn2n:
        if st.button("📋 EMR-N", use_container_width=False):
            paste_emr_n_dialog(selected_date)

    with top_btn3:
        if st.button("🔄 새로고침"):
            st.session_state["procedures"] = load_data()
            st.session_state["vacation_notes"] = load_vacation_data()
            st.rerun()

    row2_left, row2_right = st.columns([1.4, 8.0])

    with row2_left:
        if st.button("🏅 Ranking", use_container_width=False):
            ranking_dialog(selected_date)

    with row2_right:
        st.markdown(
            f"""
            <div style="padding-top: 0.15rem; font-size: 2.0rem; font-weight: 700; line-height: 1.3;">
                <span style="color: #000000;">Total {total_count}</span>
                <span style="color: #000000;"> = </span>
                <span style="color: {STATUS_COLORS[STATUS_INROOM]};">입실</span>
                <span style="color: #000000;"> {inroom_count} </span>
                <span style="color: #000000;"> + </span>
                <span style="color: {STATUS_COLORS[STATUS_CALLED]};">호출</span>
                <span style="color: #000000;"> {called_count} </span>
                <span style="color: #000000;"> + </span>
                <span style="color: {STATUS_COLORS[STATUS_PLANNED]};">예정</span>
                <span style="color: #000000;"> {planned_count} </span>
                <span style="color: #000000;"> + </span>
                <span style="color: {STATUS_COLORS[STATUS_DONE]};">완료</span>
                <span style="color: #000000;"> {done_count}</span>
            </div>
            """,
            unsafe_allow_html=True
        )
    st.markdown("---")

    col_widths = [0.35, 1.85, 1.00, 0.70, 2.10, 0.50, 0.80, 2.70, 1.35, 0.45, 0.60, 0.35, 0.35, 0.40]
    h0, h1, h2, h3, h4, h5, h6, h7, h8, h9, h10, h11, h12, h13 = st.columns(col_widths)
    h0.markdown('<div class="nowrap-header"> </div>', unsafe_allow_html=True)
    h1.markdown('<div class="nowrap-header">등록번호/이름/병실</div>', unsafe_allow_html=True)
    h2.markdown('<div class="nowrap-header">시술전확인</div>', unsafe_allow_html=True)
    h3.markdown('<div class="nowrap-header">시술과</div>', unsafe_allow_html=True)
    h4.markdown('<div class="nowrap-header">시술명</div>', unsafe_allow_html=True)
    h5.markdown('<div class="nowrap-header">Room</div>', unsafe_allow_html=True)
    h6.markdown('<div class="nowrap-header">시술의</div>', unsafe_allow_html=True)
    h7.markdown('<div class="nowrap-header">진행상황</div>', unsafe_allow_html=True)
    h8.markdown('<div class="nowrap-header">의뢰과/의뢰의</div>', unsafe_allow_html=True)
    h9.markdown('<div class="nowrap-header">응급</div>', unsafe_allow_html=True)
    h10.markdown('<div class="nowrap-header">수정</div>', unsafe_allow_html=True)
    h11.markdown('<div class="nowrap-header">↑</div>', unsafe_allow_html=True)
    h12.markdown('<div class="nowrap-header">↓</div>', unsafe_allow_html=True)
    h13.markdown('<div class="nowrap-header">삭제</div>', unsafe_allow_html=True)

    st.markdown("---")

    day_df = get_display_day_df(df, selected_date)

    if len(day_df) == 0:
        st.info("등록된 시술 없음")
    else:
        for i, row in display_df.iterrows():
            row_id = row["id"]
            is_done = row["진행상황"] == STATUS_DONE

            is_done = row["진행상황"] == STATUS_DONE
            is_inroom = row["진행상황"] == STATUS_INROOM

            if is_done:
                row_key = f"row_done_{row_id}"
            elif is_inroom:
                row_key = f"row_inroom_{row_id}"
            else:
                row_key = f"row_{row_id}"

            with st.container(key=row_key):
                col0, col1, col2, col3, col4, col5, col6, col7, col8, col9, col10, col11, col12, col13 = st.columns(col_widths)

                reg_no = "" if pd.isna(row["등록번호"]) else str(row["등록번호"]).strip()

                if reg_no:
                    if col0.button("🔍", key=f"history_btn_{selected_date}_{reg_no}_{i}"):
                        replace_query_params(date=selected_date, history=reg_no, history_date=selected_date)
                        st.rerun()
                else:
                    col0.write("")

                patient_info = f"{row['등록번호']} | {row['이름']} ({row['성별']}/{row['나이']}) | {row['병실']}"
                col1.write(patient_info)

            with col2:
                consent_on = bool(row["동의서"]) if "동의서" in row else False

                infection_on = bool(row["감염"]) if "감염" in row else False
                infection_memo = "" if pd.isna(row.get("감염메모", "")) else str(row.get("감염메모", "")).strip()

                adr_on = bool(row["ADR"]) if "ADR" in row else False
                adr_memo = "" if pd.isna(row.get("ADR메모", "")) else str(row.get("ADR메모", "")).strip()

                renal_on = bool(row["신기능"]) if "신기능" in row else False
                cr_value = "" if pd.isna(row.get("Cr", "")) else str(row.get("Cr", "")).strip()

                bleeding_on = bool(row["출혈"]) if "출혈" in row else False
                pt_inr_value = "" if pd.isna(row.get("PT_INR", "")) else str(row.get("PT_INR", "")).strip()
                plt_value = "" if pd.isna(row.get("PLT", "")) else str(row.get("PLT", "")).strip()

                consent_btn_type = "primary" if consent_on else "secondary"
                if st.button(
                    "동의(+)",
                    key=f"precheck_consent_{i}",
                    use_container_width=True,
                    type=consent_btn_type
                ):
                    toggle_consent(row_id)
                    st.rerun()

                infect_btn_type = "primary" if infection_on else "secondary"
                infect_c1, infect_c2 = st.columns([2.0, 1.0])
                with infect_c1:
                    if st.button(
                        "감염",
                        key=f"precheck_infect_btn_{i}",
                        use_container_width=True,
                        type=infect_btn_type
                    ):
                        st.session_state[f"infect_pop_open_{i}"] = not st.session_state.get(f"infect_pop_open_{i}", False)
                        st.rerun()

                with infect_c2:
                    if infection_on and infection_memo:
                        st.markdown(
                            f"<div style='font-size:0.82rem; padding-top:0.28rem; color:#444; white-space:nowrap;'>{infection_memo}</div>",
                            unsafe_allow_html=True
                        )        

                if st.session_state.get(f"infect_pop_open_{i}", False):
                    infection_input = st.text_input(
                        "감염내용",
                        value=infection_memo,
                        key=f"infection_text_{i}",
                        label_visibility="collapsed",
                    )

                    if st.button("확인", key=f"infection_save_{i}", use_container_width=True):
                        save_infection_info(row_id, st.session_state.get(f"infection_text_{i}", ""))
                        st.session_state[f"infect_pop_open_{i}"] = False
                        st.rerun()

                adr_btn_type = "primary" if adr_on else "secondary"
                if st.button(
                    "ADR",
                    key=f"precheck_adr_btn_{i}",
                    use_container_width=True,
                    type=adr_btn_type,
                    help=adr_memo if adr_memo else None
                ):
                    st.session_state[f"adr_pop_open_{i}"] = not st.session_state.get(f"adr_pop_open_{i}", False)
                    st.rerun()

                if st.session_state.get(f"adr_pop_open_{i}", False):
                    adr_input = st.text_input(
                        "ADR내용",
                        value=adr_memo,
                        key=f"adr_text_{i}",
                        label_visibility="collapsed",
                    )

                    if st.button("확인", key=f"adr_save_{i}", use_container_width=True):
                        save_adr_info(row_id, st.session_state.get(f"adr_text_{i}", ""))
                        st.session_state[f"adr_pop_open_{i}"] = False
                        st.rerun()

                renal_btn_type = "primary" if renal_on else "secondary"
                if st.button(
                    "신기능",
                    key=f"precheck_renal_btn_{i}",
                    use_container_width=True,
                    type=renal_btn_type,
                    help=f"Cr: {cr_value}" if cr_value else None
                ):
                    st.session_state[f"renal_pop_open_{i}"] = not st.session_state.get(f"renal_pop_open_{i}", False)
                    st.rerun()

                if st.session_state.get(f"renal_pop_open_{i}", False):
                    r_label, r_input = st.columns([1, 2])
                    r_label.markdown("<div style='padding-top:0.42rem; font-weight:600;'>Cr</div>", unsafe_allow_html=True)
                    r_input.text_input(
                        "Cr",
                        value=cr_value,
                        key=f"renal_cr_text_{i}",
                        label_visibility="collapsed",
                    )

                    if st.button("확인", key=f"renal_save_{i}", use_container_width=True):
                        save_renal_info(row_id, st.session_state.get(f"renal_cr_text_{i}", ""))
                        st.session_state[f"renal_pop_open_{i}"] = False
                        st.rerun()

                bleeding_help = None
                if pt_inr_value or plt_value:
                    bleeding_help = f"PT INR: {pt_inr_value} / Plt: {plt_value}"

                bleeding_btn_type = "primary" if bleeding_on else "secondary"
                if st.button(
                    "출혈",
                    key=f"precheck_bleeding_btn_{i}",
                    use_container_width=True,
                    type=bleeding_btn_type,
                    help=bleeding_help
                ):
                    st.session_state[f"bleeding_pop_open_{i}"] = not st.session_state.get(f"bleeding_pop_open_{i}", False)
                    st.rerun()

                if st.session_state.get(f"bleeding_pop_open_{i}", False):
                    b1, b2 = st.columns([1.2, 1.8])
                    b1.markdown("<div style='padding-top:0.42rem; font-weight:600;'>PT INR</div>", unsafe_allow_html=True)
                    b2.text_input(
                        "PT INR",
                        value=pt_inr_value,
                        key=f"bleeding_pt_inr_text_{i}",
                        label_visibility="collapsed",
                    )

                    b3, b4 = st.columns([1.2, 1.8])
                    b3.markdown("<div style='padding-top:0.42rem; font-weight:600;'>Plt</div>", unsafe_allow_html=True)
                    b4.text_input(
                        "Plt",
                        value=plt_value,
                        key=f"bleeding_plt_text_{i}",
                        label_visibility="collapsed",
                    )

                    if st.button("확인", key=f"bleeding_save_{i}", use_container_width=True):
                        save_bleeding_info(
                            row_id,
                            st.session_state.get(f"bleeding_pt_inr_text_{i}", ""),
                            st.session_state.get(f"bleeding_plt_text_{i}", "")
                        )
                        st.session_state[f"bleeding_pop_open_{i}"] = False
                        st.rerun()

            proc_dept_value = row["시술과"] if row["시술과"] in ["IR", "NS", "NU"] else "IR"
            memo_value = "" if pd.isna(row.get("메모", "")) else str(row.get("메모", ""))

            col3.write(proc_dept_value)
            procedure_text = "" if pd.isna(row["시술명"]) else str(row["시술명"])
            if memo_value.strip():
                col4.markdown(
                    f"<div>{procedure_text}</div><div style='font-size:0.9rem; color:#555; padding-top:0.22rem; white-space:pre-wrap;'>{memo_value}</div>",
                    unsafe_allow_html=True,
                )
            else:
                col4.write(procedure_text)


            col5.write("" if pd.isna(row["Room"]) else str(row["Room"]))
            col6.write("" if pd.isna(row["교수"]) else str(row["교수"]))

            current_status = row["진행상황"] if row["진행상황"] in VALID_STATUSES else STATUS_PLANNED

            with col7:
                s1, s2, s3, s4, s5 = st.columns(5)

                if current_status == STATUS_PLANNED:
                    with s1:
                        status_badge("예정", STATUS_PLANNED)
                else:
                    if s1.button("예정", key=f"status_plan_{i}", use_container_width=True):
                        set_status(row_id, STATUS_PLANNED)
                        st.rerun()

                if current_status == STATUS_CALLED:
                    with s2:
                        status_badge("호출", STATUS_CALLED)
                else:
                    if s2.button("호출", key=f"status_call_{i}", use_container_width=True):
                        set_status(row_id, STATUS_CALLED)
                        st.rerun()

                if current_status == STATUS_ARRIVED:
                    with s3:
                        status_badge("도착", STATUS_ARRIVED)
                else:
                    if s3.button("도착", key=f"status_arrived_{i}", use_container_width=True):
                        set_status(row_id, STATUS_ARRIVED)
                        st.rerun()
                        
                if current_status == STATUS_INROOM:
                    with s4:
                        status_badge("입실", STATUS_INROOM)
                else:
                    if s4.button("입실", key=f"status_in_{i}", use_container_width=True):
                        set_status(row_id, STATUS_INROOM)
                        st.rerun()

                if current_status == STATUS_DONE:
                    with s5:
                        status_badge("완료", STATUS_DONE)
                else:
                    if s5.button("완료", key=f"status_done_{i}", use_container_width=True):
                        set_status(row_id, STATUS_DONE)
                        st.rerun()

            col8.write(f"{row['의뢰과']} / {row['의뢰의']}")

            current_emergency = "🚑" if bool(row["응급"]) else "N"
            col9.write(current_emergency)

            if col10.button("수정", key=f"edit_{i}", use_container_width=True):
                edit_procedure_dialog(row_id)

            if col11.button("↑", key=f"up_{i}"):
                move_up(row_id)
                st.rerun()

            if col12.button("↓", key=f"down_{i}"):
                move_down(row_id)
                st.rerun()

            if col13.button("🗑", key=f"del_{i}"):
                delete_dialog(row_id)

            st.markdown("<hr style='margin:4px 0;'>", unsafe_allow_html=True)

    st.stop()

# -------------------------
# 메인 캘린더 페이지
# -------------------------
st.markdown(
    """
    <div style="text-align:center; font-size:2rem; font-weight:700; margin-bottom:0.4rem;">
        인터벤션 시술일정관리
    </div>
    """,
    unsafe_allow_html=True
)

top_btn_row_left, top_btn_row_mid, top_btn_row_rest = st.columns([1.2, 1.2, 7.6])

with top_btn_row_left:
    if st.button("🏅 Ranking", use_container_width=True):
        monthly_ranking_dialog(st.session_state["calendar_year"], st.session_state["calendar_month"])

with top_btn_row_mid:
    if st.button("📅 당직", use_container_width=True):
        replace_query_params(duty=f"{st.session_state['calendar_year']}-{st.session_state['calendar_month']:02d}")
        st.rerun()

year = st.session_state["calendar_year"]
month = st.session_state["calendar_month"]
df = st.session_state["procedures"]

month_total = get_month_case_total(df, year, month)

nav1, nav2, nav3 = st.columns([1, 6, 1])

with nav1:
    st.button("◀", on_click=prev_month, use_container_width=True)

with nav2:
    st.markdown(
        f"""
        <div style='display:flex; justify-content:center; align-items:center; gap:18px; margin-top:0.2rem;'>
            <div style='font-size:25pt; font-weight:700;'>
                {year}년 {month}월
            </div>
            <div style='font-size:22pt; font-weight:600; color:#0000FF;'>
                Total {month_total}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

with nav3:
    st.button("▶", on_click=next_month, use_container_width=True)

year = st.session_state["calendar_year"]
month = st.session_state["calendar_month"]

calendar.setfirstweekday(calendar.SUNDAY)
cal = calendar.monthcalendar(year, month)

weekdays = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]

cols = st.columns(7)
for i, day_name in enumerate(weekdays):
    color = "#111111"
    if day_name == "Sun":
        color = "#dc2626"
    elif day_name == "Sat":
        color = "#2563eb"

    cols[i].markdown(
        f"<div style='text-align:center; font-size:1.5rem; font-weight:700; color:{color};'>{day_name}</div>",
        unsafe_allow_html=True
    )

df = st.session_state["procedures"]

today = datetime.today().date()

for week in cal:
    cols = st.columns(7)

    for i, day_num in enumerate(week):
        if day_num == 0:
            cols[i].write("")
            continue

        date_str = f"{year}-{month:02d}-{day_num:02d}"
        this_date = datetime(year, month, day_num).date()

        total_count, ir_count, ns_count, nu_count, neuro_total = get_day_case_summary(df, date_str)

        memo_text, _ = get_vacation_note(date_str)

        content = (
            f"{day_num}\n"
            f"{total_count}\n"
            f"IR {ir_count}\n"
            f"N {neuro_total}\n"
            f"(NS {ns_count} NU {nu_count})\n"
        )

        is_today = this_date == today

        if is_today:
            if i == 0:
                cal_key = f"cal_today_sun_{year}_{month:02d}_{day_num:02d}"
            elif i == 6:
                cal_key = f"cal_today_sat_{year}_{month:02d}_{day_num:02d}"
            else:
                cal_key = f"cal_today_{year}_{month:02d}_{day_num:02d}"
        else:
            if i == 0:
                cal_key = f"cal_sun_{year}_{month:02d}_{day_num:02d}"
            elif i == 6:
                cal_key = f"cal_sat_{year}_{month:02d}_{day_num:02d}"
            else:
                cal_key = f"cal_{year}_{month:02d}_{day_num:02d}"

        with cols[i]:
            if st.button(content, key=cal_key, use_container_width=True):
                replace_query_params(date=date_str)
                st.rerun()

            plus_col, note_col = st.columns([1, 5])

            with plus_col:
                if st.button("+", key=f"vac_plus_{date_str}", use_container_width=True):
                    popup_key = f"vac_popup_text_{date_str}"
                    st.session_state[popup_key] = memo_text
                    vacation_note_dialog(date_str)

            with note_col:
                if memo_text.strip():
                    st.markdown(
                        f"""
                        <div style="
                            font-size:0.78rem;
                            line-height:1.2;
                            padding-top:0.35rem;
                            white-space:normal;
                            word-break:keep-all;
                            min-height:1.5rem;
                        ">
                            {memo_text}
                        </div>
                        """,
                        unsafe_allow_html=True
                    )
                else:
                    st.markdown(
                        """
                        <div style="min-height:1.5rem;"></div>
                        """,
                        unsafe_allow_html=True
                    )

render_logout()
st.write(f"접속자: {st.session_state['username']}")

