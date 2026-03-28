import io
import json
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
from datetime import datetime
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from google.oauth2 import service_account

# ── Page Config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="IoT Streetlight Dashboard",
    page_icon="💡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Dark Theme CSS ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
body, .stApp { background-color: #0f172a !important; color: #f1f5f9; }
[data-testid="stSidebar"] { background-color: #1e293b !important; }
[data-testid="stSidebar"] * { color: #e2e8f0 !important; }
.kpi-row { display: flex; gap: 16px; margin-bottom: 24px; }
.kpi-card {
    flex: 1;
    background: #1e293b;
    border-radius: 10px;
    padding: 22px 20px;
    border: 1px solid #334155;
}
.kpi-label { color: #94a3b8; font-size: 14px; margin-bottom: 6px; }
.kpi-value { color: #f8fafc; font-size: 38px; font-weight: 700; letter-spacing: -1px; }
.section-title {
    color: #f1f5f9;
    font-size: 18px;
    font-weight: 600;
    margin: 8px 0 12px 0;
}
.legend-row {
    display: flex;
    gap: 20px;
    align-items: center;
    justify-content: flex-end;
    margin-bottom: 8px;
}
.legend-item { display: flex; align-items: center; gap: 6px; font-size: 13px; color: #cbd5e1; }
.leg-dot { width: 12px; height: 12px; border-radius: 3px; display: inline-block; }
table.dt { width: 100%; border-collapse: collapse; font-size: 14px; margin-top: 4px; }
table.dt th {
    background: #1e293b;
    color: #94a3b8;
    padding: 10px 14px;
    text-align: right;
    border-bottom: 2px solid #334155;
    white-space: nowrap;
    font-weight: 500;
}
table.dt th:first-child, table.dt th:nth-child(2) { text-align: left; }
table.dt td {
    padding: 10px 14px;
    border-bottom: 1px solid #1e293b;
    color: #f1f5f9;
    text-align: right;
    vertical-align: middle;
}
table.dt td:first-child, table.dt td:nth-child(2) { text-align: left; }
table.dt tr:hover td { background: #1e293b; }
.bar-wrap { width: 130px; height: 16px; background: #374151; border-radius: 4px; overflow: hidden; display: inline-block; vertical-align: middle; }
.bar-fill { height: 100%; border-radius: 4px; }
.c-green  { color: #22c55e; }
.c-orange { color: #f97316; }
.c-red    { color: #ef4444; }
.c-gray   { color: #64748b; }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
#  GOOGLE DRIVE
# ══════════════════════════════════════════════════════════════════════════════

SCOPES = ["https://www.googleapis.com/auth/drive"]


@st.cache_resource
def _drive_service():
    gd = st.secrets["gdrive"]
    info = {
        "type": "service_account",
        "project_id": gd["project_id"],
        "private_key_id": gd["private_key_id"],
        "private_key": gd["private_key"],
        "client_email": gd["client_email"],
        "client_id": gd["client_id"],
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token",
        "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
        "client_x509_cert_url": (
            "https://www.googleapis.com/robot/v1/metadata/x509/"
            + gd["client_email"].replace("@", "%40")
        ),
    }
    creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
    return build("drive", "v3", credentials=creds)


def _folder_id() -> str:
    return st.secrets["gdrive"]["folder_id"]


@st.cache_data(ttl=5)
def _find_file_id(name: str):
    """Find a file's ID in the Drive folder by name. Cached 5 sec."""
    svc = _drive_service()
    fid = _folder_id()
    q = f"name='{name}' and '{fid}' in parents and trashed=false"
    r = svc.files().list(q=q, fields="files(id)", pageSize=1).execute()
    files = r.get("files", [])
    return files[0]["id"] if files else None


def _read_json(name: str):
    """Read a JSON file from Drive. Returns None if not found."""
    file_id = _find_file_id(name)
    if not file_id:
        return None
    svc = _drive_service()
    req = svc.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    dl = MediaIoBaseDownload(buf, req)
    done = False
    while not done:
        _, done = dl.next_chunk()
    buf.seek(0)
    return json.loads(buf.read().decode("utf-8"))


def _write_json(name: str, data):
    """Write (create or update) a JSON file in Drive."""
    svc = _drive_service()
    content = json.dumps(data, ensure_ascii=False).encode("utf-8")
    media = MediaIoBaseUpload(
        io.BytesIO(content), mimetype="application/json", resumable=False
    )
    file_id = _find_file_id(name)
    if file_id:
        svc.files().update(fileId=file_id, media_body=media).execute()
    else:
        meta = {"name": name, "parents": [_folder_id()], "mimeType": "application/json"}
        svc.files().create(body=meta, media_body=media, fields="id").execute()
    _find_file_id.clear()  # clear cache so next read finds latest


def _delete_drive_file(name: str):
    """Delete a file from Drive by name (silently if not found)."""
    file_id = _find_file_id(name)
    if file_id:
        _drive_service().files().delete(fileId=file_id).execute()
        _find_file_id.clear()


def _load_db() -> dict:
    """Load db.json from Drive (or return empty structure)."""
    data = _read_json("db.json")
    if data is None:
        return {"uploads": [], "rows": []}
    return data


def _save_db(db: dict):
    _write_json("db.json", db)


# ══════════════════════════════════════════════════════════════════════════════
#  DATA HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def _next_id(arr: list) -> int:
    return max((x["id"] for x in arr), default=0) + 1


def get_uploads() -> list:
    """Return upload list (newest first)."""
    db = _load_db()
    return list(reversed(db["uploads"]))


def get_data(upload_id: int) -> pd.DataFrame:
    db = _load_db()
    rows = [r for r in db["rows"] if r["upload_id"] == upload_id]
    return pd.DataFrame(rows) if rows else pd.DataFrame(
        columns=["id","upload_id","contractor","contract_no",
                 "max_poles","actual_poles","iot_available","api_connected"]
    )


def save_upload(df: pd.DataFrame, filename: str) -> int:
    db = _load_db()
    now = datetime.now()
    uid = _next_id(db["uploads"])
    db["uploads"].append({
        "id": uid,
        "upload_date": now.strftime("%Y-%m-%d"),
        "upload_time": now.strftime("%H:%M:%S"),
        "filename": filename,
    })
    for _, r in df.iterrows():
        db["rows"].append({
            "id": _next_id(db["rows"]),
            "upload_id": uid,
            "contractor":    str(r["contractor"]),
            "contract_no":   str(r["contract_no"]),
            "max_poles":     int(r["max_poles"]),
            "actual_poles":  int(r["actual_poles"]),
            "iot_available": int(r["iot_available"]),
            "api_connected": int(r["api_connected"]),
        })
    _save_db(db)
    return uid


def delete_upload(upload_id: int):
    db = _load_db()
    db["uploads"] = [u for u in db["uploads"] if u["id"] != upload_id]
    db["rows"]    = [r for r in db["rows"]    if r["upload_id"] != upload_id]
    _save_db(db)
    _delete_drive_file(f"raw_{upload_id}.json")


# ══════════════════════════════════════════════════════════════════════════════
#  EXCEL PARSING
# ══════════════════════════════════════════════════════════════════════════════

COLUMN_HINTS = {
    "contractor":    ["ผู้รับเหมา", "contractor", "บริษัท", "ผู้รับจ้าง", "ผรม"],
    "contract_no":   ["เลขที่สัญญา", "สัญญา", "contract", "contract_no", "เลขสัญญา"],
    "max_poles":     ["max โคม", "maxโคม", "โคมสูงสุด", "max", "จำนวนสัญญา"],
    "actual_poles":  ["จำนวนจริง", "จำนวนเสาจริง", "actual", "จำนวนเสา", "จำนวน"],
    "iot_available": ["iot ได้", "iotได้", "ควบคุม iot", "iot", "เสา iot"],
    "api_connected": ["api เชื่อม", "apiเชื่อม", "เชื่อม api", "api", "เชื่อมต่อ"],
}


def detect_columns(columns) -> dict:
    mapping = {}
    for col in columns:
        col_l = str(col).lower().strip()
        for field, hints in COLUMN_HINTS.items():
            if field in mapping:
                continue
            if any(h.lower() in col_l or col_l in h.lower() for h in hints):
                mapping[field] = col
    return mapping


def parse_excel(file):
    """Returns (df, error_str). df is None on failure."""
    try:
        raw = pd.read_excel(file)
    except Exception as e:
        return None, f"ไม่สามารถอ่านไฟล์ได้: {e}"

    mapping = detect_columns(raw.columns)
    missing = [f for f in COLUMN_HINTS if f not in mapping]
    if missing:
        readable = {
            "contractor": "ผู้รับเหมา", "contract_no": "เลขที่สัญญา",
            "max_poles": "max โคม",     "actual_poles": "จำนวนจริง",
            "iot_available": "IOT ได้", "api_connected": "API เชื่อม",
        }
        miss_names = [readable.get(m, m) for m in missing]
        return None, (
            f"ไม่พบคอลัมน์ที่ต้องการ: **{', '.join(miss_names)}**\n\n"
            f"คอลัมน์ที่พบในไฟล์: {list(raw.columns)}\n\n"
            "กรุณาดาวน์โหลด Template เพื่อดูรูปแบบที่ถูกต้อง"
        )

    df = raw.rename(columns={v: k for k, v in mapping.items()})[list(COLUMN_HINTS.keys())]
    df = df.dropna(how="all")
    for c in ["max_poles", "actual_poles", "iot_available", "api_connected"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)
    df["contractor"]  = df["contractor"].fillna("").astype(str)
    df["contract_no"] = df["contract_no"].fillna("").astype(str).str.strip()
    df = df[df["contract_no"] != ""]
    return df, None


# ── Template ───────────────────────────────────────────────────────────────────
def make_template() -> bytes:
    sample = pd.DataFrame({
        "ผู้รับเหมา":   ["บริษัท ตัวอย่าง จำกัด", "บริษัท ABC จำกัด"],
        "เลขที่สัญญา": ["สกบ.1/2568", "สกบ.2/2568"],
        "max โคม":     [100, 50],
        "จำนวนจริง":   [100, 48],
        "IOT ได้":     [80, 0],
        "API เชื่อม":  [75, 0],
    })
    buf = io.BytesIO()
    sample.to_excel(buf, index=False)
    return buf.getvalue()


# ── Export Excel ───────────────────────────────────────────────────────────────
def export_excel(df: pd.DataFrame, group_by: str) -> bytes:
    df = df.copy()
    df["ส่วนต่าง"]    = df["actual_poles"] - df["max_poles"]
    df["ยังไม่เชื่อม"] = df["iot_available"] - df["api_connected"]

    export = df.rename(columns={
        "contractor":    "ผู้รับเหมา",
        "contract_no":   "เลขที่สัญญา",
        "max_poles":     "max โคม",
        "actual_poles":  "จำนวนจริง",
        "iot_available": "IOT ได้",
        "api_connected": "API เชื่อม",
    })[["ผู้รับเหมา", "เลขที่สัญญา", "max โคม", "จำนวนจริง",
        "ส่วนต่าง", "IOT ได้", "API เชื่อม", "ยังไม่เชื่อม"]]

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        if group_by == "contractor":
            for name, grp in export.groupby("ผู้รับเหมา"):
                grp.to_excel(writer, sheet_name=str(name)[:31], index=False)
        else:
            export.to_excel(writer, sheet_name="รายเลขที่สัญญา", index=False)
    return buf.getvalue()


# ── Table HTML ─────────────────────────────────────────────────────────────────
def render_table(df: pd.DataFrame):
    df = df.copy()
    df["diff"]          = df["actual_poles"] - df["max_poles"]
    df["not_connected"] = df["iot_available"] - df["api_connected"]
    df["iot_ratio"]     = df.apply(
        lambda r: r["iot_available"] / r["actual_poles"] if r["actual_poles"] > 0 else 0,
        axis=1,
    )

    def row_status(r):
        if r["iot_available"] > 0 and r["api_connected"] > 0:
            return "green"
        elif r["iot_available"] > 0:
            return "orange"
        return "red"

    rows_html = []
    for _, r in df.iterrows():
        if r["diff"] > 0:
            diff = f'<span class="c-green">+{r["diff"]:,}</span>'
        elif r["diff"] < 0:
            diff = f'<span class="c-red">{r["diff"]:,}</span>'
        else:
            diff = '<span class="c-gray">0</span>'

        status  = row_status(r)
        iot_cls = "c-green"  if r["iot_available"] > 0  else "c-gray"
        api_cls = "c-green"  if r["api_connected"] > 0  else "c-gray"
        nc_cls  = "c-orange" if r["not_connected"] > 0  else "c-gray"

        pct = min(r["iot_ratio"] * 100, 100)
        bar_colors = {"green": "#22c55e", "orange": "#f97316", "red": "#ef4444"}
        bar = (
            f'<div class="bar-wrap">'
            f'<div class="bar-fill" style="width:{pct:.0f}%;background:{bar_colors[status]}"></div>'
            f'</div>'
        )

        rows_html.append(f"""
        <tr>
          <td>{r["contractor"]}</td>
          <td>{r["contract_no"]}</td>
          <td>{r["max_poles"]:,}</td>
          <td>{r["actual_poles"]:,}</td>
          <td>{diff}</td>
          <td><span class="{iot_cls}">{r["iot_available"]:,}</span></td>
          <td><span class="{api_cls}">{r["api_connected"]:,}</span></td>
          <td><span class="{nc_cls}">{r["not_connected"]:,}</span></td>
          <td>{bar}</td>
        </tr>""")

    rows_joined = "".join(rows_html)
    html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
  body {{ margin:0; padding:0; background:transparent; color:#f1f5f9;
          font-family:'Segoe UI',sans-serif; font-size:14px; }}
  .legend-row {{ display:flex; gap:20px; align-items:center; justify-content:flex-end;
                 margin-bottom:8px; }}
  .legend-item {{ display:flex; align-items:center; gap:6px; font-size:13px; color:#cbd5e1; }}
  .leg-dot {{ width:12px; height:12px; border-radius:3px; display:inline-block; }}
  table {{ width:100%; border-collapse:collapse; }}
  th {{ background:#1e293b; color:#94a3b8; padding:10px 14px; text-align:right;
        border-bottom:2px solid #334155; white-space:nowrap; font-weight:500; }}
  th:first-child, th:nth-child(2) {{ text-align:left; }}
  td {{ padding:10px 14px; border-bottom:1px solid #0f172a; text-align:right;
        vertical-align:middle; }}
  td:first-child, td:nth-child(2) {{ text-align:left; }}
  tr:hover td {{ background:#1e293b; }}
  .bar-wrap {{ width:130px; height:16px; background:#374151; border-radius:4px;
               overflow:hidden; display:inline-block; vertical-align:middle; }}
  .bar-fill {{ height:100%; border-radius:4px; }}
  .c-green  {{ color:#22c55e; }}
  .c-orange {{ color:#f97316; }}
  .c-red    {{ color:#ef4444; }}
  .c-gray   {{ color:#64748b; }}
</style>
</head>
<body>
<div class="legend-row">
  <div class="legend-item"><span class="leg-dot" style="background:#22c55e"></span> IOT+API OK</div>
  <div class="legend-item"><span class="leg-dot" style="background:#f97316"></span> IOT OK ยังไม่ API</div>
  <div class="legend-item"><span class="leg-dot" style="background:#ef4444"></span> IOT ไม่ได้</div>
</div>
<table>
  <thead>
    <tr>
      <th>ผู้รับเหมา</th><th>เลขที่สัญญา</th><th>max โคม</th><th>จำนวนจริง</th>
      <th>ส่วนต่าง</th><th>IOT ได้</th><th>API เชื่อม</th><th>ยังไม่เชื่อม</th><th>สัดส่วน IOT</th>
    </tr>
  </thead>
  <tbody>{rows_joined}</tbody>
</table>
</body>
</html>"""
    height = max(300, len(df) * 41 + 120)
    components.html(html, height=height, scrolling=True)


# ══════════════════════════════════════════════════════════════════════════════
#  SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## 💡 IoT Streetlight")
    st.markdown("---")

    # ── Upload ──────────────────────────────────────────────────
    st.markdown("### 📤 อัพโหลดข้อมูล")
    uploaded_file = st.file_uploader("เลือกไฟล์ Excel (.xlsx)", type=["xlsx", "xls"])

    if uploaded_file:
        parsed_df, parse_err = parse_excel(uploaded_file)
        if parse_err:
            st.error(parse_err)
        else:
            st.success(f"✅ พบข้อมูล {len(parsed_df):,} แถว")
            if st.button("💾 บันทึกข้อมูล", type="primary", use_container_width=True):
                with st.spinner("กำลังบันทึกไปยัง Google Drive..."):
                    save_upload(parsed_df, uploaded_file.name)
                st.success("บันทึกเรียบร้อยแล้ว!")
                st.rerun()

    st.download_button(
        label="📋 ดาวน์โหลด Template Excel",
        data=make_template(),
        file_name="template_streetlight.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    st.markdown("---")

    # ── Select Upload ────────────────────────────────────────────
    st.markdown("### 📅 เลือกข้อมูล")
    uploads = get_uploads()
    selected_uid = None

    if not uploads:
        st.info("ยังไม่มีข้อมูล\nกรุณาอัพโหลดไฟล์ Excel")
    else:
        options = {
            f"{u['upload_date']}  {u['upload_time']}  —  {u['filename']}": u["id"]
            for u in uploads
        }
        chosen_label = st.selectbox(
            "วันที่อัพโหลด", list(options.keys()), label_visibility="collapsed"
        )
        selected_uid = options[chosen_label]

        with st.expander("🗑 ลบข้อมูลนี้"):
            if st.button("ลบ", type="secondary", use_container_width=True):
                with st.spinner("กำลังลบ..."):
                    delete_upload(selected_uid)
                st.rerun()

    st.markdown("---")

    # ── Export ──────────────────────────────────────────────────
    if selected_uid is not None:
        st.markdown("### 📥 Export Excel")
        export_opt = st.radio(
            "แยกตาม",
            ["รายผู้รับเหมา (แยก sheet)", "รายเลขที่สัญญา (sheet เดียว)"],
        )
        group = "contractor" if "ผู้รับเหมา" in export_opt else "contract"
        data_for_export = get_data(selected_uid)
        st.download_button(
            label="⬇️ ดาวน์โหลด Excel",
            data=export_excel(data_for_export, group),
            file_name=f"export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("# IoT Streetlight Dashboard")

if selected_uid is None:
    st.markdown("""
    <div style="text-align:center;padding:80px 0;color:#64748b;">
        <div style="font-size:48px;">💡</div>
        <div style="font-size:18px;margin-top:12px;">
            กรุณาอัพโหลดไฟล์ Excel จาก Sidebar ด้านซ้าย
        </div>
        <div style="font-size:14px;margin-top:8px;">
            ดาวน์โหลด Template เพื่อดูรูปแบบคอลัมน์ที่รองรับ
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

data_df = get_data(selected_uid)

if data_df.empty:
    st.warning("ไม่พบข้อมูลสำหรับการอัพโหลดนี้")
    st.stop()

# ── KPI Cards ─────────────────────────────────────────────────────────────────
total_contracts = len(data_df)
total_poles     = int(data_df["actual_poles"].sum())
total_iot       = int(data_df["iot_available"].sum())
total_api       = int(data_df["api_connected"].sum())

st.markdown(f"""
<div class="kpi-row">
  <div class="kpi-card">
    <div class="kpi-label">สัญญา</div>
    <div class="kpi-value">{total_contracts:,}</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-label">เสาไฟจริง</div>
    <div class="kpi-value">{total_poles:,}</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-label">ควบคุม IOT ได้</div>
    <div class="kpi-value">{total_iot:,}</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-label">API เชื่อมแล้ว</div>
    <div class="kpi-value">{total_api:,}</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ── Table ─────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">ตารางสรุปแยกตามสัญญา</div>', unsafe_allow_html=True)
render_table(data_df)
