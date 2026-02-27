"""
Harvest Audit Export — Streamlit App
Commit Consulting | Financial Operations

OAuth 2.0 login via Harvest. Requires client_id / client_secret in
.streamlit/secrets.toml (local) or Streamlit Cloud secrets dashboard.
"""

import io
import sys
import os
import urllib.parse
import requests
import streamlit as st
from datetime import date, timedelta
from openpyxl import Workbook

# Make src/ importable
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

from harvest_export import (
    fetch_time_entries,
    fetch_timesheet_approvals,
    parse_entries,
    add_audit_columns,
    build_summary,
    detect_duplicates,
    write_summary_sheet,
    write_duplicates_sheet,
    write_blank_notes_sheet,
    write_raw_sheet,
)

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="Harvest Audit Export | Commit Consulting",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------------
# OAuth constants
# ---------------------------------------------------------------------------

try:
    CLIENT_ID     = st.secrets["harvest"]["client_id"]
    CLIENT_SECRET = st.secrets["harvest"]["client_secret"]
    REDIRECT_URI  = st.secrets["harvest"]["redirect_uri"]
except KeyError:
    st.error(
        "⚠️ App secrets are not configured. "
        "Please add your Harvest OAuth credentials in the Streamlit Cloud secrets dashboard.\n\n"
        "See DEPLOY.md for instructions."
    )
    st.stop()

HARVEST_AUTH_URL     = "https://id.getharvest.com/oauth2/authorize"
HARVEST_TOKEN_URL    = "https://id.getharvest.com/api/v2/oauth2/token"
HARVEST_ACCOUNTS_URL = "https://id.getharvest.com/api/v1/accounts"


# ---------------------------------------------------------------------------
# OAuth helpers
# ---------------------------------------------------------------------------

def get_auth_url(state: str) -> str:
    params = {
        "client_id":     CLIENT_ID,
        "redirect_uri":  REDIRECT_URI,
        "scope":         "all",
        "response_type": "code",
        "state":         state,
    }
    return f"{HARVEST_AUTH_URL}?{urllib.parse.urlencode(params)}"


def exchange_code_for_token(code: str) -> dict:
    resp = requests.post(HARVEST_TOKEN_URL, data={
        "code":          code,
        "client_id":     CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "redirect_uri":  REDIRECT_URI,
        "grant_type":    "authorization_code",
    }, timeout=15)
    resp.raise_for_status()
    return resp.json()


def get_harvest_accounts(access_token: str) -> dict:
    """Returns {"user": {...}, "accounts": [...]}"""
    resp = requests.get(HARVEST_ACCOUNTS_URL, headers={
        "Authorization": f"Bearer {access_token}",
        "User-Agent":    "CommitConsulting-HarvestAudit/1.0",
    }, timeout=15)
    resp.raise_for_status()
    return resp.json()


# ---------------------------------------------------------------------------
# Excel generation
# ---------------------------------------------------------------------------

def generate_excel_bytes(df, summary, dupes, from_date: str, to_date: str) -> bytes:
    wb = Workbook()
    del wb[wb.sheetnames[0]]
    write_summary_sheet(wb, summary, from_date, to_date)
    write_duplicates_sheet(wb, dupes)
    write_blank_notes_sheet(wb, df)
    write_raw_sheet(wb, df)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Session state init
# ---------------------------------------------------------------------------

defaults = {
    "oauth_state":    None,
    "access_token":   None,
    "account_id":     None,
    "accounts":       [],
    "user":           None,
    "df":             None,
    "summary":        None,
    "dupes":          None,
    "excel_bytes":    None,
    "report_dates":   None,
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v


# ---------------------------------------------------------------------------
# OAuth callback handler  (runs before any UI is rendered)
# ---------------------------------------------------------------------------

params = st.query_params
if "code" in params and "state" in params:
    if params["state"] == st.session_state.oauth_state:
        try:
            token_data    = exchange_code_for_token(params["code"])
            access_token  = token_data["access_token"]
            accounts_data = get_harvest_accounts(access_token)

            st.session_state.access_token = access_token
            st.session_state.user         = accounts_data.get("user", {})
            st.session_state.accounts     = accounts_data.get("accounts", [])

            # Auto-select if only one account
            if len(st.session_state.accounts) == 1:
                st.session_state.account_id = str(st.session_state.accounts[0]["id"])

            st.query_params.clear()
            st.rerun()

        except Exception as e:
            st.error(f"OAuth error: {e}")
            st.query_params.clear()
    else:
        st.warning("OAuth state mismatch — please try logging in again.")
        st.query_params.clear()


# ---------------------------------------------------------------------------
# Custom CSS
# ---------------------------------------------------------------------------

st.markdown("""
<style>
    [data-testid="stSidebar"] { background-color: #1F3864; }
    [data-testid="stSidebar"] * { color: #FFFFFF !important; }
    [data-testid="stSidebar"] .stSelectbox label,
    [data-testid="stSidebar"] .stDateInput label { color: #D6E4F0 !important; }
    div[data-testid="metric-container"] {
        background: #EBF3FB;
        border: 1px solid #D6E4F0;
        border-radius: 8px;
        padding: 12px 16px;
    }
    .commit-header {
        background: #1F3864;
        color: white;
        padding: 18px 24px;
        border-radius: 8px;
        margin-bottom: 20px;
    }
    .stTabs [data-baseweb="tab"] { font-size: 14px; font-weight: 500; }
</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Login page
# ---------------------------------------------------------------------------

def show_login():
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("""
        <div style="text-align:center; padding: 60px 0 30px 0;">
            <div style="font-size:48px;">📊</div>
            <h1 style="color:#1F3864; margin:8px 0 4px 0;">Harvest Audit Export</h1>
            <p style="color:#595959; font-size:15px;">Commit Consulting — Financial Operations</p>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("""
        <div style="background:#EBF3FB; border:1px solid #D6E4F0; border-radius:8px;
                    padding:20px 24px; margin-bottom:24px; font-size:14px; color:#333;">
            Pull time entry data from Harvest and generate a formatted audit report
            with duplicate detection, late submission analysis, and notes review.
        </div>
        """, unsafe_allow_html=True)

        import secrets as _secrets
        if not st.session_state.oauth_state:
            st.session_state.oauth_state = _secrets.token_urlsafe(16)
        state = st.session_state.oauth_state

        auth_url = get_auth_url(state)
        st.link_button("🌾  Login with Harvest", auth_url, use_container_width=True, type="primary")

        st.caption("You will be redirected to Harvest to authorise access. No passwords are stored by this application.")


# ---------------------------------------------------------------------------
# Account selector (for users with multiple Harvest accounts)
# ---------------------------------------------------------------------------

def show_account_selector():
    st.info("You have access to multiple Harvest accounts. Please select one to continue.")
    options = {a["name"]: str(a["id"]) for a in st.session_state.accounts}
    choice  = st.selectbox("Select account", list(options.keys()))
    if st.button("Continue", type="primary"):
        st.session_state.account_id = options[choice]
        st.rerun()


# ---------------------------------------------------------------------------
# Main app
# ---------------------------------------------------------------------------

def show_app():
    user       = st.session_state.user or {}
    first_name = user.get("first_name", "")
    last_name  = user.get("last_name", "")
    full_name  = f"{first_name} {last_name}".strip() or "Unknown"

    account_name = next(
        (a["name"] for a in st.session_state.accounts
         if str(a["id"]) == st.session_state.account_id),
        "Unknown",
    )

    # ── Sidebar ──────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown(f"### 📊 Harvest Audit Export")
        st.markdown(f"**{account_name}**")
        st.markdown(f"Logged in as **{full_name}**")
        st.divider()

        st.markdown("#### Date Range")
        default_start = date.today().replace(day=1)
        default_end   = date.today()

        from_date = st.date_input("Start date", value=default_start)
        to_date   = st.date_input("End date",   value=default_end)

        if from_date > to_date:
            st.error("End date must be after start date.")

        st.divider()
        run_clicked = st.button("▶ Run Report", type="primary", use_container_width=True)

        if st.session_state.excel_bytes and st.session_state.report_dates:
            fd, td = st.session_state.report_dates
            st.download_button(
                label="⬇ Download Excel Report",
                data=st.session_state.excel_bytes,
                file_name=f"Harvest_Audit_{fd}_to_{td}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        st.divider()
        if st.button("Logout", use_container_width=True):
            for k in defaults:
                st.session_state[k] = defaults[k]
            st.rerun()

    # ── Header ────────────────────────────────────────────────────────────
    st.markdown("""
    <div class="commit-header">
        <span style="font-size:22px; font-weight:700;">Harvest Audit Export</span>
        <span style="font-size:14px; opacity:0.7; margin-left:16px;">Commit Consulting — Financial Operations</span>
    </div>
    """, unsafe_allow_html=True)

    # ── Run report ────────────────────────────────────────────────────────
    if run_clicked and from_date <= to_date:
        from_str = from_date.strftime("%Y-%m-%d")
        to_str   = to_date.strftime("%Y-%m-%d")

        with st.status("Fetching data from Harvest...", expanded=True) as status:
            st.write(f"Fetching time entries ({from_str} to {to_str})...")
            try:
                entries = fetch_time_entries(
                    st.session_state.access_token,
                    st.session_state.account_id,
                    from_str, to_str,
                )
            except requests.HTTPError as e:
                if e.response.status_code == 401:
                    status.update(label="Session expired — please log in again.", state="error")
                    st.session_state.access_token = None
                    st.rerun()
                raise

            if not entries:
                status.update(label="No time entries found for that date range.", state="error")
                st.stop()

            st.write("Fetching timesheet approvals...")
            approvals = fetch_timesheet_approvals(
                st.session_state.access_token,
                st.session_state.account_id,
                from_str, to_str,
            )

            st.write("Processing data and computing audit metrics...")
            df      = parse_entries(entries, approvals)
            df      = add_audit_columns(df)
            summary = build_summary(df)
            dupes   = detect_duplicates(df)

            st.write("Generating Excel report...")
            excel_bytes = generate_excel_bytes(df, summary, dupes, from_str, to_str)

            st.session_state.df          = df
            st.session_state.summary     = summary
            st.session_state.dupes       = dupes
            st.session_state.excel_bytes = excel_bytes
            st.session_state.report_dates = (from_str, to_str)

            status.update(
                label=f"Report ready — {len(entries)} entries loaded.",
                state="complete",
                expanded=False,
            )
        st.rerun()

    # ── Report display ────────────────────────────────────────────────────
    if st.session_state.df is None:
        st.markdown("""
        <div style="text-align:center; padding:80px 0; color:#999;">
            <div style="font-size:40px;">📋</div>
            <p style="font-size:16px;">Select a date range and click <strong>Run Report</strong> to get started.</p>
        </div>
        """, unsafe_allow_html=True)
        return

    df      = st.session_state.df
    summary = st.session_state.summary
    dupes   = st.session_state.dupes
    fd, td  = st.session_state.report_dates

    st.caption(f"Report period: **{fd}** to **{td}**  ·  {len(df):,} entries")

    # ── KPI metrics ───────────────────────────────────────────────────────
    kpis     = summary["kpis"]
    kpi_keys = list(kpis.items())

    cols = st.columns(4)
    for i, (label, value) in enumerate(kpi_keys[:4]):
        cols[i].metric(label, value)

    cols = st.columns(4)
    for i, (label, value) in enumerate(kpi_keys[4:8]):
        cols[i].metric(label, value)

    st.divider()

    # ── Tabs ──────────────────────────────────────────────────────────────
    tab_summary, tab_dupes, tab_blank, tab_raw = st.tabs([
        "📊 Audit Summary",
        f"🔁 Duplicate Entries ({dupes['Duplicate Group'].nunique() if not dupes.empty else 0} groups)",
        f"📝 Blank Notes ({int(df['Blank Notes'].sum())})",
        "📄 Raw Data",
    ])

    # ── Audit Summary tab ─────────────────────────────────────────────────
    with tab_summary:
        st.subheader("Breakdown by Employee")
        st.dataframe(
            summary["by_employee"],
            use_container_width=True,
            hide_index=True,
        )

        st.subheader("Breakdown by Client / Project")
        st.dataframe(
            summary["by_project"],
            use_container_width=True,
            hide_index=True,
        )

        if not summary["flags"].empty:
            st.subheader(f"⚑ Audit Flags — {len(summary['flags'])} entries")
            st.caption("Late submissions and entries edited after being locked.")
            st.dataframe(
                summary["flags"],
                use_container_width=True,
                hide_index=True,
            )
        else:
            st.success("No audit flags for this period.")

    # ── Duplicate Entries tab ─────────────────────────────────────────────
    with tab_dupes:
        if dupes.empty:
            st.success("No potential duplicate entries found for this period.")
        else:
            hours_groups = dupes[dupes["Match Type"].str.startswith("Hours")]["Duplicate Group"].nunique()
            notes_groups = dupes[dupes["Match Type"].str.startswith("Notes")]["Duplicate Group"].nunique()

            c1, c2 = st.columns(2)
            c1.metric("Hours match groups", hours_groups,
                      help="Same employee, client, project, date & hours")
            c2.metric("Notes match groups", notes_groups,
                      help="Same employee, client, project, date & notes")

            st.dataframe(dupes, use_container_width=True, hide_index=True)

    # ── Blank Notes tab ───────────────────────────────────────────────────
    with tab_blank:
        blank_df = df[df["Blank Notes"] == True]
        if blank_df.empty:
            st.success("All entries have notes for this period.")
        else:
            display_cols = [
                "Entry ID", "Employee Name", "Client", "Project", "Task",
                "Work Date", "Hours", "Billable", "Is Locked",
                "Late Submission", "Created At",
            ]
            keep = [c for c in display_cols if c in blank_df.columns]
            st.dataframe(
                blank_df[keep].sort_values(["Employee Name", "Work Date"]),
                use_container_width=True,
                hide_index=True,
            )

    # ── Raw Data tab ──────────────────────────────────────────────────────
    with tab_raw:
        st.caption(f"{len(df):,} total entries. Use the column headers to sort and filter.")
        st.dataframe(df, use_container_width=True, hide_index=True)


# ---------------------------------------------------------------------------
# Routing
# ---------------------------------------------------------------------------

if not st.session_state.access_token:
    show_login()
elif not st.session_state.account_id:
    show_account_selector()
else:
    show_app()
