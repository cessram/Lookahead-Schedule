"""
Construction Lookahead Dashboard (2 / 3 / 4 / 6 week) with Gantt chart
----------------------------------------------------------------------
Reads the Bird "6wk Schedule" lookahead template (.xlsx / .xlsm), shows a Gantt
chart for the selected lookahead window and exports the same Excel format.

Run:
    pip install streamlit pandas plotly openpyxl
    streamlit run lookahead_dashboard.py
"""

import io
from datetime import date, datetime, timedelta

import openpyxl
import pandas as pd
import plotly.express as px
import streamlit as st
from openpyxl.styles import PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ------------------------------------------------------------------ page setup
st.set_page_config(page_title="Lookahead Schedule Dashboard",
                   page_icon="🏗️", layout="wide")

st.markdown("""
<style>
.block-container {padding-top: 1.5rem;}
.kpi {background:#F7F9FC;border-left:6px solid #F58025;border-radius:8px;
      padding:12px 16px;margin-bottom:8px;}
.kpi .lbl {font-size:0.8rem;color:#5b6573;text-transform:uppercase;letter-spacing:.04em;}
.kpi .val {font-size:1.25rem;font-weight:700;color:#1f2a37;}
</style>
""", unsafe_allow_html=True)

WEEK_OPTIONS = {"2-Week Lookahead": 2, "3-Week Lookahead": 3,
                "4-Week Lookahead": 4, "6-Week Lookahead": 6}
BAR_FILL = PatternFill("solid", fgColor="F58025")   # Bird orange for Excel bars
NO_FILL = PatternFill(fill_type=None)


# ------------------------------------------------------------------ helpers
def find_cell(ws, text, max_row=15):
    """Return (row, col) of the first cell whose value starts with `text`."""
    text = text.upper()
    for row in ws.iter_rows(min_row=1, max_row=max_row):
        for c in row:
            if isinstance(c.value, str) and c.value.strip().upper().startswith(text):
                return c.row, c.column
    return None


def value_right_of(ws, label):
    """Template stores header values (project, super, dates) right of the label (label is merged)."""
    pos = find_cell(ws, label)
    if not pos:
        return None
    r, c = pos
    for cc in range(c + 1, c + 5):
        v = ws.cell(r, cc).value
        if v not in (None, ""):
            return v
    return None


def to_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%d/%m/%Y"):
            try:
                return datetime.strptime(v.strip(), fmt).date()
            except ValueError:
                pass
    return None


def is_marked(v):
    return isinstance(v, str) and v.strip().lower().startswith("x")


@st.cache_data(show_spinner=False)
def parse_workbook(file_bytes: bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
    ws = wb["6wk Schedule"] if "6wk Schedule" in wb.sheetnames else wb.worksheets[0]

    hdr = find_cell(ws, "WORK DESCRIPTION")
    if not hdr:
        raise ValueError("Could not find the 'WORK DESCRIPTION' header. Please use the lookahead template.")
    hdr_row, desc_col = hdr
    id_col = desc_col - 1
    status_pos = find_cell(ws, "STATUS")
    status_col = status_pos[1]
    man_col = status_col + 1
    notes_col = status_col + 2

    # first day column = first column after Duration (or after description)
    dur_pos = find_cell(ws, "DURATION")
    day_start = (dur_pos[1] + 1) if dur_pos else desc_col + 1
    day_cols = list(range(day_start, status_col))

    info = {
        "title": ws.cell(hdr_row - 3, desc_col).value or "",          # e.g. "MMCC 370012"
        "project": value_right_of(ws, "PROJECT NAME") or "",
        "number": str(value_right_of(ws, "PROJECT NUMBER") or ""),
        "super": value_right_of(ws, "SUPERINTENDENT") or "",
        "from": to_date(value_right_of(ws, "FROM")),
        "to": to_date(value_right_of(ws, "TO:")),
        "sheet": ws.title,
        "hdr_row": hdr_row, "desc_col": desc_col, "id_col": id_col,
        "day_cols": day_cols, "status_col": status_col,
        "man_col": man_col, "notes_col": notes_col,
    }
    if not info["from"]:
        raise ValueError("The 'FROM:' date is missing or not a valid date.")

    # read activities; rows with a description but no marks become section headers (trade/area)
    rows, section = [], "General"
    first_data = hdr_row + 2          # header is 2 rows tall (letters + day numbers)
    for r in range(first_data, ws.max_row + 1):
        desc = ws.cell(r, desc_col).value
        if desc is None or str(desc).strip() in ("", "0"):
            continue
        desc = str(desc).strip()
        offsets = [i for i, c in enumerate(day_cols) if is_marked(ws.cell(r, c).value)]
        if not offsets:
            section = desc
            continue
        rows.append({
            "row": r, "Section": section, "Activity": desc,
            "Sched ID": ws.cell(r, id_col).value or "",
            "offsets": offsets,
            "Status / % Complete": ws.cell(r, status_col).value or "",
            "Manpower": ws.cell(r, man_col).value or "",
            "Notes": ws.cell(r, notes_col).value or "",
        })
    return info, rows


def build_gantt_df(info, rows, n_days):
    """Split each activity into continuous bars within the selected window."""
    start = info["from"]
    recs = []
    for a in rows:
        offs = [o for o in a["offsets"] if o < n_days]
        if not offs:
            continue
        seg = [offs[0]]
        for o in offs[1:] + [None]:
            if o is not None and o == seg[-1] + 1:
                seg.append(o)
                continue
            recs.append({
                "Section": a["Section"],
                "Activity": a["Activity"],
                "Start": start + timedelta(days=seg[0]),
                "Finish": start + timedelta(days=seg[-1] + 1),
                "Days": len(seg),
                "Status": a["Status / % Complete"],
                "Manpower": a["Manpower"],
            })
            if o is not None:
                seg = [o]
    return pd.DataFrame(recs)


def export_excel(file_bytes, info, rows, n_weeks, is_xlsm):
    """Return the uploaded template trimmed to the selected window (same format)."""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), keep_vba=is_xlsm)
    ws = wb[info["sheet"]]
    n_days = min(n_weeks * 7, len(info["day_cols"]))
    start = info["from"]
    hdr_row = info["hdr_row"]

    # header dates
    for lbl, val in (("TO:", start + timedelta(days=n_days - 1)),):
        r, c = find_cell(ws, lbl)
        for cc in range(c + 1, c + 5):
            if ws.cell(r, cc).value not in (None, ""):
                ws.cell(r, cc).value = val
                ws.cell(r, cc).number_format = "mm/dd/yyyy"
                break

    # weekday letters + real day numbers, colour bars, hide columns outside window
    for i, col in enumerate(info["day_cols"]):
        d = start + timedelta(days=i)
        letter = get_column_letter(col)
        in_window = i < n_days
        ws.column_dimensions[letter].hidden = not in_window
        ws.cell(hdr_row, col).value = "MTWTFSS"[d.weekday()]
        ws.cell(hdr_row + 1, col).value = d.day
        ws.cell(hdr_row + 1, col).number_format = "0"
        for a in rows:
            cell = ws.cell(a["row"], col)
            if in_window and i in a["offsets"]:
                cell.fill = BAR_FILL
                cell.alignment = Alignment(horizontal="center")
            elif not in_window and is_marked(cell.value):
                cell.value = None          # remove marks beyond the selected window
                cell.fill = NO_FILL

    # hide week banners (PLANNED THIS WEEK ... 6 WEEK FORECASTED) beyond the window
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def kpi(col, label, value):
    col.markdown(f'<div class="kpi"><div class="lbl">{label}</div>'
                 f'<div class="val">{value}</div></div>', unsafe_allow_html=True)


# ------------------------------------------------------------------ sidebar
with st.sidebar:
    st.header("📂 1. Upload Lookahead")
    up = st.file_uploader("Excel lookahead (6wk Schedule template)",
                          type=["xlsx", "xlsm"])
    st.header("📅 2. Select Lookahead")
    choice = st.radio("Window", list(WEEK_OPTIONS), index=3)
    n_weeks = WEEK_OPTIONS[choice]
    st.header("🎨 3. Display")
    color_by = st.selectbox("Colour bars by", ["Section", "Activity"])
    st.caption("Tip: rows with a description but no 'x' marks (e.g. *R3 demo*, "
               "*I-Mark*, *Bird scopes*) are treated as trade/area section headers.")

# ------------------------------------------------------------------ main
if not up:
    st.title("🏗️ Construction Lookahead Dashboard")
    st.info("Upload your lookahead Excel file in the sidebar to generate the Gantt chart.")
    st.markdown("""
**How it works**
1. Upload the *6wk Schedule* Excel file (same template as MMCC_Lookahead.xlsx).
2. Pick a 2, 3, 4 or 6-week lookahead.
3. Review the Gantt chart and activity table.
4. Download the lookahead in the same Excel format.
""")
    st.stop()

file_bytes = up.getvalue()
try:
    info, rows = parse_workbook(file_bytes)
except Exception as e:
    st.error(f"Could not read the file: {e}")
    st.stop()

n_days = min(n_weeks * 7, len(info["day_cols"]))
win_start = info["from"]
win_end = win_start + timedelta(days=n_days - 1)

title = info["title"] or f'{info["project"]} {info["number"]}'
st.title(f"🏗️ {title} — {choice}")

c1, c2, c3, c4 = st.columns(4)
kpi(c1, "Project Name", f'{info["project"]} ({info["number"]})')
kpi(c2, "Superintendent", info["super"] or "—")
kpi(c3, "Planned From", win_start.strftime("%b %d, %Y"))
kpi(c4, "Lookahead To", win_end.strftime("%b %d, %Y"))
if n_weeks * 7 > len(info["day_cols"]):
    st.caption(f"Note: the template contains {len(info['day_cols'])} day columns; "
               f"the window ends on the last day in the file.")

gdf = build_gantt_df(info, rows, n_days)

m1, m2, m3 = st.columns(3)
m1.metric("Activities in window", gdf["Activity"].nunique() if not gdf.empty else 0)
m2.metric("Trade / area sections", gdf["Section"].nunique() if not gdf.empty else 0)
m3.metric("Total planned activity-days", int(gdf["Days"].sum()) if not gdf.empty else 0)

tab1, tab2 = st.tabs(["📊 Gantt Chart", "📋 Activity List"])

with tab1:
    if gdf.empty:
        st.warning("No activities are planned within this lookahead window.")
    else:
        order = list(dict.fromkeys(gdf["Activity"]))
        fig = px.timeline(gdf, x_start="Start", x_end="Finish", y="Activity",
                          color=color_by, hover_data=["Section", "Days", "Status", "Manpower"],
                          category_orders={"Activity": order},
                          color_discrete_sequence=px.colors.qualitative.Set2)
        fig.update_yaxes(autorange="reversed", title="")
        fig.update_xaxes(range=[win_start, win_end + timedelta(days=1)],
                         dtick=86400000, tickformat="%a<br>%b %d", side="top",
                         showgrid=True, gridcolor="#e5e7eb")
        # shade weekends and mark week breaks
        for i in range(n_days):
            d = win_start + timedelta(days=i)
            if d.weekday() >= 5:
                fig.add_vrect(x0=d, x1=d + timedelta(days=1), fillcolor="#9ca3af",
                              opacity=0.12, line_width=0)
        for w in range(1, n_weeks):
            fig.add_vline(x=win_start + timedelta(days=7 * w), line_dash="dot",
                          line_color="#F58025")
        today = date.today()
        if win_start <= today <= win_end:
            fig.add_vline(x=today, line_color="red", line_width=2)
        fig.update_layout(height=max(380, 42 * len(order) + 140),
                          margin=dict(l=10, r=10, t=60, b=10),
                          legend=dict(orientation="h", y=-0.08, title=""))
        st.plotly_chart(fig, use_container_width=True)
        st.caption("Grey = weekends · Orange dotted = week breaks · Red = today")

with tab2:
    table = []
    for a in rows:
        offs = [o for o in a["offsets"] if o < n_days]
        if not offs:
            continue
        table.append({
            "Section": a["Section"], "Sched ID": a["Sched ID"], "Activity": a["Activity"],
            "Start": win_start + timedelta(days=offs[0]),
            "Finish": win_start + timedelta(days=offs[-1]),
            "Duration (days)": len(offs),
            "Status / % Complete": a["Status / % Complete"],
            "Manpower": a["Manpower"], "Notes": a["Notes"],
        })
    tdf = pd.DataFrame(table)
    if not tdf.empty:
        sec = st.multiselect("Filter by section", sorted(tdf["Section"].unique()))
        if sec:
            tdf = tdf[tdf["Section"].isin(sec)]
    st.dataframe(tdf, use_container_width=True, hide_index=True)

# ------------------------------------------------------------------ download
st.divider()
is_xlsm = up.name.lower().endswith(".xlsm")
ext = "xlsm" if is_xlsm else "xlsx"
out_bytes = export_excel(file_bytes, info, rows, n_weeks, is_xlsm)
fname = f'{info["project"] or "Project"}_{n_weeks}wk_Lookahead_{win_start:%Y%m%d}.{ext}'
st.download_button(f"⬇️ Download {choice} (Excel)", data=out_bytes, file_name=fname,
                   mime="application/vnd.ms-excel.sheet.macroEnabled.12" if is_xlsm else
                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                   type="primary", use_container_width=True)
st.caption("The download keeps the original template layout. Days outside the selected "
           "window are hidden, planned days are shaded orange, and the TO: date is updated.")
