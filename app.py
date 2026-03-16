import streamlit as st
import pandas as pd
from datetime import timedelta
import plotly.express as px
import plotly.graph_objects as go
import numpy as np

st.set_page_config(layout="wide")

# -------------------------------------------------
# NETWORK FILE PATHS (CHANGE THIS)
# -------------------------------------------------

ATTENDANCE_FILE = "data/attendance.xlsx"
ESSENTIAL_FILE = "data/essential_list.xlsx"
# -------------------------------------------------
# STYLE
# -------------------------------------------------

st.markdown("""
<style>
.header-bar{
background-color:#0086E2;
padding:18px;
border-radius:8px;
margin-bottom:20px;
}
.header-text{
color:white;
font-size:28px;
font-weight:bold;
text-align:center;
}
thead tr th{
background-color:#0086E2 !important;
color:white !important;
text-align:center;
}
tbody tr:nth-child(even){
background-color:#F2F2F2;
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="header-bar">
<div class="header-text">OT Monitoring Dashboard</div>
</div>
""", unsafe_allow_html=True)

# -------------------------------------------------
# SIDEBAR
# -------------------------------------------------

st.sidebar.markdown("### Dashboard Controls")

month_option = st.sidebar.selectbox(
"Select Month",
["January","February","March","April","May","June",
"July","August","September","October","November","December"]
)

# -------------------------------------------------
# HOLIDAYS
# -------------------------------------------------

holiday_list = pd.to_datetime([
"14-01-2026","26-01-2026","03-03-2026","21-03-2026",
"01-05-2026","15-08-2026","28-08-2026","02-10-2026",
"20-10-2026","09-11-2026","11-11-2026"
], format="%d-%m-%Y")

# -------------------------------------------------
# ESSENTIAL EMPLOYEES
# -------------------------------------------------

essential_list = []

try:

    essential_df = pd.read_excel(ESSENTIAL_FILE)
    essential_df.columns = essential_df.columns.astype(str).str.strip()

    emp_col = None
    for c in essential_df.columns:
        if "employee" in c.lower() or "name" in c.lower():
            emp_col = c
            break

    if "Month" in essential_df.columns:

        essential_list = essential_df[
            essential_df["Month"] == month_option
        ][emp_col].astype(str).tolist()

    else:

        essential_list = essential_df[emp_col].astype(str).tolist()

except Exception as e:

    st.error("Essential Employee File Not Found")
    st.stop()

# -------------------------------------------------
# HELPERS
# -------------------------------------------------

def normalize_hours(x):

    try:
        x = float(x)
    except:
        return 0

    return round(x)

def is_essential(emp):

    return str(emp) in essential_list

# -------------------------------------------------
# PROCESS EXCEL SHEET
# -------------------------------------------------

def process_sheet(sheet):

    raw = pd.read_excel(ATTENDANCE_FILE, sheet_name=sheet, header=None)

    header_row = None

    for i,row in raw.iterrows():

        row_str = row.astype(str).str.lower()

        if row_str.str.contains("personnel").any():
            header_row = i
            break

    if header_row is None:
        raise Exception("Header row not found")

    df = pd.read_excel(ATTENDANCE_FILE, sheet_name=sheet, header=header_row)

    df.columns = df.columns.astype(str).str.strip()

    personnel_col = None
    name_col = None
    area_col = None

    for col in df.columns:

        col_low = col.lower()

        if "personnel" in col_low:
            personnel_col = col

        if "employee" in col_low or "name" in col_low:
            name_col = col

        if "area" in col_low:
            area_col = col

    if personnel_col is None:
        personnel_col = df.columns[0]

    if name_col is None:
        name_col = df.columns[1]

    df = df.rename(columns={
        personnel_col:"Personnel Number",
        name_col:"Employee/app.name"
    })

    if area_col:
        df = df.rename(columns={area_col:"Area"})
    else:
        df["Area"] = "Unknown"

    df = df.dropna(subset=["Personnel Number"])

    date_columns = []

    for col in df.columns:

        col_str = str(col)

        if "/" in col_str or "-" in col_str:

            try:
                pd.to_datetime(col_str, errors="raise")
                date_columns.append(col)
            except:
                pass

    df_long = df.melt(
        id_vars=["Personnel Number","Employee/app.name","Area"],
        value_vars=date_columns,
        var_name="Date",
        value_name="Daily_Hours"
    )

    df_long["Date"] = pd.to_datetime(df_long["Date"])

    df_long["Daily_Hours"] = pd.to_numeric(
        df_long["Daily_Hours"],
        errors="coerce"
    ).fillna(0)

    df_long["Daily_Hours"] = df_long["Daily_Hours"].apply(normalize_hours)

    df_long["Weekday"] = df_long["Date"].dt.weekday

    df_long["Daily_OT"] = np.where(
        df_long["Daily_Hours"]>9,
        df_long["Daily_Hours"]-8,
        0
    )

    df_long["Week"] = (
        df_long["Date"]
        - pd.to_timedelta((df_long["Date"].dt.weekday + 1) % 7, unit="D")
    ).dt.date

    return df_long

# -------------------------------------------------
# MAIN
# -------------------------------------------------

try:

    xls = pd.ExcelFile(ATTENDANCE_FILE)

except:

    st.error("Attendance file not found in network folder.")
    st.stop()

all_data = []

for sheet in xls.sheet_names:

    try:

        data = process_sheet(sheet)
        all_data.append(data)

    except Exception as e:

        st.warning(f"Skipping sheet {sheet}")

if len(all_data) == 0:

    st.error("No valid sheets found.")
    st.stop()

combined_data = pd.concat(all_data, ignore_index=True)

# -------------------------------------------------
# SUMMARY
# -------------------------------------------------

def create_summary(data):

    weekly = data.groupby(
        ["Personnel Number","Employee/app.name","Area","Week"]
    )["Daily_Hours"].sum().reset_index()

    weekly_max = weekly.groupby(
        ["Personnel Number","Employee/app.name","Area"]
    )["Daily_Hours"].max().reset_index()

    weekly_max.rename(
        columns={"Daily_Hours":"Working hr/week"},
        inplace=True
    )

    ot_total = data.groupby(
        ["Personnel Number","Employee/app.name","Area"]
    )["Daily_OT"].sum().reset_index()

    ot_total.rename(
        columns={"Daily_OT":"Total OT Hours"},
        inplace=True
    )

    # -----------------------------
    # NEW → OT DATE + HOURS
    # -----------------------------

    ot_data = data[data["Daily_OT"] > 0]

    ot_details = ot_data.groupby(
        ["Personnel Number","Employee/app.name","Area"]
    ).apply(
        lambda x: ", ".join(
            f"{d.strftime('%d-%m-%Y')} ({ot}h)"
            for d, ot in zip(x["Date"], x["Daily_OT"])
        )
    ).reset_index(name="OT Details")

    # -----------------------------
    # MERGE DATA
    # -----------------------------

    final = pd.merge(
        weekly_max,
        ot_total,
        on=["Personnel Number","Employee/app.name","Area"]
    )

    final = pd.merge(
        final,
        ot_details,
        on=["Personnel Number","Employee/app.name","Area"],
        how="left"
    )

    final["OT Details"] = final["OT Details"].fillna("No OT")

    final["Continous Working >10 days"] = "No"

    final["OT hrs/qtr >50"] = final["Total OT Hours"].apply(
        lambda x: "Yes" if x > 50 else "No"
    )

    final["Remark"] = final["Working hr/week"].apply(
        lambda x: "Exceeded 60 hr/week" if x > 60 else ""
    )

    final.rename(
        columns={
            "Employee/app.name":"Name",
            "Personnel Number":"E.No"
        },
        inplace=True
    )

    return final

summary = create_summary(combined_data)

st.subheader("Combined Summary")

st.dataframe(summary, use_container_width=True)

st.markdown("---")

st.subheader("Workforce OT Analytics")

area_ot = summary.groupby("Area")["Total OT Hours"].sum().reset_index()

fig_area = px.pie(
    area_ot,
    names="Area",
    values="Total OT Hours"
)

st.plotly_chart(fig_area, use_container_width=True)

top10 = summary.sort_values(
    by="Total OT Hours",
    ascending=False
).head(10)

fig_top = px.bar(
    top10,
    x="Name",
    y="Total OT Hours",
    color="Area"
)

st.plotly_chart(fig_top, use_container_width=True)

daily_ot = combined_data.groupby("Date")["Daily_OT"].sum().reset_index()

fig_trend = px.line(
    daily_ot,
    x="Date",
    y="Daily_OT",
    markers=True
)

st.plotly_chart(fig_trend, use_container_width=True)

ot_nonzero = summary[summary["Total OT Hours"] > 0]

counts, bins = np.histogram(
    ot_nonzero["Total OT Hours"],
    bins=15
)

centers = 0.5 * (bins[:-1] + bins[1:])

fig_hist = go.Figure()

fig_hist.add_bar(
    x=centers,
    y=counts
)

fig_hist.update_layout(
    title="OT Distribution",
    xaxis_title="OT Hours",
    yaxis_title="Employees"
)

st.plotly_chart(fig_hist, use_container_width=True)