### OT Dashboard Code
import streamlit as st
import pandas as pd
from datetime import timedelta
import plotly.express as px
import plotly.graph_objects as go
import numpy as np

st.set_page_config(layout="wide")
st.markdown("""
<style>

/* Remove extra vertical spacing between blocks */
div[data-testid="stVerticalBlock"]{
    gap:0.5rem;
}

/* Remove space inside containers */
div[data-testid="stVerticalBlock"] > div{
    padding-top:0px;
    padding-bottom:0px;
}

/* Reduce heading spacing */
h1, h2, h3, h4 {
    margin-top: 0px !important;
    margin-bottom: 2px !important;
}

/* Reduce button spacing */
div.stButton {
    margin-bottom:-5px !important;
}

</style>
""", unsafe_allow_html=True)

# -------------------------------------------------
# NETWORK FILE PATHS (CHANGE THIS)
# -------------------------------------------------

ATTENDANCE_FILE = "data/attendance.xlsx"
ESSENTIAL_FILE = "data/essential_list.xlsx"
# --------------------------------------
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
# SESSION STATE INITIALIZATION
# -------------------------------------------------

if "selected_months" not in st.session_state:
    st.session_state.selected_months = []

# -------------------------------------------------
# -------------------------------------------------
# MONTH TILE SELECTOR
# -------------------------------------------------



months = [
    "Jan","Feb","Mar","Apr","May","Jun",
    "Jul","Aug","Sep","Oct","Nov","Dec"
]

# session state
selected_months = st.session_state.selected_months

    # color logic
   # color = "#198754" if m in st.session_state.selected_months else "#E9ECEF"

  

   
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

    if "Month" in essential_df.columns and selected_months:
        essential_list = essential_df[
        essential_df["Month"].isin(selected_months)
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
    
    

    df_long["Month"] = sheet
    
    
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

if selected_months:

    month_map = {
        "Jan":1,"Feb":2,"Mar":3,"Apr":4,"May":5,"Jun":6,
        "Jul":7,"Aug":8,"Sep":9,"Oct":10,"Nov":11,"Dec":12
    }

    selected_month_numbers = [month_map[m] for m in selected_months]

    combined_data = combined_data[
        combined_data["Date"].dt.month.isin(selected_month_numbers)
    ]
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

# -------------------------------------------------

# -------------------------------------------------
# -------------------------------------------------
# POLICY KPI CARDS (PROPER LAYOUT)
# -------------------------------------------------
# -------------------------------------------------
# POLICY KPI CARDS (CLICKABLE WITHOUT VIEW BUTTON)
# -------------------------------------------------

week_violation = summary[summary["Working hr/week"] > 60]
ot_violation = summary[summary["Total OT Hours"] > 50]
cont_violation = summary[summary["Continous Working >10 days"] == "Yes"]

#st.markdown("## 🚨 Policy Monitoring")

# Invisible button style
st.markdown("""
<style>
div.stButton > button {
    width:100%;
    height:50px;
    border:none;
    background:transparent;
}
</style>
""", unsafe_allow_html=True)

c1, c2, c3 = st.columns(3)

# -------------------------------------------------
# TILE 1
# -------------------------------------------------

with c1:
    st.markdown(f"""
    <div style="
        background:#ffe6e6;
        padding:25px;
        border-radius:12px;
        text-align:center;
        border:2px solid red;
        cursor:pointer;
    ">
        <h2>⏱</h2>
        <h4>Working hrs</h4>
        <h4>>60 / Week</h4>
        <h2 style="color:red;">{len(week_violation)}</h2>
        <p>Employees</p>
    </div>
    """, unsafe_allow_html=True)

    week_btn = st.button("", key="week_tile")

# -------------------------------------------------
# TILE 2
# -------------------------------------------------

with c2:
    st.markdown(f"""
    <div style="
        background:#fff3e0;
        padding:25px;
        border-radius:12px;
        text-align:center;
        border:2px solid orange;
        cursor:pointer;
    ">
        <h2>📈</h2>
        <h4>OT hrs</h4>
        <h4>>50 / Quarter</h4>
        <h2 style="color:orange;">{len(ot_violation)}</h2>
        <p>Employees</p>
    </div>
    """, unsafe_allow_html=True)

    ot_btn = st.button("", key="ot_tile")

# -------------------------------------------------
# TILE 3
# -------------------------------------------------

with c3:
    st.markdown(f"""
    <div style="
        background:#f3e5f5;
        padding:25px;
        border-radius:12px;
        text-align:center;
        border:2px solid purple;
        cursor:pointer;
    ">
        <h2>🔁</h2>
        <h4>Continuous Punch</h4>
        <h4>>10 Days</h4>
        <h2 style="color:purple;">{len(cont_violation)}</h2>
        <p>Employees</p>
    </div>
    """, unsafe_allow_html=True)

    cont_btn = st.button("", key="cont_tile")


# -------------------------------------------------
# SHOW DETAILS BASED ON TILE CLICK
# -------------------------------------------------

# ⏱ Working Hours >60
if week_btn:

    st.markdown("### ⏱ Employees Working >60 hrs/week")

    if len(week_violation) > 0:

        for _, row in week_violation.iterrows():

            st.markdown(f"""
            <div style="
                border:3px solid red;
                border-radius:10px;
                padding:15px;
                margin-bottom:10px;
                background:#ffe6e6;
            ">
            <b>{row['Name']}</b><br>
            E.No : {row['E.No']}<br>
            Area : {row['Area']}<br>
            Working hr/week : {row['Working hr/week']}<br>
            OT Details : {row['OT Details']}
            </div>
            """, unsafe_allow_html=True)

    else:
        st.success("No employees exceeding 60 hrs/week")


# 📈 OT >50
if ot_btn:

    st.markdown("### 📈 Employees OT >50 hrs / Quarter")

    if len(ot_violation) > 0:

        for _, row in ot_violation.iterrows():

            st.markdown(f"""
            <div style="
                border:3px solid orange;
                border-radius:10px;
                padding:15px;
                margin-bottom:10px;
                background:#fff3e0;
            ">
            <b>{row['Name']}</b><br>
            E.No : {row['E.No']}<br>
            Area : {row['Area']}<br>
            Total OT Hours : {row['Total OT Hours']}<br>
            OT Details : {row['OT Details']}
            </div>
            """, unsafe_allow_html=True)

    else:
        st.success("No employees exceeding 50 OT hours")


# 🔁 Continuous Punch
if cont_btn:

    st.markdown("### 🔁 Continuous Punch >10 Days")

    if len(cont_violation) > 0:

        for _, row in cont_violation.iterrows():

            st.markdown(f"""
            <div style="
                border:3px solid purple;
                border-radius:10px;
                padding:15px;
                margin-bottom:10px;
                background:#f3e5f5;
            ">
            <b>{row['Name']}</b><br>
            E.No : {row['E.No']}<br>
            Area : {row['Area']}<br>
            Status : Continuous Working >10 days
            </div>
            """, unsafe_allow_html=True)

    else:
        st.success("No continuous punch violations")

# ---------------------------------------
# DEPT WISE DEVIATION COUNT
# ---------------------------------------

dept_deviation = (
    summary[summary["Remark"] == "Exceeded 60 hr/week"]
    .groupby("Area")
    .size()
    .reset_index(name="Deviation Count")
)

dept_text = ""

for _, row in dept_deviation.iterrows():
    dept_text += f"{row['Area']} : {row['Deviation Count']} no's  |  "


# -------------------------------------------------
# DEVIATION MONITORING
# -------------------------------------------------

st.markdown(
    "<h3 style='margin-top:5px;margin-bottom:5px;'>🚨 Deviation Monitoring</h3>",
    unsafe_allow_html=True
)

st.markdown("### Department Deviation")

cols = st.columns(6)

for i, row in dept_deviation.iterrows():

    dept = row["Area"]
    count = row["Deviation Count"]

    tile = f"""
    <div style="
        background-color:#F5F5F5;
        padding:15px;
        border-radius:10px;
        text-align:center;
        border:1px solid #DDD;
        margin-bottom:10px;
    ">
        <h4 style="margin:0;">{dept}</h4>
        <h2 style="margin:0;color:#E53935;">{count}</h2>
        <span style="font-size:12px;">Deviations</span>
    </div>
    """

    cols[i % 6].markdown(tile, unsafe_allow_html=True)

violated = summary[summary["Working hr/week"] > 60]

risk = summary[
    (summary["Working hr/week"] >= 55) &
    (summary["Working hr/week"] <= 60)
]

# KPI NUMBERS

col1, col2 = st.columns(2)

with col1:
    st.error(f"🔴 Violations : {len(violated)}")

with col2:
    st.warning(f"🟠 Risk : {len(risk)}")

# -------------------------------------------------
# VIOLATION DETAILS (EXPANDABLE)
# -------------------------------------------------

with st.expander("🔴 View Policy Violations"):

    if len(violated) > 0:

        for _, row in violated.iterrows():

            st.markdown(f"""
            <div style="
                border:3px solid red;
                border-radius:10px;
                padding:15px;
                margin-bottom:12px;
                background-color:#ffe6e6;
            ">
            <b>👤 {row['Name']}</b><br>
            <b>E.No :</b> {row['E.No']} <br>
            <b>Area :</b> {row['Area']} <br>
            <b>Working hr/week :</b> {row['Working hr/week']} <br><br>

            <b style="color:red;">Violation :</b> {row['Remark']}<br>
            <b>OT Dates :</b> {row['OT Details']}
            </div>
            """, unsafe_allow_html=True)

    else:

        st.success("No policy violations detected")

# -------------------------------------------------
# RISK DETAILS (EXPANDABLE)
# -------------------------------------------------

with st.expander("🟠 View Employees Close to Violation"):

    if len(risk) > 0:

        for _, row in risk.iterrows():

            st.markdown(f"""
            <div style="
                border:3px solid orange;
                border-radius:10px;
                padding:15px;
                margin-bottom:12px;
                background-color:#fff3e0;
            ">
            <b>👤 {row['Name']}</b><br>
            <b>E.No :</b> {row['E.No']} <br>
            <b>Area :</b> {row['Area']} <br>
            <b>Working hr/week :</b> {row['Working hr/week']} <br><br>

            <b style="color:orange;">Warning :</b> Close to 60 hr/week limit<br>
            <b>OT Dates :</b> {row['OT Details']}
            </div>
            """, unsafe_allow_html=True)

    else:

        st.info("No employees close to violation")







# -------------------------------------------------
# DEPTWISE DEVIATION COUNT
# -------------------------------------------------



# -------------------------------------------------
# -------------------------------------------------
# -------------------------------------------------
# DEPT WISE OT TREND (% SHARE)
# -------------------------------------------------

st.subheader("📊 Dept wise OT Trend (%)")

# -------------------------------------------------
# DEPT WISE OT TREND (%)
# -------------------------------------------------



#st.markdown("## 📊 Dept Wise OT Trend (%)")

# Create bar chart
# -------------------------------------------------
# DEPT WISE OT TREND (%)
# -------------------------------------------------

# -------------------------------------------------
# DEPT WISE OT TREND (%)
# -------------------------------------------------

# -------------------------------------------------
# DEPT WISE OT TREND (%)
# -------------------------------------------------

# -------------------------------------------------
# DEPT WISE OT TREND (%)
# -------------------------------------------------

# -------------------------------------------------
# DEPT WISE OT TREND (%)
# -------------------------------------------------

# Clean month values
combined_data["Month"] = combined_data["Month"].astype(str).str[:3]

month_order = [
    "Jan","Feb","Mar","Apr","May","Jun",
    "Jul","Aug","Sep","Oct","Nov","Dec"
]

# Department OT calculation
dept_ot = combined_data.groupby(["Month","Area"]).agg(
    total_emp=("Personnel Number", "nunique"),
    ot_emp=("Daily_OT", lambda x: (x > 0).sum())
).reset_index()

# Calculate percentage
dept_ot["OT_Percent"] = (dept_ot["ot_emp"] / dept_ot["total_emp"]) * 100
dept_ot["OT_Percent"] = dept_ot["OT_Percent"].round(1)

# Set month order
dept_ot["Month"] = pd.Categorical(
    dept_ot["Month"],
    categories=month_order,
    ordered=True
)

dept_ot = dept_ot.sort_values("Month")

# Plot chart
fig_ot = px.bar(
    dept_ot,
    x="Area",
    y="OT_Percent",
    color="Month",
    barmode="group",
    text="OT_Percent",
    category_orders={"Month": month_order}
)

fig_ot.update_traces(texttemplate='%{text}%', textposition='outside')

fig_ot.update_layout(
    xaxis_title="Department",
    yaxis_title="OT %",
    yaxis=dict(range=[0,100]),
    height=420
)

st.plotly_chart(fig_ot, use_container_width=True)

# -------------------------------------------------
# DEPT OT HEATMAP
# -------------------------------------------------
# -------------------------------------------------
# DEPT OT HEATMAP
# -------------------------------------------------

st.subheader("🔥 Dept OT Heatmap (%)")

# Calculate department OT %
dept_ot_heat = combined_data.groupby(["Month","Area"]).agg(
    total_emp=("Personnel Number","nunique"),
    ot_emp=("Daily_OT", lambda x: (x > 0).sum())
).reset_index()

dept_ot_heat["OT_Percent"] = (
    dept_ot_heat["ot_emp"] / dept_ot_heat["total_emp"]
) * 100

dept_ot_heat["OT_Percent"] = dept_ot_heat["OT_Percent"].round(1)

# Pivot for heatmap
heatmap_data = dept_ot_heat.pivot(
    index="Area",
    columns="Month",
    values="OT_Percent"
)

# Fill missing values
heatmap_data = heatmap_data.fillna(0)

# Create heatmap
fig_heat = px.imshow(
    heatmap_data,
    text_auto=True,
    aspect="auto",
    color_continuous_scale="YlOrRd"
)

fig_heat.update_layout(
    xaxis_title="Month",
    yaxis_title="Department",
    height=450
)

st.plotly_chart(fig_heat, use_container_width=True)

#st.write(combined_data[["Month","Area","Daily_OT"]].head(20))
# -------------------------------------------------
# DETAILED SUMMARY (DROPDOWN)
# -------------------------------------------------

st.markdown("## 📊 Detailed Summary")

with st.expander("View Detailed Summary"):

    st.subheader("Combined Summary")

    st.dataframe(summary, use_container_width=True)

st.markdown("---")
