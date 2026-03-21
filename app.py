import pandas as pd
import numpy as np

# -------------------------------
# FILE PATH
# -------------------------------
ATTENDANCE_FILE = "attendance.xlsx"

# -------------------------------
# HOLIDAYS
# -------------------------------
holiday_list = pd.to_datetime([
    "14-01-2026","26-01-2026","03-03-2026","21-03-2026",
    "01-05-2026","15-08-2026","28-08-2026","02-10-2026",
    "20-10-2026","09-11-2026","11-11-2026"
], format="%d-%m-%Y")

# -------------------------------
# HELPER FUNCTION
# -------------------------------
def normalize_hours(x):
    try:
        return round(float(x))
    except:
        return 0

# -------------------------------
# PROCESS SINGLE SHEET
# -------------------------------
def process_sheet(sheet):

    raw = pd.read_excel(ATTENDANCE_FILE, sheet_name=sheet, header=None)

    header_row = None
    for i, row in raw.iterrows():
        if row.astype(str).str.lower().str.contains("personnel").any():
            header_row = i
            break

    if header_row is None:
        return None

    df = pd.read_excel(ATTENDANCE_FILE, sheet_name=sheet, header=header_row)
    df.columns = df.columns.astype(str).str.strip()
    #print(df.columns.tolist())
    personnel_col = None
    name_col = None

    for col in df.columns:
        col_low = col.lower()
        if "personnel" in col_low:
            personnel_col = col
        if "employee" in col_low or "name" in col_low:
            name_col = col

    df = df.rename(columns={
        personnel_col: "Personnel Number",
        name_col: "Employee/app.name"
    })

    df = df.dropna(subset=["Personnel Number"])
    # Ensure required columns exist
    required_cols = ["Pay Scale Group", "Area"]
    
    # Ensure columns exist (avoid crash)
    for col in ["Pay Scale Group", "Area"]:
        if col not in df.columns:
            df[col] = "Unknown"

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
        id_vars=[
            "Personnel Number",
            "Employee/app.name",
            "Pay Scale Group",
            "Area"
        ],
        value_vars=date_columns,
        var_name="Date",
        value_name="Daily_Hours"
    )
    print("a")
    #print(df_long.head(5))
    print("b")
    df_long["Date"] = pd.to_datetime(df_long["Date"])
    df_long["Daily_Hours"] = pd.to_numeric(df_long["Daily_Hours"], errors="coerce").fillna(0)
    df_long["Daily_Hours"] = df_long["Daily_Hours"].apply(normalize_hours)

    df_long["Weekday"] = df_long["Date"].dt.weekday

    is_sunday = df_long["Weekday"] == 6
    is_holiday = df_long["Date"].isin(holiday_list)

    df_long["Daily_OT"] = np.where(
        is_sunday | is_holiday,
        df_long["Daily_Hours"],
        np.where(df_long["Daily_Hours"] > 9, df_long["Daily_Hours"] - 8, 0)
    )

    return df_long


# -------------------------------
# MAIN EXECUTION
# -------------------------------
xls = pd.ExcelFile(ATTENDANCE_FILE)

all_data = []
for sheet in xls.sheet_names:
    data = process_sheet(sheet)
    if data is not None:
        all_data.append(data)

combined_data = pd.concat(all_data, ignore_index=True)
print("c")
#print(combined_data.columns)
print("ds")

# -------------------------------
# WEEKLY
# -------------------------------
df_week = combined_data.copy()
df_week = df_week.sort_values(by=["Personnel Number", "Date"])

df_week["Week_Start"] = df_week["Date"] - pd.to_timedelta(
    (df_week["Date"].dt.weekday + 1) % 7, unit='D'
)
df_week["Week_End"] = df_week["Week_Start"] + pd.Timedelta(days=6)

weekly_summary = df_week.groupby(
    ["Personnel Number", "Employee/app.name","Area", "Week_Start", "Week_End"],
    as_index=False
)["Daily_Hours"].sum()

weekly_summary = weekly_summary.rename(columns={"Daily_Hours": "Weekly_Hours"})

weekly_summary["Weekly_OT"] = np.where(
    weekly_summary["Weekly_Hours"] > 48,
    weekly_summary["Weekly_Hours"] - 48,
    0
)


# -------------------------------
# CRITERIA 3: WEEKLY HOURS > 60 HEATMAP
# -------------------------------

# Sort for safety
weekly_summary = weekly_summary.sort_values(by=["Personnel Number", "Week_Start"])

# Add Week Number (for labeling Wk1, Wk2...)
weekly_summary["Week_Num"] = weekly_summary["Week_Start"].rank(method="dense").astype(int)
weekly_summary["Week_Label"] = "Wk" + weekly_summary["Week_Num"].astype(str)

# Filter only employees with weekly hours > 60
week_violation_df = weekly_summary[weekly_summary["Weekly_Hours"] > 60]

# Pivot for heatmap
heatmap_weekly = pd.pivot_table(
    week_violation_df,
    index="Week_Label",
    columns="Area",               # or "Pay Scale Group" if you want
    values="Personnel Number",
    aggfunc=pd.Series.nunique,
    fill_value=0
)

# Ensure order for weeks
week_order = ["Wk"+str(i) for i in range(1,13)]
heatmap_weekly = heatmap_weekly.reindex(week_order, fill_value=0)

print("✅ Weekly >60 hrs Heatmap Ready")













# -------------------------------
# QUARTERLY
# -------------------------------
df_quarter = weekly_summary.copy()
df_quarter["Year"] = df_quarter["Week_Start"].dt.year
df_quarter["Quarter"] = df_quarter["Week_Start"].dt.to_period("Q").astype(str)

quarterly_summary = df_quarter.groupby(
    ["Personnel Number", "Employee/app.name", "Year", "Quarter"],
    as_index=False
)["Weekly_OT"].sum()


# -------------------------------
# HEATMAP DATA (BCA ONLY)
# -------------------------------

df_heat = quarterly_summary.copy()

# -------------------------------
# HEATMAP DATA (BCA ONLY)
# -------------------------------
# -------------------------------
# HEATMAP DATA (BCA ONLY)
# -------------------------------

df_heat = combined_data.copy()

# ✅ ADD THIS (IMPORTANT FIX)
print("DEBUG → Columns in quarterly_summary:")
print(quarterly_summary.columns)

df_heat = df_heat.merge(
    quarterly_summary[[
        "Personnel Number",
        "Employee/app.name",
        "Weekly_OT"
    ]],
    on=["Personnel Number", "Employee/app.name"],
    how="left"
)

# Remove duplicates (one row per employee)
df_heat = df_heat.drop_duplicates(subset=["Personnel Number", "Employee/app.name"])

# Filter BCA
df_heat = df_heat[df_heat["Pay Scale Group"] == "BCA"]

def ot_bucket(x):
    if x < 10:
        return "<10 hrs"
    elif x < 20:
        return "10-20 hrs"
    elif x < 30:
        return "20-30 hrs"
    elif x < 40:
        return "30-40 hrs"
    elif x < 50:
        return "40-50 hrs"
    else:
        return ">50 hrs"

df_heat["OT_Range"] = df_heat["Weekly_OT"].fillna(0).apply(ot_bucket)

heatmap_table = pd.pivot_table(
    df_heat,
    index="OT_Range",
    columns="Area",
    values="Personnel Number",
    aggfunc="count",
    fill_value=0
)

order = ["<10 hrs", "10-20 hrs", "20-30 hrs", "30-40 hrs", "40-50 hrs", ">50 hrs"]
heatmap_table = heatmap_table.reindex(order)


quarterly_summary = quarterly_summary.rename(columns={"Weekly_OT": "Quarterly_OT"})

# -------------------------------
# CONTINUOUS
# -------------------------------
df_cont = combined_data.copy()

# Add Month
df_cont["Month"] = df_cont["Date"].dt.strftime("%b")

df_cont = df_cont.sort_values(by=["Personnel Number", "Date"])

df_cont["Working_Day"] = df_cont["Daily_Hours"] > 0
df_cont["Break"] = (~df_cont["Working_Day"]).astype(int)
df_cont["Streak_Group"] = df_cont.groupby(
    ["Personnel Number", "Month"]
)["Break"].cumsum()

df_cont["Continuous_Days"] = df_cont.groupby(
    ["Personnel Number", "Month", "Streak_Group"]
)["Working_Day"].cumsum()

print("✅ Continuous working days calculated")

continuous_summary = df_cont.groupby(
    ["Personnel Number", "Employee/app.name"],
    as_index=False
)["Continuous_Days"].max()

continuous_summary = continuous_summary.rename(columns={
    "Continuous_Days": "Max_Continuous_Days"
})

continuous_summary["Violation"] = np.where(
    continuous_summary["Max_Continuous_Days"] > 10,
    "Yes",
    "No"
)


#### new block 

# -------------------------------
# MONTH-WISE CONTINUOUS SUMMARY
# -------------------------------

df_cont["Month"] = df_cont["Date"].dt.strftime("%b")

df_cont["Streak_Group"] = df_cont.groupby(
    ["Personnel Number", "Month"]
)["Break"].cumsum()

df_cont["Continuous_Days"] = df_cont.groupby(
    ["Personnel Number", "Month", "Streak_Group"]
)["Working_Day"].cumsum()

monthly_cont_summary = df_cont.groupby(
    ["Personnel Number", "Employee/app.name", "Area", "Month"],
    as_index=False
)["Continuous_Days"].max()

monthly_cont_summary = monthly_cont_summary.rename(columns={
    "Continuous_Days": "Max_Continuous_Days"
})

monthly_cont_summary["Violation"] = np.where(
    monthly_cont_summary["Max_Continuous_Days"] > 10,
    "Yes",
    "No"
)

# Filter violations
df_violation = monthly_cont_summary[
    monthly_cont_summary["Violation"] == "Yes"
]

# Heatmap table
heatmap_cont = pd.pivot_table(
    df_violation,
    index="Month",
    columns="Area",
    values="Personnel Number",
    aggfunc=pd.Series.nunique,
    fill_value=0
)

# Month order fix
month_order = ["Jan","Feb","Mar","Apr","May","Jun",
               "Jul","Aug","Sep","Oct","Nov","Dec"]

heatmap_cont = heatmap_cont.reindex(month_order, fill_value=0)

print("✅ Continuous Heatmap Ready")







# =================================================
# STREAMLIT UI
# =================================================

import streamlit as st

st.set_page_config(layout="wide")

st.markdown("""<style>
div[data-testid="stVerticalBlock"]{gap:0.5rem;}
div[data-testid="stVerticalBlock"] > div{padding-top:0px;padding-bottom:0px;}
h1, h2, h3, h4 {margin-top:0px !important;margin-bottom:2px !important;}
div.stButton {margin-bottom:-5px !important;}
</style>""", unsafe_allow_html=True)

st.markdown("""<style>
.header-bar{background-color:#0086E2;padding:18px;border-radius:8px;margin-bottom:20px;}
.header-text{color:white;font-size:28px;font-weight:bold;text-align:center;}
</style>""", unsafe_allow_html=True)

st.markdown("""
<div class="header-bar">
<div class="header-text">OT Monitoring Dashboard</div>
</div>
""", unsafe_allow_html=True)


# -------------------------------
# SESSION STATE (TOGGLE)
# -------------------------------
if "show_week" not in st.session_state:
    st.session_state.show_week = False

if "show_ot" not in st.session_state:
    st.session_state.show_ot = False

if "show_cont" not in st.session_state:
    st.session_state.show_cont = False





# KPI
week_violation = weekly_summary[weekly_summary["Weekly_Hours"] > 60]
ot_violation = quarterly_summary[quarterly_summary["Quarterly_OT"] > 50]
cont_violation = continuous_summary[continuous_summary["Violation"] == "Yes"]

st.markdown("""<style>
div.stButton > button {width:100%;height:50px;border:none;background:transparent;}
</style>""", unsafe_allow_html=True)

c1, c2, c3 = st.columns(3)

with c1:
    st.markdown(f"""<div style="background:#ffe6e6;padding:25px;border-radius:12px;text-align:center;border:2px solid red;">
    <h2>⏱</h2><h4>Working hrs</h4><h4>>60 / Week</h4>
    <h2 style="color:red;">{len(week_violation)}</h2><p>Employees</p></div>""", unsafe_allow_html=True)
    if st.button("", key="week_tile"):
        st.session_state.show_week = not st.session_state.get("show_week", False)

with c2:
    st.markdown(f"""<div style="background:#fff3e0;padding:25px;border-radius:12px;text-align:center;border:2px solid orange;">
    <h2>📈</h2><h4>OT hrs</h4><h4>>50 / Quarter</h4>
    <h2 style="color:orange;">{len(ot_violation)}</h2><p>Employees</p></div>""", unsafe_allow_html=True)
    if st.button("", key="ot_tile"):
        st.session_state.show_week = not st.session_state.get("show_ot", False)

with c3:
    st.markdown(f"""<div style="background:#f3e5f5;padding:25px;border-radius:12px;text-align:center;border:2px solid purple;">
    <h2>🔁</h2><h4>Continuous Punch</h4><h4>>10 Days</h4>
    <h2 style="color:purple;">{len(cont_violation)}</h2><p>Employees</p></div>""", unsafe_allow_html=True)
    if st.button("", key="cont_tile"):
        st.session_state.show_week = not st.session_state.get("show_cont", False)





        
# -------------------------------------------------
# DETAILS SECTION (TOGGLE VIEW)
# -------------------------------------------------

if st.session_state.show_week:
    st.markdown("### 🚨 Weekly Working Hours Violation (>60 hrs)")
    st.dataframe(week_violation, use_container_width=True)

if st.session_state.show_ot:
    st.markdown("### 🚨 Quarterly OT Violation (>50 hrs)")
    st.dataframe(ot_violation, use_container_width=True)

if st.session_state.show_cont:
    st.markdown("### 🚨 Continuous Working Days Violation (>10 days)")
    st.dataframe(cont_violation, use_container_width=True)


# -------------------------------------------------
# HEATMAP VIEW
# -------------------------------------------------

st.markdown("### 🔥 OT Distribution Heatmap (BCA Employees)")

#st.dataframe(heatmap_table, use_container_width=True)


    
# -------------------------------
# HEATMAP DATA (BCA ONLY)
# -------------------------------

# Filter BCA employees
df_heat = quarterly_summary.copy()

# 🔴 IMPORTANT: Merge department + pay scale from original data
df_heat = df_heat.merge(
    combined_data[["Personnel Number", "Employee/app.name", "Pay Scale Group" , "Area"]].drop_duplicates(),
    on=["Personnel Number", "Employee/app.name"],
    how="left"
)

# Filter BCA
df_heat = df_heat[df_heat["Pay Scale Group"] == "BCA"]

# -------------------------------
# CREATE BUCKETS
# -------------------------------
def ot_bucket(x):
    if x < 10:
        return "<10 hrs"
    elif x < 20:
        return "10-20 hrs"
    elif x < 30:
        return "20-30 hrs"
    elif x < 40:
        return "30-40 hrs"
    elif x < 50:
        return "40-50 hrs"
    else:
        return ">50 hrs"

df_heat["OT_Range"] = df_heat["Quarterly_OT"].apply(ot_bucket)

# -------------------------------
# CREATE HEATMAP TABLE
# -------------------------------
heatmap_table = pd.pivot_table(
    df_heat,
    index="OT_Range",
    columns="Area",
    values="Personnel Number",
    aggfunc=pd.Series.nunique,   # ✅ UNIQUE COUNT FIX
    fill_value=0
)

# Ensure proper order
order = ["<10 hrs", "10-20 hrs", "20-30 hrs", "30-40 hrs", "40-50 hrs", ">50 hrs"]
#heatmap_table = heatmap_table.reindex(order)
heatmap_table = heatmap_table.reindex(order, fill_value=0)
import plotly.express as px

fig = px.imshow(
    heatmap_table,
    text_auto=True,
    aspect="auto",
    color_continuous_scale="Reds"
)

fig.update_layout(
    height=400,
    coloraxis_colorbar=dict(
        title="Employee Count",
        lenmode="fraction",
        len=0.5
    )
)

st.plotly_chart(fig, use_container_width=True)

print(heatmap_table.index.tolist())



st.markdown("### 🔁 Continuous Working >10 Days (Monthly Heatmap)")

import plotly.express as px

fig2 = px.imshow(
    heatmap_cont,
    text_auto=True,
    aspect="auto",
    color_continuous_scale="Purples"
)

fig2.update_layout(
    height=400,
    coloraxis_colorbar=dict(title="Employees")
)

st.plotly_chart(fig2, use_container_width=True)




# -------------------------------
# MONTH-WISE CONTINUOUS SUMMARY
# -------------------------------

monthly_cont_summary = df_cont.groupby(
    ["Personnel Number", "Employee/app.name", "Area", "Month"],
    as_index=False
)["Continuous_Days"].max()

monthly_cont_summary = monthly_cont_summary.rename(columns={
    "Continuous_Days": "Max_Continuous_Days"
})

# Violation flag (>10 days)
monthly_cont_summary["Violation"] = np.where(
    monthly_cont_summary["Max_Continuous_Days"] > 10,
    "Yes",
    "No"
)

df_violation = monthly_cont_summary[
    monthly_cont_summary["Violation"] == "Yes"
]

heatmap_cont = pd.pivot_table(
    df_violation,
    index="Month",
    columns="Area",
    values="Personnel Number",
    aggfunc=pd.Series.nunique,
    fill_value=0
)
month_order = ["Jan","Feb","Mar","Apr","May","Jun",
               "Jul","Aug","Sep","Oct","Nov","Dec"]

heatmap_cont = heatmap_cont.reindex(month_order, fill_value=0)


st.markdown("### 🔥 Weekly Working Hours >60 hrs Heatmap")
import plotly.express as px

fig3 = px.imshow(
    heatmap_weekly,
    text_auto=True,
    aspect="auto",
    color_continuous_scale="Reds"
)

fig3.update_layout(
    height=400,
    coloraxis_colorbar=dict(title="Employees")
)

st.plotly_chart(fig3, use_container_width=True)

