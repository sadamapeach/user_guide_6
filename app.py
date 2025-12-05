import streamlit as st
import pandas as pd
import numpy as np
import time
import re
import io
import zipfile
from io import BytesIO

def format_rupiah(x):
    if pd.isna(x):
        return ""
    # pastikan bisa diubah ke float
    try:
        x = float(x)
    except:
        return x  # biarin apa adanya kalau bukan angka

    # kalau tidak punya desimal (misal 7000.0), tampilkan tanpa ,00
    if x.is_integer():
        formatted = f"{int(x):,}".replace(",", ".")
    else:
        formatted = f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        # hapus ,00 kalau desimalnya 0 semua (misal 7000,00 → 7000)
        if formatted.endswith(",00"):
            formatted = formatted[:-3]
    return formatted

def highlight_total(row):
    # Cek apakah ada kolom yang berisi "TOTAL" (case-insensitive)
    if any(str(x).strip().upper() == "TOTAL" for x in row):
        return ["font-weight: bold; background-color: #D9EAD3; color: #1A5E20;"] * len(row)
    else:
        return [""] * len(row)
    
def highlight_1st_2nd_vendor(row, columns):
    styles = [""] * len(columns)
    first_vendor = row.get("1st Vendor")
    second_vendor = row.get("2nd Vendor")

    for i, col in enumerate(columns):
        if col == first_vendor:
            # styles[i] = "background-color: #f8c8dc; color: #7a1f47;"
            styles[i] = "background-color: #C6EFCE; color: #006100;"
        elif col == second_vendor:
            # styles[i] = "background-color: #d7c6f3; color: #402e72;"
            styles[i] = "background-color: #FFEB9C; color: #9C6500;"
    return styles

st.subheader("🧑‍🏫 User Guide: UPL Comparison Round by Round")
st.markdown(
    ":red-badge[Indosat] :orange-badge[Ooredoo] :green-badge[Hutchison]"
)
st.caption("INSPIRE 2025 | Oktaviana Sadama Nur Azizah")

# Divider custom
st.markdown(
    """
    <hr style="margin-top:-5px; margin-bottom:10px; border: none; height: 2px; background-color: #ddd;">
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
    <div style="
        display: flex;
        align-items: center;
        height: 65px;
        margin-bottom: 10px;
    ">
        <div style="text-align: justify; font-size: 15px;">
            <span style="color: #FFCB09; font-weight: 800;">
            UPL Comparison Round by Round</span>
            tracks UPL adjustments throughout negotiation rounds to understand pricing dynamics
            at the item level.
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

st.markdown("#### Input Structure")

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 20px;">
            The input file required for this menu should be a 
            <span style="color: #FF69B4; font-weight: 500;">multi-file containing multiple sheets</span>, in eather 
            <span style="background:#C6EFCE; padding:1px 4px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">.xlsx</span> or 
            <span style="background:#FFEB9C; padding:2px 4px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">.xls</span> format. 
            The file name represents the 
            <span style="font-weight: bold;">"ROUND"</span>, while the sheet names represent the vendor name. 
            Below is the table structure for each sheet.
        </div>
    """,
    unsafe_allow_html=True
)

# Dataframe
columns = ["Scope", "Desc", "Category", "UoM", "PRICE"]
df = pd.DataFrame([[""] * len(columns) for _ in range(3)], columns=columns)

st.dataframe(df, hide_index=True)

# Buat DataFrame 1 row
st.markdown("""
<table style="width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 15px;">
    <tr>
        <td style="border: 1px solid gray; width: 15%;">Vendor A</td>
        <td style="border: 1px solid gray; width: 15%;">Vendor B</td>
        <td style="border: 1px solid gray; width: 15%;">Vendor C</td>
        <td style="border: 1px solid gray; font-style: italic; color: #26BDAD">multiple sheets</td>
    </tr>
</table>
""", unsafe_allow_html=True)

st.markdown("###### Description:")
st.markdown(
    """
    <div style="font-size:15px;">
        <ul>
            <li>
                <span style="display:inline-block; width:100px;">Scope - UoM</span>: non-numeric columns
            </li>
            <li>
                <span style="display:inline-block; width:100px;">PRICE</span>: numeric column
            </li>
        </ul>
    </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 20px">
            The system accommodates a 
            <span style="font-weight: bold;">dynamic table</span>, but it is 
            <span style="color: #FF69B4; font-weight: 500;">ONLY APPLICABLE</span> to 
            <span style="color: #FF69B4; font-weight: 500;">non-numeric columns</span>. Unlike other menus, 
            <span style="background:#FFCB09; padding:2px 4px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">NUMERIC COLUMN</span> is permitted 
            <span style="color: #ED1C24; font-weight: bold;">ONLY ONCE</span> and 
            <span style="color: #ED1C24; font-weight: bold;">MUST</span> be placed in the last column.
            Also, users have the freedom to name the columns as they wish. The system logic relies on 
            <span style="font-weight: bold;">column indices</span>, not specific column names.
        </div>
    """,
    unsafe_allow_html=True 
)

st.markdown("**:violet-badge[Ensure that each sheet has the same table structure and column names!]**")

st.divider()
st.markdown("#### Constraint")

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 20px; margin-top: -10px">
            To ensure this menu works correctly, users need to follow certain rules regarding
            the dataset structure.
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown("**:red-badge[1. MULTIPLE FILE NAME]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 15px; margin-top: -10px">
            This menu operates using 
            <span style="color: #FF69B4; font-weight: 500;">multiple files</span>, where each filename is extracted and used as the value for the 
            <span style="background:#FFCB09; padding:2px 4px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">ROUND</span> column. 
            Therefore, please ensure that each filename correctly represents its corresponding round and 
            <span style="color: #ED1C24; font-weight: bold;">AVOID</span> using ambiguous names.
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px;">
            Because the filenames are parsed into a non-numeric column, the system uses 
            <span style="color: #FF69B4; font-weight: 700;">REGEX</span> to detect and sort the rounds in the correct order. 
            For example:
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
        <div style="text-align: center; font-size: 15px; margin-bottom: 10px;">
            <span style="background: #FF5E5E; padding:1px 4px; border-radius:6px; font-weight:600; font-size: 13px; color: black">L2R2.xlsx</span>  |
            <span style="background: #FF00AA; padding:1px 4px; border-radius:6px; font-weight:600; font-size: 13px; color: black">L2R4.xlsx</span>  |
            <span style="background: #FF5E5E; padding:1px 4px; border-radius:6px; font-weight:600; font-size: 13px; color: black">L2R3.xlsx</span>  |
            <span style="background: #FF00AA; padding:1px 4px; border-radius:6px; font-weight:600; font-size: 13px; color: black">L2R1.xlsx</span>
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 15px">
            will automatically be sorted as: 
            <span style="font-weight: bold;">L2R1</span> →
            <span style="font-weight: bold;">L2R2</span> →
            <span style="font-weight: bold;">L2R3</span> →
            <span style="font-weight: bold;">L2R4</span>
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 3px; font-weight: bold;">
            Why is this important?
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 25px;">
            Because the order of the rounds directly affects the analysis of 
            <span style="color: #69FFB4; font-weight: 700;">PRICE MOVEMENT</span>.
            It is highly recommended to use clear and consistent naming, such as 
            <span style="background:#FF9A09; padding:2px 4px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">Round 1</span>
            <span style="background:#FF9A09; padding:2px 4px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">Round 2</span> 
            and so on.
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown("**:orange-badge[2. COLUMN ORDER]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top: -10px">
            When creating tables, it is important to follow the specified column structure. Columns 
            <span style="font-weight: bold;">must</span> be arranged in the following order:
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
        <div style="text-align: center; font-size: 15px; margin-bottom: 10px; font-weight: bold">
            Non-Numeric Columns → Numeric Column (only one)
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 25px">
            this order is <span style="color: #FF69B4; font-weight: 700;">strict</span> and 
            <span style="color: #FF69B4; font-weight: 700;">cannot be altered</span>!
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown("**:green-badge[3. NUMBER COLUMN]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            Please refer the table below:
        </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["No", "Scope", "Desc", "Category", "UoM", "PRICE"]
data = [
    [1] + [""] * (len(columns) - 1),
    [2] + [""] * (len(columns) - 1),
    [3] + [""] * (len(columns) - 1)
]
df = pd.DataFrame(data, columns=columns)

st.dataframe(df, hide_index=True)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 25px; margin-top: -5px;">
            The table above is an <span style="color: #FF69B4; font-weight: 700;">incorrect example</span>
            and is <span style="color: #FF69B4; font-weight: 700;">not allowed</span> because it contains 
            a <span style="font-weight: bold;">"No"</span> column. The "No" column is prohibited in this
            menu, as it will be treated as a numeric column by the system, which violates the constraint
            described in point 2.
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown("**:blue-badge[4. FLOATING TABLE]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            Floating tables are allowed, meaning tables <span style="color: #FF69B4; font-weight: 700;">
            do not need to start from cell A1</span>. However, ensure
            that the cells above and to the left of the table are empty, as shown in the example below:
        </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["", "A", "B", "C", "D", "E", "F"]

# Buat 5 baris kosong
df = pd.DataFrame([[""] * len(columns) for _ in range(6)], columns=columns)

# Isi kolom pertama dengan 1–6
df.iloc[:, 0] = [1, 2, 3, 4, 5, 6]

# Header bagian kedua
df.loc[1, ["B", "C", "D", "E"]] = ["Desc", "Category", "UoM", "PRICE"]

# Data Software & Hardware
df.loc[2, ["B", "C", "D", "E"]] = ["Optical Cable", "Non-Services Area & Material", "M", "3.600"]
df.loc[3, ["B", "C", "D", "E"]] = ["Cross Connect", "Non-Services Area & Material", "Link", "29.800"]
df.loc[4, ["B", "C", "D", "E"]] = ["Dismantle RAU", "Non-Services Area & Material", "Pcs", "274.450"]

st.dataframe(df, hide_index=True)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 25px; margin-top:-10px;">
            To provide additional explanations or notes on the sheet, you can include them using an image or a text box.
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown("**:violet-badge[5. TOTAL ROW]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            You are not allowed to add a 
            <span style="font-weight: 700;">TOTAL</span> row at the bottom of the table! 
            Please refer to the example table below:
        </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["Desc", "Category", "UoM", "PRICE"]
data = [
    ["Optical Cable", "Non-Services Area & Material", "M", "3.600"],
    ["Cross Connect", "Non-Services Area & Material", "Link", "29.800"],
    ["TOTAL", "", "", "33.400"],
]
df = pd.DataFrame(data, columns=columns)

def red_highlight(row):
    if any(str(x).strip().upper() == "TOTAL" for x in row):
        return [
            "background-color: #FFE5E5; color: #D00000; font-weight: 700;"
        ] * len(row)
    return [""] * len(row)


df_styled = df.style.apply(red_highlight, axis=1)

st.dataframe(df_styled, hide_index=True)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 20px; margin-top: -5px;">
            The table above is an 
            <span style="color: #FF69B4; font-weight: 700;">incorrect example</span> and is 
            <span style="color: #FF69B4; font-weight: 700;">not permitted</span>! 
            The total row is generated automatically during
            <span style="font-weight: 700;">MERGE DATA</span> — 
            do not add one manually, or the system will treat it as a regular row and include it in the calculations.
        </div>
    """,
    unsafe_allow_html=True
)

st.divider()

st.markdown("#### What is Displayed?")

# Path file Excel yang sudah ada
file_paths = ["Round 1.xlsx", "Round 2.xlsx", "Round 3.xlsx", "Round 4.xlsx"]

# Buat ZIP di memory
zip_buffer = io.BytesIO()
with zipfile.ZipFile(zip_buffer, "w") as zf:
    for file_path in file_paths:
        zf.write(file_path, arcname=file_path.split("/")[-1])  # arcname = nama file di ZIP
zip_buffer.seek(0)

# Markdown teks
st.markdown(
    """
    <div style="text-align: justify; font-size: 15px; margin-bottom: 5px; margin-top: -10px">
        You can try this menu by downloading the dummy dataset using the button below: 
    </div>
    """,
    unsafe_allow_html=True
)

@st.fragment
def release_the_balloons():
    st.balloons()

# Download button untuk file Excel
st.download_button(
    label="Dummy Dataset",
    data=zip_buffer,
    file_name="Dummy Dataset - UPL Comparison Round by Round.zip",
    mime="application/zip",
    on_click=release_the_balloons,
    type="primary",
    use_container_width=True,
)

st.markdown(
    """
    <div style="text-align: justify; font-size: 15px; margin-bottom: 20px">
        Based on this dummy dataset, the menu will produce the following results.
    </div>
    """,
    unsafe_allow_html=True
)

st.markdown("**:red-badge[1. MERGE DATA]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            The system will merge the tables from each sheet into a single table and add
            a <span style="background:#FFCB09; padding:2px 4px; border-radius:6px; font-weight:600; 
            font-size: 0.75rem; color: black">TOTAL ROW</span> for each vendor, as shown below.
        </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["ROUND", "VENDOR", "Scope", "PRICE"]
data = [
    ["Round 1", "Vendor A", "Site Survey", 15000],
    ["Round 1", "Vendor A", "DG Dismantle", 55000],
    ["Round 1", "Vendor A", "AirCon Dismantle", 3230],
    ["Round 1", "Vendor A", "TOTAL", 73230],

    ["Round 1", "Vendor B", "Site Survey", 14800],
    ["Round 1", "Vendor B", "DG Dismantle", 55100],
    ["Round 1", "Vendor B", "AirCon Dismantle", 3240],
    ["Round 1", "Vendor B", "TOTAL", 73140],

    ["Round 1", "Vendor C", "Site Survey", 15050],
    ["Round 1", "Vendor C", "DG Dismantle", 54900],
    ["Round 1", "Vendor C", "AirCon Dismantle", 3200],
    ["Round 1", "Vendor C", "TOTAL", 73150],

    ["Round 2", "Vendor A", "Site Survey", 14950],
    ["Round 2", "Vendor A", "DG Dismantle", 54980],
    ["Round 2", "Vendor A", "AirCon Dismantle", 3200],
    ["Round 2", "Vendor A", "TOTAL", 73130],

    ["Round 2", "Vendor B", "Site Survey", 14800],
    ["Round 2", "Vendor B", "DG Dismantle", 55000],
    ["Round 2", "Vendor B", "AirCon Dismantle", 3240],
    ["Round 2", "Vendor B", "TOTAL", 73040],

    ["Round 2", "Vendor C", "Site Survey", 15000],
    ["Round 2", "Vendor C", "DG Dismantle", 54900],
    ["Round 2", "Vendor C", "AirCon Dismantle", 3200],
    ["Round 2", "Vendor C", "TOTAL", 73100],

    ["Round 3", "Vendor A", "Site Survey", 14900],
    ["Round 3", "Vendor A", "DG Dismantle", 54950],
    ["Round 3", "Vendor A", "AirCon Dismantle", 3150],
    ["Round 3", "Vendor A", "TOTAL", 73000],

    ["Round 3", "Vendor B", "Site Survey", 14750],
    ["Round 3", "Vendor B", "DG Dismantle", 54900],
    ["Round 3", "Vendor B", "AirCon Dismantle", 3220],
    ["Round 3", "Vendor B", "TOTAL", 73870],

    ["Round 3", "Vendor C", "Site Survey", 14900],
    ["Round 3", "Vendor C", "DG Dismantle", 54800],
    ["Round 3", "Vendor C", "AirCon Dismantle", 3175],
    ["Round 3", "Vendor C", "TOTAL", 72875],

    ["Round 4", "Vendor A", "Site Survey", 14900],
    ["Round 4", "Vendor A", "DG Dismantle", 54900],
    ["Round 4", "Vendor A", "AirCon Dismantle", 3150],
    ["Round 4", "Vendor A", "TOTAL", 73950],

    ["Round 4", "Vendor B", "Site Survey", 14750],
    ["Round 4", "Vendor B", "DG Dismantle", 54900],
    ["Round 4", "Vendor B", "AirCon Dismantle", 3200],
    ["Round 4", "Vendor B", "TOTAL", 73850],

    ["Round 4", "Vendor C", "Site Survey", 14850],
    ["Round 4", "Vendor C", "DG Dismantle", 54800],
    ["Round 4", "Vendor C", "AirCon Dismantle", 3175],
    ["Round 4", "Vendor C", "TOTAL", 72825],
]
df_merge = pd.DataFrame(data, columns=columns)

num_cols = ["PRICE"]
df_merge_styled = (
    df_merge.style
    .format({col: format_rupiah for col in num_cols})
    .apply(highlight_total, axis=1)
)

st.dataframe(df_merge_styled, hide_index=True)

st.write("")
st.markdown("**:orange-badge[2. PIVOT TABLE]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            After merging the data, the system will generate a pivot table by rearranging the table structure horizontally
            based on each 
            <span style="color: #FF69B4; font-weight: 700;">SCOPE</span>, as shown below.
        </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["Scope", "VENDOR A Round 1", "VENDOR A Round 2", "VENDOR A Round 3", "VENDOR A Round 4", "VENDOR B Round 1", "VENDOR B Round 2", "VENDOR B Round 3", "VENDOR B Round 4", "VENDOR C Round 1", "VENDOR C Round 2", "VENDOR C Round 3", "VENDOR C Round 4"]
data = [
    ["AirCon Dismantle", 3230,3200,3150,3150,3240,3240,3220,3200,3200,3200,3175,3175],
    ["DG Dismantle", 55000,54980,54950,54900,55100,55000,54900,54900,54900,54900,54800,54800],
    ["Site Survey", 15000,14950,14900,14900,14800,14800,14750,14750,15050,15000,14900,14850],
    ["TOTAL", 73230,73130,73000,72950,73140,73040,72870,72850,73150,73100,72875,72825]
]
df_pivot = pd.DataFrame(data, columns=columns)

num_cols = ["VENDOR A Round 1", "VENDOR A Round 2", "VENDOR A Round 3", "VENDOR A Round 4", "VENDOR B Round 1", "VENDOR B Round 2", "VENDOR B Round 3", "VENDOR B Round 4", "VENDOR C Round 1", "VENDOR C Round 2", "VENDOR C Round 3", "VENDOR C Round 4"]
df_pivot_styled = (
    df_pivot.style
    .format({col: format_rupiah for col in num_cols})
    .apply(highlight_total, axis=1)
)
st.dataframe(df_pivot_styled, hide_index=True)

st.write("")
st.markdown("**:yellow-badge[3. BID & PRICE ANALYSIS]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            This menu also displays an analysis table that provides a comprehensive overview of the pricing structure 
            submitted by each vendor, as follows.
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
    <div style="text-align:left; margin-bottom: 8px">
        <span style="background:#C6EFCE; padding:2px 8px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">1st Lowest</span>
        &nbsp;
        <span style="background:#FFEB9C; padding:2px 8px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">2nd Lowest</span>
    </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["ROUND", "Scope", "VENDOR A", "VENDOR B", "VENDOR C", "1st Lowest", "1st Vendor", "2nd Lowest", "2nd Vendor", "Gap 1 to 2 (%)", "Median Price", "Vendor A to Median (%)", "Vendor B to Median (%)", "Vendor C to Median (%)"]
data = [
    ["ROUND 1", "Site Survey", 15000, 14800, 15050, 14800, "VENDOR B", 15000, "VENDOR A", "1.4%", 15000, "+0.0%", "-1.3%", "+0.3%"],
    ["ROUND 1", "DG Dismantle", 55000, 55100, 54900, 54900, "VENDOR C", 55000, "VENDOR A", "0.2%", 55000, "+0.0%", "+0.2%", "-0.2%"],
    ["ROUND 1", "AirCon Dismantle", 3230, 3240, 3200, 3200, "VENDOR C", 3230, "VENDOR A", "0.9%", 3230, "+0.0%", "+0.3%", "-0.9%"],

    ["ROUND 2", "Site Survey", 14950, 14800, 15000, 14800, "VENDOR B", 14950, "VENDOR A", "1.0%", 14950, "+0.0%", "-1.0%", "+0.3%"],
    ["ROUND 2", "DG Dismantle", 54980, 55000, 54900, 54900, "VENDOR C", 54980, "VENDOR A", "0.1%", 54980, "+0.0%", "+0.0%", "-0.1%"],
    ["ROUND 2", "AirCon Dismantle", 3200, 3240, 3200, 3200, "VENDOR A", 3240, "VENDOR B", "1.2%", 3200, "+0.0%", "+1.2%", "+0.0%"],

    ["ROUND 3", "Site Survey", 14900, 14750, 14900, 14750 ,"VENDOR B", 14900, "VENDOR A", "1.0%", 14900, "+0.0%", "-1.0%", "+0.0%"],
    ["ROUND 3", "DG Dismantle", 54950, 54900, 54800, 54800 ,"VENDOR C", 54900, "VENDOR B", "0.2%", 54900, "+0.1%", "+0.0%", "-0.2%"],
    ["ROUND 3", "AirCon Dismantle", 3150, 3220, 3175, 3150, "VENDOR A", 3175, "VENDOR C", "0.8%", 3175, "-0.8%", "+1.4%", "+0.0%"],

    ["ROUND 4", "Site Survey", 14900, 14750, 14850, 14750 ,"VENDOR B", 14850,"VENDOR C", "0.7%", 14850, "+0.3%", "-0.7%", "+0.0%"],
    ["ROUND 4", "DG Dismantle", 54900, 54900, 54800, 54800 ,"VENDOR C", 54900,"VENDOR A", "0.2%", 54900, "+0.0%", "+0.0%", "-0.2%"],
    ["ROUND 4", "AirCon Dismantle", 3150, 3200, 3175, 3150, "VENDOR A", 3175, "VENDOR C", "0.8%", 3175, "-0.8%", "+0.8%", "+0.0%"],
]
df_analysis = pd.DataFrame(data, columns=columns)

num_cols = ["VENDOR A", "VENDOR B", "VENDOR C", "1st Lowest", "2nd Lowest", "Median Price"]
df_analysis_styled = (
    df_analysis.style
    .format({col: format_rupiah for col in num_cols})
    .apply(lambda row: highlight_1st_2nd_vendor(row, df_analysis.columns), axis=1)
)

st.dataframe(df_analysis_styled, hide_index=True)

st.write("")
st.markdown("**:green-badge[4. PRICE MOVEMENT ANALYSIS]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            The system also generates a price analysis table to compare price 
            <span style="color: #FF69B4">decreases</span> or 
            <span style="color: #FF69B4">increases</span> 
            across each round, as follows.
        </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["VENDOR", "Scope", "Round 1", "Round 2", "Round 3", "Round 4", "PRICE REDUCTION (VALUE)", "PRICE REDUCTION (%)", "PRICE TREND", "STANDARD DEVIATION", "PRICE STABILITY INDEX (%)"]
data = [
    ["VENDOR A", "AirCon Dismantle", 3230,3200,3150,3150,-80,"-2.5%","Fluctuating",34.187,"2.5%"],
    ["VENDOR A", "DG Dismantle", 55000,54980,54950,54900,-100,"-0.2%","Consistently Down",37.6663,"0.2%"],
    ["VENDOR A", "Site Survey", 15000,14950,14900,14900,-100,"-0.7%","Fluctuating",41.4578,"0.7%"],
    ["VENDOR A", "TOTAL", 73230,73130,73000,72950,"","","","",""],

    ["VENDOR B", "AirCon Dismantle", 3240,3240,3220,3200,-40,"-1.2%","Fluctuating",16.5831,"1.2%"],
    ["VENDOR B", "DG Dismantle", 55100,55000,54900,54900,-200,"-0.4%","Fluctuating",82.9156,"0.4%"],
    ["VENDOR B", "Site Survey", 14800,14800,14750,14750,-50,"-0.3%","Fluctuating",25,"0.3%"],
    ["VENDOR B", "TOTAL", 73140,73040,72870,72850,"","","","",""],

    ["VENDOR C", "AirCon Dismantle", 3200,3200,3175,3175,-25,"-0.8%","Fluctuating",12.5,"0.8%"],
    ["VENDOR C", "DG Dismantle", 54900,54900,54800,54800,-100,"-0.2%","Fluctuating",50,"0.2%"],
    ["VENDOR C", "Site Survey", 15050,15000,14900,14850,-200,"-1.3%","Consistently Down",79.0569,"1.3%"],
    ["VENDOR C", "TOTAL",73150,73100,72875,72825,"","","","",""],
]

df_pmove = pd.DataFrame(data, columns=columns)
df_pmove = df_pmove.map(lambda x: None if x == "" else x)

num_cols = ["Round 1", "Round 2", "Round 3", "Round 4", "PRICE REDUCTION (VALUE)", "STANDARD DEVIATION"]
df_pmove_styled = (
    df_pmove.style
    .format({col: format_rupiah for col in num_cols})
    .apply(highlight_total, axis=1)
)

st.dataframe(df_pmove_styled, hide_index=True)

st.write("")
st.markdown("**:blue-badge[5. VISUALIZATION]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            This menu displays visualizations focusing on two key aspects: 
            <span style="background: #FF5E5E; padding:1px 4px; border-radius:6px; font-weight:600; font-size: 13px; color: black">Winning Performance</span> and 
            <span style="background: #FF00AA; padding:2px 4px; border-radius:6px; font-weight:600; font-size: 13px; color: black">Price Trend</span>, 
            each presented in its own tab.
        </div>
    """,
    unsafe_allow_html=True
)

tab1, tab2 = st.tabs(["Winning Performance", "Price Trend"])

with tab1:
    st.image("assets/1.png")
    with st.expander("See explanation"):
            st.caption('''
                The visualization above shows the number of wins each vendor
                achieves in every tender round. A win is counted based on which
                vendor becomes the best bidder **(1st Vendor)** for each scope.
                     
                **💡 How to interpret the chart**
                     
                - High Wins Value  
                     Vendor is highly competitive in that round and wins more scopes
                     than others.  
                - Increasing Wins Across Rounds  
                     Indicates improving perfomance or more competitive pricing in later 
                     rounds.  
                - Decreasing Wins Across Rounds  
                     Shows declining competitiveness, with the vendor losing more scopes
                     compared the previous rounds.  
                - Zero Wins in a Round  
                     Vendor did not win any scope in that round, indicating weak competitiveness
                     for that stage.
            ''')

with tab2:
    st.image("assets/2.png")
    with st.expander("See explanation"):
            st.caption('''
                The chart above shows the number of occurrences of each **Price 
                Trend** for every vendor based on the pivoted tender data.
                     
                **💡 How to interpret the chart**
                     
                - No Change  
                     The vendor's price remains stable across all rounds or periods.
                - Consistently Down  
                     The vendor's price decreases continuously from one round to the next.
                - Consistently Up  
                     The vendor's price increases in every subsequent round.
                - Fluctuating  
                     The vendor's price moves up and down across the rounds.
            ''')
    
st.write("")
st.markdown("**:violet-badge[6. SUPER BUTTON]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            Lastly, there is a <span style="background:#FFCB09; padding:2px 4px; border-radius:6px; font-weight:600; 
            font-size: 0.75rem; color: black">Super Button</span> feature where all dataframes generated by the system 
            can be downloaded as a single file with multiple sheets. You can also customize the order of the sheets.
            The interface looks more or less like this.
        </div>
    """,
    unsafe_allow_html=True
)

dataframes = {
    "Merge Data": df_merge,
    "Pivot Table": df_pivot,
    "Bid & Price Analysis": df_analysis,
    "Price Movement Analysis": df_pmove,
}

# Tampilkan multiselect
selected_sheets = st.multiselect(
    "Select sheets to download in a single Excel file:",
    options=list(dataframes.keys()),
    default=list(dataframes.keys())  # default semua dipilih
)

# Fungsi "Super Button" & Formatting
def generate_multi_sheet_excel(selected_sheets, df_dict):
    """
    Buat Excel multi-sheet dengan formatting:
    - Sheet 'Bid & Price Analysis' → highlight 1st & 2nd vendor
    - Sheet lainnya → highlight TOTAL
    """
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for sheet in selected_sheets:
            df = df_dict[sheet].copy()

            # Identifikasi num cols
            df_to_write = df.copy()
            numeric_cols = []

            for col in df.columns:
                # TRY CONVERT -> kalo gagal, tetap pakai kolom asli
                try:
                    coerced = pd.to_numeric(df[col])
                except Exception:
                    coerced = df[col]

                # Check apakah kolom benar-benar numerik
                coerced_check = pd.to_numeric(df[col], errors="coerce")

                if coerced_check.notna().any():
                    numeric_cols.append(col)
                    df_to_write[col] = coerced_check
                else:
                    df_to_write[col] = df[col]

            # vendor columns (hanya untuk Bid & Price Analysis)
            vendor_cols = [c for c in numeric_cols] if sheet == "Bid & Price Analysis" else []

            # Tulis dataframe ke excel 
            df_to_write.to_excel(writer, index=False, sheet_name=sheet)
            workbook  = writer.book
            worksheet = writer.sheets[sheet]

            # Format
            fmt_rupiah = workbook.add_format({"num_format": "#,##0"})
            fmt_pct    = workbook.add_format({"num_format": '#,##0.0"%"'})

            fmt_total  = workbook.add_format({
                "bold": True,
                "bg_color": "#D9EAD3",
                "font_color": "#1A5E20",
                "num_format": "#,##0"
            })

            fmt_first  = workbook.add_format({"bg_color": "#C6EFCE", "num_format": "#,##0"})
            fmt_second = workbook.add_format({"bg_color": "#FFEB9C", "num_format": "#,##0"})

            # Format lagii
            for col_idx, col_name in enumerate(df_to_write.columns):
                if col_name in numeric_cols:
                    worksheet.set_column(col_idx, col_idx, 15, fmt_rupiah)
                if "%" in col_name:
                    worksheet.set_column(col_idx, col_idx, 15, fmt_pct)

            # Highlight
            for row_idx, row in enumerate(df_to_write.itertuples(index=False), start=1):

                # Cek apakah baris TOTAL
                is_total_row = any(
                    isinstance(x, str) and x.strip().upper() == "TOTAL"
                    for x in row
                    if pd.notna(x)
                )

                # Ambil nama 1st & 2nd vendor (jika sheet Bid & Price Analysis)
                if sheet == "Bid & Price Analysis":
                    first_vendor_name = row[df.columns.get_loc("1st Vendor")]
                    second_vendor_name = row[df.columns.get_loc("2nd Vendor")]

                    first_idx = df.columns.get_loc(first_vendor_name) if first_vendor_name in vendor_cols else None
                    second_idx = df.columns.get_loc(second_vendor_name) if second_vendor_name in vendor_cols else None

                # Loop tiap kolom → apply format
                for col_idx, col_name in enumerate(df_to_write.columns):
                    value = row[col_idx]
                    fmt = None

                    # TOTAL row highlight
                    if is_total_row and sheet != "Bid & Price Analysis":
                        fmt = fmt_total

                    # Highlight 1st & 2nd vendor
                    if sheet == "Bid & Price Analysis":
                        if first_idx is not None and col_idx == first_idx:
                            fmt = fmt_first
                        elif second_idx is not None and col_idx == second_idx:
                            fmt = fmt_second

                    # Replace NaN/inf with empty string
                    if pd.isna(value) or (isinstance(value, float) and np.isinf(value)):
                        worksheet.write_blank(row_idx, col_idx, None, fmt)
                    else:
                        worksheet.write(row_idx, col_idx, value, fmt)

    output.seek(0)
    return output

# ---- DOWNLOAD BUTTON ----
if selected_sheets:
    excel_bytes = generate_multi_sheet_excel(selected_sheets, dataframes)

    st.download_button(
        label="Download",
        data=excel_bytes,
        file_name="Super Botton - UPL Comparison Round by Round.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

st.write("")
st.divider()

st.markdown("#### Video Tutorial")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            I have also included a video tutorial, which you can access through the 
            <span style="background:#FF0000; padding:2px 4px; border-radius:6px; font-weight:600; 
            font-size: 0.75rem; color: black">YouTube</span> link below.
        </div>
    """,
    unsafe_allow_html=True
)

st.video("https://youtu.be/yZTRQbr3sqA?si=-eNXGSLwrhV2by0C")