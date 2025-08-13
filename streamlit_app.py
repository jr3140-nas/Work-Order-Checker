
import io
import math
import pandas as pd
import streamlit as st
from datetime import datetime, date
from typing import Dict, Any, List
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

st.set_page_config(page_title="Craft-Based Daily Report", layout="wide")

CRAFT_MAP = {
  "1145480": "Alloy Mech Days",
  "1145560": "AOD Elec Days",
  "1146669": "AOD Elec Days",
  "1145463": "AOD Mech Days",
  "1145498": "Baghouse Mech Days",
  "1145594": "Caster Elec Days",
  "1145501": "Caster Mech Days",
  "1145551": "EAF Elec Days",
  "1145674": "Turns",
  "1145455": "EAF Mech Days",
  "1145631": "HVAC Elec Days",
  "1145623": "Preheater Elec Days",
  "1157755": "Segment Shop",
  "1145658": "Turns",
  "1145666": "Turns",
  "1146757": "Utilities Mech Days",
  "1162511": "Utilities Mech Days",
  "1152989": "WTP Mech Days"
}

EXPECTED_COLS = [
    "AddressBookNumber","Name","Production Date","OrderNumber","Sum of Hours.","Hours Estimated",
    "Status","Type","PMFrequency","Description","Problem","Lead Area","Craft","CostCenter","UnitNumber","StructureTag"
]

def normalize_excel_date(v) -> str | None:
    if v is None or (isinstance(v, float) and math.isnan(v)) or v == "":
        return None
    if isinstance(v, (datetime, date)):
        return datetime(v.year, v.month, v.day).strftime("%m/%d/%Y")
    if isinstance(v, (int, float)):
        try:
            d = pd.to_datetime(v, unit="D", origin="1899-12-30")
            return d.strftime("%m/%d/%Y")
        except Exception:
            pass
        try:
            d = pd.to_datetime(v, unit="ms", origin="unix")
            return d.strftime("%m/%d/%Y")
        except Exception:
            pass
    try:
        d = pd.to_datetime(str(v), errors="coerce")
        if pd.notnull(d):
            return d.strftime("%m/%d/%Y")
    except Exception:
        pass
    return None

def numberish(v) -> float:
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        try:
            return float("".join(ch for ch in v if (ch.isdigit() or ch in ".-")))
        except Exception:
            return 0.0
    return 0.0

def build_report(df: pd.DataFrame, selected_date: str) -> Dict[str, List[Dict[str, Any]]]:
    df = df.copy()
    df["__ProdDate"] = df["Production Date"].apply(normalize_excel_date)
    df = df[df["__ProdDate"] == selected_date]
    df["__CraftDesc"] = df["Craft"].astype(str).str.strip().map(CRAFT_MAP).fillna("(Unmapped Craft)")

    groups = {}
    for _, r in df.iterrows():
        ck = r["__CraftDesc"]
        key = (ck, r["Name"], r["OrderNumber"])
        if key not in groups:
            groups[key] = {
                "Craft": ck,
                "Name": r["Name"],
                "OrderNumber": r["OrderNumber"],
                "SumOfHours": 0.0,
                "Type": set(),
                "Description": set(),
                "Problem": set(),
            }
        g = groups[key]
        g["SumOfHours"] += numberish(r.get("Sum of Hours.", 0))
        if isinstance(r.get("Type", ""), str) and r.get("Type", "").strip():
            g["Type"].add(r["Type"].strip())
        if isinstance(r.get("Description", ""), str) and r.get("Description", "").strip():
            g["Description"].add(r["Description"].strip())
        if isinstance(r.get("Problem", ""), str) and r.get("Problem", "").strip():
            g["Problem"].add(r["Problem"].strip())

    crafts: Dict[str, List[Dict[str, Any]]] = {}
    for (_, name, wo), v in groups.items():
        crafts.setdefault(v["Craft"], []).append({
            "Name": v["Name"],
            "Work Order #": v["OrderNumber"],
            "Sum of Hours": round(v["SumOfHours"], 2),
            "Type": "; ".join(sorted(v["Type"])),
            "Description": "; ".join(sorted(v["Description"])),
            "Problem": "; ".join(sorted(v["Problem"])),
        })

    for k in list(crafts.keys()):
        def wo_key(x):
            s = str(x.get("Work Order #", ""))
            return float(s) if s.isdigit() else float("inf")
        crafts[k].sort(key=wo_key)
    return crafts

def make_pdf(selected_date: str, crafts: Dict[str, List[Dict[str, Any]]]) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    story: List = []

    story += [Paragraph(f"Daily Report — {selected_date}", styles["Title"]), Spacer(1, 6),
              Paragraph("Sorted by Work Order # within each craft", styles["Normal"]), Spacer(1, 12)]

    header_style = styles["Heading2"]
    table_style = TableStyle([
        ("GRID", (0,0), (-1,-1), 0.25, colors.grey),
        ("BACKGROUND", (0,0), (-1,0), colors.whitesmoke),
        ("ALIGN", (0,0), (-1,0), "LEFT"),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,0), 9),
        ("FONTSIZE", (0,1), (-1,-1), 8),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
    ])

    first = True
    for craft, rows in crafts.items():
        if not first:
            story.append(Spacer(1, 12))
        story.append(Paragraph(craft, header_style))
        data = [["Name", "Work Order #", "Sum of Hours", "Type", "Description", "Problem"]]
        for r in rows:
            data.append([
                str(r.get("Name","")),
                str(r.get("Work Order #","")),
                f'{{r.get("Sum of Hours",0):.2f}}',
                str(r.get("Type","")),
                str(r.get("Description","")),
                str(r.get("Problem","")),
            ])
        tbl = Table(data, repeatRows=1, colWidths=[110, 90, 90, 90, 170, 170])
        tbl.setStyle(table_style)
        story.append(tbl)
        first = False
        story.append(Spacer(1, 6))

    doc.build(story)
    pdf = buf.getvalue()
    buf.close()
    return pdf

st.title("Craft-Based Daily Report (Excel → PDF)")

with st.sidebar:
    st.markdown(
        "**Instructions**\\n"
        "1. Upload the exported **Time on Work Order** spreadsheet (.xlsx).\\n"
        "2. Pick a **Production Date** (format **MM/DD/YYYY**).\\n"
        "3. Review the grouped report by **Craft Description**.\\n"
        "4. Click **Download PDF** to export.\\n\\n"
        "**Notes**\\n"
        "- Craft descriptions are **hard-coded** from your address sheet.\\n"
        "- Unmapped craft codes show as *(Unmapped Craft)*.\\n"
    )

uploaded = st.file_uploader("Upload Time on Work Order (.xlsx)", type=["xlsx"])

df = None
dates = []
if uploaded is not None:
    try:
        df = pd.read_excel(uploaded, header=2)  # 3rd row as header
        df.columns = [str(c).strip() for c in df.columns]
        missing = [c for c in EXPECTED_COLS if c not in df.columns]
        if missing:
            st.error(f"Missing expected columns: {{missing}}")
        else:
            dates = sorted({d for d in (df["Production Date"].apply(normalize_excel_date).dropna().tolist())})
    except Exception as e:
        st.exception(e)

selected_date = st.text_input("Production Date (MM/DD/YYYY)", value=(dates[-1] if dates else ""))

if df is not None and selected_date:
    crafts = build_report(df, selected_date)
    st.caption(f"Detected dates: {{(dates[0] if dates else '—')}} → {{(dates[-1] if dates else '—')}} • Unique dates: {{len(dates)}}")
    for craft, rows in crafts.items():
        st.subheader(craft)
        st.dataframe(pd.DataFrame(rows, columns=["Name","Work Order #","Sum of Hours","Type","Description","Problem"]))
    pdf_bytes = make_pdf(selected_date, crafts)
    st.download_button("Download PDF", data=pdf_bytes, file_name=f"nas_report_{{selected_date.replace('/', '-')}}.pdf", mime="application/pdf")

