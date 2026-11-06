#!/usr/bin/env python3
"""Build the complete Kaldrix Data Analysis Mastery Programme.

Generates the same public entry-point names used by learners/trainers, 50 lesson
PDFs, 50 day-specific exercise datasets, programme PDFs, and a distributable ZIP.
"""
from __future__ import annotations

import base64
import json
import random
import shutil
import zipfile
from datetime import datetime, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak

ROOT = Path(__file__).resolve().parent
COURSE = json.loads((ROOT / "curriculum.json").read_text(encoding="utf-8"))
TEMPLATE = (ROOT / "src" / "Kaldrix_Data_Analysis_Mastery.template.html").read_text(encoding="utf-8")
LOGO_B64 = (ROOT / "assets" / "kaldrix_logo.base64.txt").read_text(encoding="ascii").strip()
LOGO_DATA = "data:image/png;base64," + LOGO_B64

LESSONS = ROOT / "lessons"
DATASETS = ROOT / "day_datasets"
LESSONS.mkdir(exist_ok=True)
DATASETS.mkdir(exist_ok=True)

random.seed(20261106)
FIRST = ["Amina","Brian","Carol","David","Edwin","Faith","Grace","Hassan","Irene","James","Kevin","Lilian","Mercy","Nancy","Oscar","Patricia","Rose","Samuel","Tracy","Victor","Winnie","Yvonne"]
LAST = ["Otieno","Omondi","Wekesa","Achieng","Mutiso","Mwangi","Ouma","Naliaka","Odhiambo","Wafula","Atieno","Kiptoo","Musyoka","Were","Okello","Barasa"]
COUNTIES = ["Kisumu","Kakamega","Siaya","Vihiga"]
PROGRAMS = ["Education","Agriculture","Digital Skills","Livelihoods"]
ACTIVITIES = ["Training","Field Visit","Mentorship","Support","Community Meeting"]
BASE_DATE = datetime(2026, 1, 1)

BLUE = "2F75B5"
NAVY = "17365D"
PALE = "D9EAF7"
WHITE = "FFFFFF"
THIN = Side(style="thin", color="D8E0E7")


def person() -> str:
    return f"{random.choice(FIRST)} {random.choice(LAST)}"


def style_header(ws, row=1):
    for cell in ws[row]:
        cell.fill = PatternFill("solid", fgColor=BLUE)
        cell.font = Font(bold=True, color=WHITE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=THIN)


def add_data_sheet(wb: Workbook, name: str, headers: list[str], rows: list[list], date_cols: list[int] | None = None):
    ws = wb.create_sheet(name)
    ws.append(headers)
    for row in rows:
        ws.append(row)
    style_header(ws, 1)
    ws.freeze_panes = "A2"
    date_cols = date_cols or []
    for idx, header in enumerate(headers, 1):
        width = 24 if "Name" in header else 16
        if header in {"Description", "Definition", "Question", "Rule"}:
            width = 38
        ws.column_dimensions[get_column_letter(idx)].width = width
        if idx in date_cols:
            for c in ws.iter_cols(min_col=idx, max_col=idx, min_row=2):
                for cell in c:
                    cell.number_format = "dd-mmm-yyyy"
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    return ws


def add_start_here(wb: Workbook, d: dict, sources: list[str]):
    ws = wb.create_sheet("START_HERE", 0)
    ws.merge_cells("A1:D1")
    c = ws["A1"]
    c.value = f"KALDRIX DAY {d['day']} EXERCISE DATASET"
    c.fill = PatternFill("solid", fgColor=NAVY)
    c.font = Font(bold=True, color=WHITE, size=14)
    c.alignment = Alignment(horizontal="center")
    meta = [
        ("Lesson", d["title"]),
        ("Level", d["level"]),
        ("Primary tool", d["tool"]),
        ("Source tab(s)", ", ".join(sources)),
        ("Independent exercise", d["exercise"]),
        ("Google Sheets workflow", "Upload/import this file into Google Sheets, complete the work there, share the Google Sheet with the trainer, then paste its URL into the Kaldrix app."),
    ]
    for r, (k, v) in enumerate(meta, 3):
        ws.cell(r, 1, k)
        ws.cell(r, 2, v)
        ws.cell(r, 1).fill = PatternFill("solid", fgColor=PALE)
        ws.cell(r, 1).font = Font(bold=True)
        ws.cell(r, 1).alignment = ws.cell(r, 2).alignment = Alignment(vertical="top", wrap_text=True)
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 90
    ws.append([])
    header_row = 10
    for col, value in enumerate(["No.", "Task", "Status"], 1):
        ws.cell(header_row, col, value)
    style_header(ws, header_row)
    tasks = list(d["objectives"]) + [
        d["practical"],
        "Create at least one QA/control check.",
        "Write 2-5 evidence-based findings or a short explanation appropriate to the lesson.",
    ]
    dv = DataValidation(type="list", formula1='"Not Started,In Progress,Completed"', allow_blank=False)
    ws.add_data_validation(dv)
    for i, task in enumerate(tasks, 1):
        r = header_row + i
        ws.cell(r, 1, i)
        ws.cell(r, 2, task)
        ws.cell(r, 3, "Not Started")
        dv.add(ws.cell(r, 3))
        for cell in ws[r]:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 80
    ws.column_dimensions["C"].width = 18


def build_sources(wb: Workbook, d: dict) -> list[str]:
    day = d["day"]
    src: list[str] = []
    if day <= 10:
        rows = []
        n = 75 if day < 10 else 140
        for i in range(n):
            total = random.choice([8, 10, 12]); att = random.randint(3, total); pre = random.randint(25, 80); post = min(100, pre + random.randint(-5, 25))
            rows.append([f"P-{day:02d}-{i+1:04d}", person(), random.choice(COUNTIES), random.choice(PROGRAMS), att, total, pre, post, random.choice(["Completed","Completed","Pending"]), random.choice([0,150,200,250,300,400])])
        if day in [7, 8, 10]:
            rows[4][1] = "  " + rows[4][1].upper() + " "; rows[11][2] = " kisumu "; rows[24][7] = 105; rows.append(list(rows[30]))
        src = ["DATA"]
        add_data_sheet(wb, "DATA", ["Record_ID","Participant_Name","County","Program","Sessions_Attended","Total_Sessions","Pre_Test","Post_Test","Status","Transport_KES"], rows)
    elif day in [11,12,16,17,18,19]:
        rows=[]
        for i in range(160):
            target=random.randint(20,140); actual=max(0,target+random.randint(-45,35))
            rows.append([f"R-{day:02d}-{i+1:04d}",f"P-{1000+random.randint(1,125):04d}",person(),random.choice(COUNTIES),random.choice(PROGRAMS),BASE_DATE+timedelta(days=random.randint(0,220)),target,actual,random.randint(500,14000),random.choice(["Low","Medium","High"])])
        if day==18:
            rows.append(list(rows[20])); rows[40][2]=None; rows[75][7]=190; rows[95][8]=-50
        src=["DATA"]
        add_data_sheet(wb,"DATA",["Record_ID","Person_ID","Person_Name","County","Program","Date","Target","Actual","Cost_KES","Risk"],rows,[6])
    elif day in [13,14]:
        master=[[f"P-{i:04d}",person(),random.choice(COUNTIES),random.choice(PROGRAMS)] for i in range(1001,1091)]
        visits=[[f"V-{i+1:04d}",random.choice([r[0] for r in master]+["P-X01","P-X02","P-X03"]),BASE_DATE+timedelta(days=random.randint(0,90)),random.choice(["Completed","Pending","Referred"])] for i in range(130)]
        src=["MASTER","VISITS"]
        add_data_sheet(wb,"MASTER",["Person_ID","Person_Name","County","Program"],master)
        add_data_sheet(wb,"VISITS",["Visit_ID","Person_ID","Visit_Date","Visit_Status"],visits,[3])
    elif day==15:
        reg=[[f"P-{i:04d}",person(),random.choice(COUNTIES),BASE_DATE+timedelta(days=random.randint(0,60))] for i in range(2001,2111)]
        ids=[r[0] for r in reg[:90]]+["P-UNREG-01","P-UNREG-02","P-UNREG-03"]
        att=[[f"A-{i+1:04d}",pid,BASE_DATE+timedelta(days=random.randint(30,100)),random.choice(["S1","S2","S3"])] for i,pid in enumerate(ids)]
        src=["REGISTRATION","ATTENDANCE"]
        add_data_sheet(wb,"REGISTRATION",["Person_ID","Person_Name","County","Registration_Date"],reg,[4])
        add_data_sheet(wb,"ATTENDANCE",["Attendance_ID","Person_ID","Attendance_Date","Session"],att,[3])
    elif day==20:
        people=[[f"P-{i:04d}",person(),random.choice(COUNTIES),random.choice(PROGRAMS)] for i in range(3001,3101)]
        acts=[[f"ACT-{i+1:04d}",random.choice([r[0] for r in people]+["P-X"]),BASE_DATE+timedelta(days=random.randint(0,180)),random.choice(ACTIVITIES),random.randint(1,20),random.randint(500,9000)] for i in range(240)]
        src=["PEOPLE","ACTIVITY"]
        add_data_sheet(wb,"PEOPLE",["Person_ID","Person_Name","County","Program"],people)
        add_data_sheet(wb,"ACTIVITY",["Activity_ID","Person_ID","Date","Activity_Type","Units","Cost_KES"],acts,[3])
    elif day in [21,22,24,28,29]:
        tx=[[f"TX-{i+1:05d}",BASE_DATE+timedelta(days=random.randint(0,330)),f"P-{random.randint(4001,4120)}",random.choice(COUNTIES),random.choice(PROGRAMS),random.choice(["Completed","Pending","Cancelled"]),random.randint(1,25),random.randint(500,18000)] for i in range(320)]
        src=["TRANSACTIONS"]
        add_data_sheet(wb,"TRANSACTIONS",["Transaction_ID","Date","Person_ID","County","Program","Status","Units","Amount_KES"],tx,[2])
    elif day==23:
        people=[[f"P-{i}",person(),random.choice(COUNTIES),random.choice(PROGRAMS)] for i in range(4001,4121)]
        acts=[[f"A-{i+1:05d}",random.choice([r[0] for r in people]+["P-9999"]),BASE_DATE+timedelta(days=random.randint(0,330)),random.choice(ACTIVITIES),random.randint(1,30)] for i in range(300)]
        src=["PEOPLE","ACTIVITY"]
        add_data_sheet(wb,"PEOPLE",["Person_ID","Person_Name","County","Program"],people)
        add_data_sheet(wb,"ACTIVITY",["Activity_ID","Person_ID","Activity_Date","Activity_Type","Actual"],acts,[3])
    elif day==25:
        a=[[f"A-{i+1:04d}",random.choice(COUNTIES),random.choice(PROGRAMS),BASE_DATE+timedelta(days=random.randint(0,150)),random.randint(1,30),random.randint(500,9000)] for i in range(120)]
        b=[[f"B-{i+1:04d}",random.choice(COUNTIES),random.choice(PROGRAMS),BASE_DATE+timedelta(days=random.randint(0,150)),random.randint(1,30),random.randint(500,9000)] for i in range(120)]
        src=["SOURCE_A","SOURCE_B"]
        add_data_sheet(wb,"SOURCE_A",["ID","County","Program","Date","Units","Amount_KES"],a,[4])
        add_data_sheet(wb,"SOURCE_B",["ID","County","Program","Date","Units","Amount_KES"],b,[4])
    elif day in [26,27]:
        people=[[f"P-{i}",person(),random.choice(COUNTIES),random.choice(PROGRAMS)] for i in range(5001,5111)]
        cal=[[BASE_DATE+timedelta(days=i),(BASE_DATE+timedelta(days=i)).year,(BASE_DATE+timedelta(days=i)).month] for i in range(365)]
        fact=[[f"F-{i+1:05d}",random.choice([r[0] for r in people]),BASE_DATE+timedelta(days=random.randint(0,364)),random.randint(10,120),random.randint(5,130),random.randint(500,12000)] for i in range(320)]
        src=["PEOPLE","CALENDAR","FACT"]
        add_data_sheet(wb,"PEOPLE",["Person_ID","Person_Name","County","Program"],people)
        add_data_sheet(wb,"CALENDAR",["Date","Year","Month"],cal,[1])
        add_data_sheet(wb,"FACT",["Fact_ID","Person_ID","Date","Target","Actual","Cost_KES"],fact,[3])
    elif day==30:
        people=[[f"P-{i}",person(),random.choice(COUNTIES),random.choice(PROGRAMS)] for i in range(6001,6131)]
        fact=[[f"F-{i+1:05d}",random.choice([r[0] for r in people]+["P-X"]),BASE_DATE+timedelta(days=random.randint(0,330)),random.randint(10,130),random.randint(5,150),random.randint(500,15000)] for i in range(420)]
        src=["PEOPLE","FACT"]
        add_data_sheet(wb,"PEOPLE",["Person_ID","Person_Name","County","Program"],people)
        add_data_sheet(wb,"FACT",["Fact_ID","Person_ID","Date","Target","Actual","Cost_KES"],fact,[3])
    elif day in [31,32,34,35,36,39]:
        rows=[]
        for i in range(400):
            target=random.randint(10,160); actual=max(0,target+random.randint(-55,45))
            rows.append([f"PY-{day:02d}-{i+1:05d}",BASE_DATE+timedelta(days=random.randint(0,365)),f"P-{random.randint(7001,7180)}",random.choice(COUNTIES),random.choice(PROGRAMS),random.choice(["A","B","C"]),target,actual,round(random.uniform(50,6000),2),random.choice(["Valid","Valid","Pending"])])
        if day in [32,35,39]:
            rows[20][3]=" kisumu "; rows[60][8]=None; rows[100][7]=999; rows.append(list(rows[150]))
        src=["DATA"]
        add_data_sheet(wb,"DATA",["Record_ID","Date","Person_ID","County","Program","Category","Target","Actual","Value","Status"],rows,[2])
    elif day==33:
        people=[[f"P-{i}",person(),random.choice(COUNTIES),random.choice(PROGRAMS)] for i in range(7001,7181)]
        acts=[[f"A-{i+1:05d}",random.choice([r[0] for r in people]+["P-X"]),BASE_DATE+timedelta(days=random.randint(0,365)),random.choice(ACTIVITIES),random.randint(1,30),random.randint(500,9000)] for i in range(400)]
        src=["PEOPLE","ACTIVITY"]
        add_data_sheet(wb,"PEOPLE",["Person_ID","Person_Name","County","Program"],people)
        add_data_sheet(wb,"ACTIVITY",["Activity_ID","Person_ID","Date","Activity_Type","Units","Cost_KES"],acts,[3])
    elif day==37:
        rows=[]
        for i in range(350):
            grp=random.choice(["Control","Treatment"]); bl=random.gauss(50,10); out=bl+(5 if grp=="Treatment" else 0)+random.gauss(0,8)
            rows.append([f"EXP-{i+1:04d}",grp,round(bl,2),round(out,2),random.choice(COUNTIES),random.choice(["Completed","Dropped"])])
        src=["EXPERIMENT"]
        add_data_sheet(wb,"EXPERIMENT",["ID","Group","Baseline","Outcome","County","Status"],rows)
    elif day==38:
        rows=[]; v=100
        for i in range(72):
            dt=datetime(2021+i//12,(i%12)+1,1); v=max(20,v+random.randint(-8,12)+(5 if i%12 in [10,11] else 0)); rows.append([dt,v])
        src=["TIMESERIES"]
        add_data_sheet(wb,"TIMESERIES",["Month","Value"],rows,[1])
    elif day==40:
        people=[[f"P-{i}",person(),random.choice(COUNTIES),random.choice(PROGRAMS)] for i in range(8001,8201)]
        fact=[[f"F-{i+1:05d}",random.choice([r[0] for r in people]+["P-X"]),BASE_DATE+timedelta(days=random.randint(0,365)),random.randint(10,160),random.randint(0,180),random.randint(500,15000)] for i in range(600)]
        src=["PEOPLE","FACT"]
        add_data_sheet(wb,"PEOPLE",["Person_ID","Person_Name","County","Program"],people)
        add_data_sheet(wb,"FACT",["Fact_ID","Person_ID","Date","Target","Actual","Cost_KES"],fact,[3])
    elif day in [41,44,45,49]:
        rows=[[f"E-{day:02d}-{i+1:04d}",random.choice(COUNTIES),random.choice(PROGRAMS),random.choice(["Intervention","Comparison"]),random.randint(20,160),random.randint(0,190),random.randint(500,15000),random.choice(["High","Medium","Low"])] for i in range(240)]
        src=["CASE_DATA"]
        add_data_sheet(wb,"CASE_DATA",["Record_ID","County","Program","Group","Target","Actual","Cost_KES","Risk"],rows)
    else:
        people=[[f"E-{i:04d}",person(),random.choice(COUNTIES),random.choice(PROGRAMS)] for i in range(1,181)]
        activity=[[f"A-{i+1:05d}",random.choice([r[0] for r in people]+["E-9999"]),BASE_DATE+timedelta(days=random.randint(0,400)),random.choice(ACTIVITIES),random.randint(1,30),random.randint(500,10000)] for i in range(420)]
        finance=[[f"T-{i+1:05d}",random.choice([r[0] for r in people]),BASE_DATE+timedelta(days=random.randint(0,400)),random.randint(100,18000),random.choice(["Approved","Pending","Rejected"])] for i in range(300)]
        src=["PEOPLE","ACTIVITY","FINANCE"]
        add_data_sheet(wb,"PEOPLE",["Entity_ID","Entity_Name","County","Program"],people)
        add_data_sheet(wb,"ACTIVITY",["Activity_ID","Entity_ID","Date","Activity_Type","Units","Cost_KES"],activity,[3])
        add_data_sheet(wb,"FINANCE",["Transaction_ID","Entity_ID","Date","Amount_KES","Approval"],finance,[3])
        if day in [47,48,50]:
            add_data_sheet(wb,"RULES",["Rule","Definition"],[["Engagement","Count activity only when Units >= 2"],["Expenditure","Use Approved finance only"],["Unmatched IDs","Must appear in exception report"],["Freshness","Show latest date from each source"]])
            src.append("RULES")
    return src


def build_dataset(d: dict) -> Path:
    wb=Workbook()
    wb.remove(wb.active)
    sources=build_sources(wb,d)
    add_start_here(wb,d,sources)
    path=DATASETS/f"Day_{d['day']:02d}_{d['level']}_Exercise_Dataset.xlsx"
    wb.save(path)
    return path

# PDF helpers
styles=getSampleStyleSheet()
styles.add(ParagraphStyle(name="KTitle", parent=styles["Title"], fontName="Helvetica-Bold", fontSize=18, leading=22, textColor=colors.HexColor("#17365D"), alignment=TA_CENTER, spaceAfter=8))
styles.add(ParagraphStyle(name="KSub", parent=styles["BodyText"], fontName="Helvetica-Bold", fontSize=10.5, leading=14, textColor=colors.HexColor("#2F75B5"), alignment=TA_CENTER, spaceAfter=10))
styles.add(ParagraphStyle(name="KH", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=11.5, leading=14, textColor=colors.HexColor("#17365D"), spaceBefore=6, spaceAfter=4))
styles.add(ParagraphStyle(name="KB", parent=styles["BodyText"], fontSize=9.4, leading=13, spaceAfter=4))
styles.add(ParagraphStyle(name="KSmall", parent=styles["BodyText"], fontSize=7.7, leading=9.5))


def footer(canvas, doc):
    canvas.saveState(); canvas.setFont("Helvetica",8); canvas.setFillColor(colors.HexColor("#6B7280")); canvas.drawString(15*mm,8*mm,"Kaldrix Limited · Data Analysis Mastery Programme"); canvas.drawRightString(195*mm,8*mm,f"Page {doc.page}"); canvas.restoreState()


def build_lesson_pdf(d: dict) -> Path:
    path=LESSONS/f"Day_{d['day']:02d}_{d['level']}.pdf"
    flow=[Paragraph(f"DAY {d['day']} - {d['level'].upper()}",styles["KTitle"]),Paragraph(d["title"],styles["KSub"]),Paragraph(f"<b>Primary tool:</b> {d['tool']}",styles["KB"]),Paragraph("What you are learning",styles["KH"]),Paragraph(d["concepts"],styles["KB"]),Paragraph("Learning objectives",styles["KH"])]
    for o in d["objectives"]: flow.append(Paragraph("- "+o,styles["KB"]))
    flow += [Paragraph("Worked practical",styles["KH"]),Paragraph(d["practical"],styles["KB"]),Spacer(1,4)]
    # compact visual table aligned with the lesson's analytical mindset
    visual=[["Source / Input","Method","Output / Check"],["Raw data","Apply the lesson skill","Result linked to source"],["Change one input","Recalculate / rerun","Confirm expected change"],["Control total / test","Validate","Explain the result"]]
    t=Table(visual,colWidths=[53*mm,58*mm,65*mm])
    t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#2F75B5")),("TEXTCOLOR",(0,0),(-1,0),colors.white),("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("GRID",(0,0),(-1,-1),0.35,colors.HexColor("#CBD5E1")),("VALIGN",(0,0),(-1,-1),"TOP"),("FONTSIZE",(0,0),(-1,-1),8),("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white,colors.HexColor("#F7FAFC")])]))
    flow += [t,Paragraph("Guided practice",styles["KH"]),Paragraph("1. Recreate the lesson idea on a small example.<br/>2. Change inputs and predict the result before recalculating.<br/>3. Complete the supplied independent dataset in the required tool.<br/>4. Create a control check or reconciliation.<br/>5. Explain the result in ordinary language.",styles["KB"]),Paragraph("AI coaching",styles["KH"]),Paragraph(d["ai_use"],styles["KB"]),Paragraph("AI discipline",styles["KH"]),Paragraph("Attempt first. Use AI for explanation, hints, debugging, critique, test cases and mock questioning. Never submit output you cannot explain and independently verify.",styles["KB"]),PageBreak(),Paragraph("Independent exercise",styles["KH"]),Paragraph(d["exercise"],styles["KB"]),Paragraph("Common failure modes",styles["KH"]),Paragraph("- Wrong grain or key.<br/>- Formula/query/code answers a different question.<br/>- Missing or unmatched records are silently ignored.<br/>- Findings are not reconciled to the numbers.<br/>- AI-generated logic is accepted without testing.",styles["KB"]),Paragraph("Mastery gate",styles["KH"]),Paragraph("[ ] I can reproduce the method without copying.<br/>[ ] I can trace the result to the source.<br/>[ ] I have a QA/control check.<br/>[ ] I can explain a limitation.<br/>[ ] I can answer an assessor's follow-up question without asking AI.",styles["KB"])]
    doc=SimpleDocTemplate(str(path),pagesize=A4,rightMargin=15*mm,leftMargin=15*mm,topMargin=13*mm,bottomMargin=14*mm)
    doc.build(flow,onFirstPage=footer,onLaterPages=footer)
    return path


def build_roadmap() -> Path:
    path=ROOT/"Data_Analysis_Mastery_Roadmap.pdf"
    flow=[Paragraph("KALDRIX DATA ANALYSIS MASTERY ROADMAP",styles["KTitle"]),Paragraph("10 weeks · 50 practical learning days · Beginner → Intermediate → Advanced → Pro → Expert",styles["KSub"]),Paragraph("The programme develops real analytical capability through practical data, trainer review, QA and progressively stronger tools.",styles["KB"])]
    for week in range(1,11):
        days=[d for d in COURSE if d["week"]==week]
        flow.append(Paragraph(f"Week {week} - {days[0]['level']}",styles["KH"]))
        rows=[["Day","Lesson","Tool","Exercise"]]+[[str(d["day"]),d["title"],d["tool"],d["exercise"]] for d in days]
        t=Table(rows,colWidths=[12*mm,65*mm,40*mm,72*mm],repeatRows=1)
        t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#2F75B5")),("TEXTCOLOR",(0,0),(-1,0),colors.white),("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("GRID",(0,0),(-1,-1),0.35,colors.HexColor("#CBD5E1")),("VALIGN",(0,0),(-1,-1),"TOP"),("FONTSIZE",(0,0),(-1,-1),7.2),("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white,colors.HexColor("#F8FAFC")])]))
        flow.append(t)
        if week<10: flow.append(PageBreak())
    doc=SimpleDocTemplate(str(path),pagesize=A4,rightMargin=12*mm,leftMargin=12*mm,topMargin=12*mm,bottomMargin=14*mm)
    doc.build(flow,onFirstPage=footer,onLaterPages=footer)
    return path


def build_trainer_guide() -> Path:
    path=ROOT/"Trainer_Assessor_Guide.pdf"
    flow=[Paragraph("TRAINER & ASSESSOR GUIDE",styles["KTitle"]),Paragraph("Kaldrix Data Analysis Mastery Programme",styles["KSub"]),Paragraph("Progress requires correct output, explanation, validation and independent reproduction. A submission alone is not mastery.",styles["KB"]),Paragraph("Recommended marking",styles["KH"])]
    rows=[["Area","Weight","Evidence"],["Accuracy","35%","Correct calculations/queries/measures/code and reconciled totals."],["Understanding","30%","Explains method, grain, keys, assumptions and limitations."],["QA","20%","Uses control totals, exceptions, tests and source-to-output tracing."],["Communication","15%","Clear findings, charts, documentation and decision relevance."]]
    t=Table(rows,colWidths=[35*mm,25*mm,118*mm]); t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#17365D")),("TEXTCOLOR",(0,0),(-1,0),colors.white),("GRID",(0,0),(-1,-1),0.35,colors.HexColor("#CBD5E1")),("VALIGN",(0,0),(-1,-1),"TOP"),("FONTSIZE",(0,0),(-1,-1),8)])); flow.append(t)
    flow += [Paragraph("AI-use rule",styles["KH"]),Paragraph("AI is encouraged for learning, debugging, critique, alternative approaches, test generation and mock defense. It is not accepted as an unexplained answer generator.",styles["KB"])]
    for level in ["Beginner","Intermediate","Advanced","Pro","Expert"]:
        flow += [PageBreak(),Paragraph(f"{level} mastery gate",styles["KH"])]
        for d in [x for x in COURSE if x["level"]==level]:
            flow.append(Paragraph(f"<b>Day {d['day']} - {d['title']}:</b> Ask the learner to trace one result to the source, explain the logic, show a QA check, and state a limitation.",styles["KB"]))
    doc=SimpleDocTemplate(str(path),pagesize=A4,rightMargin=16*mm,leftMargin=16*mm,topMargin=14*mm,bottomMargin=14*mm)
    doc.build(flow,onFirstPage=footer,onLaterPages=footer)
    return path


def as_data_url(path: Path, mime: str) -> str:
    return f"data:{mime};base64," + base64.b64encode(path.read_bytes()).decode("ascii")


def build_html(lesson_paths: dict[str,Path], dataset_paths: dict[str,Path], roadmap: Path):
    pdfs={k:as_data_url(v,"application/pdf") for k,v in lesson_paths.items()}
    data={k:as_data_url(v,"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") for k,v in dataset_paths.items()}
    master=as_data_url(roadmap,"application/pdf")
    common=(TEMPLATE.replace("__COURSE__",json.dumps(COURSE,ensure_ascii=False,separators=(",",":")))
                    .replace("__PDFS__",json.dumps(pdfs,separators=(",",":")))
                    .replace("__DATASETS__",json.dumps(data,separators=(",",":")))
                    .replace("__MASTER__",json.dumps(master))
                    .replace("__LOGO_DATA__",LOGO_DATA))
    learner=common.replace("__TRAINER_APP__","false")
    trainer=common.replace("__TRAINER_APP__","true")
    (ROOT/"Kaldrix_Data_Analysis_Mastery_Learner.html").write_text(learner,encoding="utf-8")
    (ROOT/"Kaldrix_Data_Analysis_Mastery_Trainer.html").write_text(trainer,encoding="utf-8")


def write_instructions():
    (ROOT/"HOW_TO_USE_KALDRIX_PROGRAMME.txt").write_text("""KALDRIX DATA ANALYSIS MASTERY PROGRAMME\n\nLEARNER\n1. Open Kaldrix_Data_Analysis_Mastery_Learner.html.\n2. Enter your official name.\n3. Study the lesson and PDF.\n4. Download the day-specific exercise dataset and import it into Google Sheets where Google Sheets is the required tool.\n5. Complete the work, share the Google Sheet and paste its URL in the app.\n6. Export the course record and send it to the trainer.\n\nTRAINER\n1. Open Kaldrix_Data_Analysis_Mastery_Trainer.html.\n2. Import the learner course record.\n3. Enter Trainer Mode, review the submission and approve or return it.\n4. Export the reviewed record back to the learner.\n\nAI LEARNING LAB\nThe AI Lab builds a lesson-aware coaching prompt. Copy & Open ChatGPT copies it using a clipboard fallback and opens ChatGPT. The learner pastes the prompt into ChatGPT and continues interactively. No API key is embedded in the HTML.\n\nCERTIFICATES\nA level certificate unlocks only after all 10 lessons in the level are trainer-approved and marked Done by the learner.\n""",encoding="utf-8")


def build_zip():
    path=ROOT/"Kaldrix_Data_Analysis_Mastery_Programme.zip"
    if path.exists(): path.unlink()
    with zipfile.ZipFile(path,"w",zipfile.ZIP_DEFLATED) as z:
        for name in ["Kaldrix_Data_Analysis_Mastery_Learner.html","Kaldrix_Data_Analysis_Mastery_Trainer.html","Data_Analysis_Mastery_Roadmap.pdf","Trainer_Assessor_Guide.pdf","HOW_TO_USE_KALDRIX_PROGRAMME.txt"]:
            z.write(ROOT/name,name)
        for p in sorted(LESSONS.glob("*.pdf")): z.write(p,f"lessons/{p.name}")
        for p in sorted(DATASETS.glob("*.xlsx")): z.write(p,f"day_datasets/{p.name}")
    return path


def main():
    # clean generated directories so obsolete artifacts do not linger
    for folder in [LESSONS,DATASETS]:
        for p in folder.iterdir():
            if p.is_file(): p.unlink()
    lesson_paths={}
    dataset_paths={}
    for d in COURSE:
        lesson_paths[str(d["day"])]=build_lesson_pdf(d)
        dataset_paths[str(d["day"])]=build_dataset(d)
    roadmap=build_roadmap()
    build_trainer_guide()
    write_instructions()
    build_html(lesson_paths,dataset_paths,roadmap)
    package=build_zip()
    print(f"Built {len(lesson_paths)} lessons, {len(dataset_paths)} datasets and {package.name}")


if __name__ == "__main__":
    main()
