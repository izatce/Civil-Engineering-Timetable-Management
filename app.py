
import io
from collections import defaultdict
from datetime import datetime

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.shared import Inches, Pt

st.set_page_config(
    page_title="Civil Engineering Department, MUET Jamshoro - Timetable Maker",
    page_icon="📅",
    layout="wide",
)

DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
DAY_HOURS = {
    "Monday": list(range(8, 15)),
    "Tuesday": list(range(8, 15)),
    "Wednesday": list(range(8, 15)),
    "Thursday": list(range(8, 15)),
    "Friday": list(range(8, 13)),
}
BATCH_SUGGESTIONS = ["22CE", "23CE", "24CE", "25CE", "26CE"]
SECTION_SUGGESTIONS = ["A", "B", "C", "D"]


def allowed_sections(batch):
    # Kept only as suggestions. The coordinator is free to enter
    # any batch and section manually.
    return SECTION_SUGGESTIONS


def hour_label(h):
    return f"{h:02d}:00-{h+1:02d}:00"


def empty_allocations():
    return pd.DataFrame(
        columns=[
            "Batch", "Section", "Subject", "Code",
            "Teacher", "Type", "CH"
        ]
    )


def empty_locks():
    return pd.DataFrame(
        columns=[
            "Batch", "Section", "Day", "Start",
            "Subject", "Code", "Teacher", "Type"
        ]
    )


def normalize_allocations(df):
    if df is None or df.empty:
        return empty_allocations()

    out = df.copy()
    for c in empty_allocations().columns:
        if c not in out.columns:
            out[c] = ""

    out = out[empty_allocations().columns].copy()
    out["Batch"] = out["Batch"].astype(str).str.strip().str.upper()
    out["Section"] = out["Section"].astype(str).str.strip().str.upper()
    out["Subject"] = out["Subject"].astype(str).str.strip()
    out["Code"] = out["Code"].astype(str).str.strip().str.upper()
    out["Teacher"] = out["Teacher"].astype(str).str.strip()
    out["Type"] = out["Type"].astype(str).str.strip().str.title()
    out["CH"] = pd.to_numeric(out["CH"], errors="coerce").fillna(0).astype(int)

    return out[
        out["Batch"].ne("")
        & out["Section"].ne("")
        & out["Subject"].ne("")
        & out["Teacher"].ne("")
        & out["Type"].isin(["Theory", "Practical"])
        & out["CH"].gt(0)
    ].reset_index(drop=True)


def validate_allocations(df):
    errors = []
    if df.empty:
        return errors

    for i, r in df.iterrows():
        row = i + 1
        if not r["Batch"]:
            errors.append(f"Row {row}: batch is missing.")
        if not r["Section"]:
            errors.append(f"Row {row}: section is missing.")
        if not r["Subject"]:
            errors.append(f"Row {row}: subject is missing.")
        if not r["Teacher"]:
            errors.append(f"Row {row}: teacher is missing.")
        if r["Type"] not in ["Theory", "Practical"]:
            errors.append(f"Row {row}: Type must be Theory or Practical.")
        if int(r["CH"]) <= 0:
            errors.append(f"Row {row}: CH must be greater than zero.")

        if r["Type"] == "Practical" and int(r["CH"]) != 1:
            errors.append(
                f"Row {row}: Practical CH should be 1 under the current timetable rule."
            )
        if r["Type"] == "Theory" and int(r["CH"]) > 3:
            errors.append(
                f"Row {row}: Theory CH above 3 may require a special rule."
            )
    return errors


def records_from_allocations(df):
    records = []
    for _, r in df.iterrows():
        base = {
            "batch": r["Batch"],
            "section": r["Section"],
            "subject": r["Subject"],
            "code": r["Code"],
            "teacher": r["Teacher"],
            "type": r["Type"],
        }

        if r["Type"] == "Practical":
            records.append({**base, "length": 3})
        else:
            for n in range(int(r["CH"])):
                records.append({**base, "length": 1, "period_no": n + 1})

    return records


def occupied(schedule, batch, section, day, start):
    return [
        x for x in schedule
        if x["batch"] == batch
        and x["section"] == section
        and x["day"] == day
        and x["start"] == start
    ]


def teacher_occupied(schedule, teacher, day, start):
    return [
        x for x in schedule
        if x["teacher"].strip().lower() == teacher.strip().lower()
        and x["day"] == day
        and x["start"] == start
    ]


def subject_on_day(schedule, batch, section, code, subject, day):
    return [
        x for x in schedule
        if x["batch"] == batch
        and x["section"] == section
        and x["day"] == day
        and x["code"].strip().upper() == code.strip().upper()
        and x["subject"].strip().lower() == subject.strip().lower()
        and x["type"] == "Theory"
    ]


def compactness_score(schedule, batch, section, day, start, length):
    starts = sorted(
        x["start"]
        for x in schedule
        if x["batch"] == batch
        and x["section"] == section
        and x["day"] == day
    )
    if not starts:
        score = 0
    else:
        lo, hi = min(starts), max(starts)
        score = 0
        if start == hi + 1:
            score -= 80
        elif start + length - 1 == lo - 1:
            score -= 60
        elif start > lo and start <= hi:
            score += 500

    # Prefer earlier completion; Friday has less available time.
    score += start * 2
    if day == "Friday":
        score += 40
    return score


def place_practical(schedule, r):
    candidates = []

    for day in DAYS:
        hours = DAY_HOURS[day]
        for i in range(len(hours) - 2):
            start = hours[i]
            block = [start, start + 1, start + 2]

            if any(
                occupied(schedule, r["batch"], r["section"], day, h)
                for h in block
            ):
                continue
            if any(teacher_occupied(schedule, r["teacher"], day, h) for h in block):
                continue

            score = compactness_score(
                schedule, r["batch"], r["section"], day, start, 3
            )
            candidates.append((score, day, start))

    if not candidates:
        return False

    _, day, start = min(candidates)

    for j in range(3):
        schedule.append({
            **r,
            "day": day,
            "start": start + j,
            "end": start + j + 1,
            "block_id": f"{r['batch']}-{r['section']}-{r['code']}-PR-{day}-{start}",
        })
    return True


def place_theory(schedule, r):
    candidates = []

    for day in DAYS:
        for start in DAY_HOURS[day]:
            if occupied(schedule, r["batch"], r["section"], day, start):
                continue
            if teacher_occupied(schedule, r["teacher"], day, start):
                continue

            # Never put the same theory subject twice on the same day.
            if subject_on_day(
                schedule, r["batch"], r["section"],
                r["code"], r["subject"], day
            ):
                continue

            score = compactness_score(
                schedule, r["batch"], r["section"], day, start, 1
            )

            # Spread theory classes through the week.
            used_days = {
                x["day"]
                for x in schedule
                if x["batch"] == r["batch"]
                and x["section"] == r["section"]
                and x["code"].strip().upper() == r["code"].strip().upper()
                and x["type"] == "Theory"
            }
            if day in used_days:
                score += 300
            else:
                score -= 30

            candidates.append((score, day, start))

    if not candidates:
        return False

    _, day, start = min(candidates)

    schedule.append({
        **r,
        "day": day,
        "start": start,
        "end": start + 1,
        "block_id": f"{r['batch']}-{r['section']}-{r['code']}-TH-{day}-{start}",
    })
    return True


def create_manual_locks(lock_df):
    locks = []
    errors = []

    if lock_df is None or lock_df.empty:
        return locks, errors

    for i, r in lock_df.iterrows():
        row = i + 1
        batch = str(r.get("Batch", "")).strip().upper()
        section = str(r.get("Section", "")).strip().upper()
        day = str(r.get("Day", "")).strip()
        subject = str(r.get("Subject", "")).strip()
        code = str(r.get("Code", "")).strip().upper()
        teacher = str(r.get("Teacher", "")).strip()
        typ = str(r.get("Type", "")).strip().title()

        try:
            start = int(r.get("Start", 0))
        except Exception:
            start = 0

        if not batch:
            errors.append(f"Lock row {row}: batch is required.")
            continue
        if not section:
            errors.append(f"Lock row {row}: section is required.")
            continue
        if day not in DAYS:
            errors.append(f"Lock row {row}: invalid day.")
            continue
        if typ not in ["Theory", "Practical"]:
            errors.append(f"Lock row {row}: Type must be Theory or Practical.")
            continue
        if start not in DAY_HOURS[day]:
            errors.append(f"Lock row {row}: invalid start time for {day}.")
            continue
        if not teacher:
            errors.append(f"Lock row {row}: teacher is required.")
            continue

        length = 3 if typ == "Practical" else 1
        if any(h not in DAY_HOURS[day] for h in range(start, start + length)):
            errors.append(
                f"Lock row {row}: {typ} does not fit completely on {day}."
            )
            continue

        for j in range(length):
            locks.append({
                "batch": batch,
                "section": section,
                "subject": subject or "Manual class",
                "code": code,
                "teacher": teacher,
                "type": typ,
                "day": day,
                "start": start + j,
                "end": start + j + 1,
                "block_id": f"LOCK-{row}",
                "locked": True,
            })

    return locks, errors


def validate_locks(locks):
    errors = []
    seen_section = set()
    seen_teacher = set()

    for x in locks:
        s_key = (x["batch"], x["section"], x["day"], x["start"])
        t_key = (x["teacher"].lower(), x["day"], x["start"])

        if s_key in seen_section:
            errors.append(
                f"Manual clash: {x['batch']}-{x['section']} has two fixed classes "
                f"at {x['day']} {hour_label(x['start'])}."
            )
        seen_section.add(s_key)

        if t_key in seen_teacher:
            errors.append(
                f"Manual teacher clash: {x['teacher']} has two fixed classes "
                f"at {x['day']} {hour_label(x['start'])}."
            )
        seen_teacher.add(t_key)

    return errors


def generate(records, locks):
    schedule = list(locks)
    failed = []

    # Hardest constraints first.
    practicals = [r for r in records if r["type"] == "Practical"]
    theories = [r for r in records if r["type"] == "Theory"]

    for r in practicals:
        if not place_practical(schedule, r):
            failed.append(r)

    # Put subjects with more hours first.
    theories.sort(key=lambda x: (-int(x.get("length", 1)), x["batch"], x["section"]))

    for r in theories:
        if not place_theory(schedule, r):
            failed.append(r)

    return schedule, failed


def clash_report(schedule):
    clashes = []
    section_seen = defaultdict(list)
    teacher_seen = defaultdict(list)

    for x in schedule:
        section_seen[
            (x["batch"], x["section"], x["day"], x["start"])
        ].append(x)
        teacher_seen[
            (x["teacher"].strip().lower(), x["day"], x["start"])
        ].append(x)

    for key, items in section_seen.items():
        if len(items) > 1:
            clashes.append(("Section clash", key, items))

    for key, items in teacher_seen.items():
        if len(items) > 1:
            clashes.append(("Teacher clash", key, items))

    return clashes


def schedule_dataframe(schedule):
    rows = []
    for x in schedule:
        rows.append({
            "Batch": x["batch"],
            "Section": x["section"],
            "Day": x["day"],
            "Start": x["start"],
            "Time": hour_label(x["start"]),
            "Type": x["type"],
            "Subject": x["subject"],
            "Code": x["code"],
            "Teacher": x["teacher"],
            "Fixed": "Yes" if x.get("locked") else "No",
        })
    return pd.DataFrame(rows)


def timetable_grid(schedule, batch, section):
    rows = []
    for h in range(8, 15):
        row = {"Time": hour_label(h)}
        for day in DAYS:
            if day == "Friday" and h >= 13:
                row[day] = ""
                continue

            items = [
                x for x in schedule
                if x["batch"] == batch
                and x["section"] == section
                and x["day"] == day
                and x["start"] == h
            ]
            if items:
                x = items[0]
                prefix = "P: " if x["type"] == "Practical" else ""
                row[day] = f"{prefix}{x['subject']}\n{x['teacher']}"
            else:
                row[day] = ""
        rows.append(row)
    return pd.DataFrame(rows)


def teacher_grid(schedule, teacher):
    rows = []
    for h in range(8, 15):
        row = {"Time": hour_label(h)}
        for day in DAYS:
            if day == "Friday" and h >= 13:
                row[day] = ""
                continue
            items = [
                x for x in schedule
                if x["teacher"].strip().lower() == teacher.strip().lower()
                and x["day"] == day
                and x["start"] == h
            ]
            row[day] = (
                f"{items[0]['batch']}-{items[0]['section']}\n{items[0]['subject']}"
                if items else ""
            )
        rows.append(row)
    return pd.DataFrame(rows)


def make_excel(schedule):
    wb = Workbook()
    ws = wb.active
    ws.title = "Master Timetable"

    title = "Civil Engineering Timetable"
    ws["A1"] = title
    ws["A1"].font = Font(size=16, bold=True)
    ws.merge_cells("A1:J1")

    df = schedule_dataframe(schedule)
    for c, col in enumerate(df.columns, 1):
        cell = ws.cell(3, c, col)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9EAF7")
        cell.alignment = Alignment(horizontal="center")

    for r_idx, row in enumerate(df.itertuples(index=False), 4):
        for c_idx, value in enumerate(row, 1):
            ws.cell(r_idx, c_idx, value)

    thin = Side(style="thin", color="B7B7B7")
    for row in ws.iter_rows():
        for cell in row:
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            cell.alignment = Alignment(vertical="center", wrap_text=True)

    for i in range(1, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(i)].width = 18

    # Batch/section sheets
    groups = sorted({(x["batch"], x["section"]) for x in schedule})
    for batch, section in groups:
        ws2 = wb.create_sheet(f"{batch}-{section}")
        ws2["A1"] = f"{batch} - Section {section}"
        ws2["A1"].font = Font(size=14, bold=True)
        ws2.merge_cells("A1:F1")

        grid = timetable_grid(schedule, batch, section)
        for c, col in enumerate(grid.columns, 1):
            cell = ws2.cell(3, c, col)
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="D9EAF7")
            cell.alignment = Alignment(horizontal="center")

        for rr, row in enumerate(grid.itertuples(index=False), 4):
            for cc, value in enumerate(row, 1):
                ws2.cell(rr, cc, value)
                ws2.cell(rr, cc).alignment = Alignment(
                    horizontal="center", vertical="center", wrap_text=True
                )

        for c in range(1, ws2.max_column + 1):
            ws2.column_dimensions[get_column_letter(c)].width = 24

    # Teacher sheets
    teachers = sorted({x["teacher"] for x in schedule if x["teacher"]})
    for teacher in teachers:
        safe = "".join(ch for ch in teacher if ch.isalnum() or ch in " -_")[:25]
        ws3 = wb.create_sheet(("T-" + safe)[:31])
        ws3["A1"] = f"Teacher Timetable: {teacher}"
        ws3["A1"].font = Font(size=13, bold=True)
        ws3.merge_cells("A1:F1")
        grid = teacher_grid(schedule, teacher)

        for c, col in enumerate(grid.columns, 1):
            ws3.cell(3, c, col).font = Font(bold=True)
            ws3.cell(3, c).fill = PatternFill("solid", fgColor="E2F0D9")

        for rr, row in enumerate(grid.itertuples(index=False), 4):
            for cc, value in enumerate(row, 1):
                ws3.cell(rr, cc, value)
                ws3.cell(rr, cc).alignment = Alignment(
                    horizontal="center", vertical="center", wrap_text=True
                )

        for c in range(1, ws3.max_column + 1):
            ws3.column_dimensions[get_column_letter(c)].width = 24

    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


def make_word(schedule):
    doc = Document()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("DEPARTMENT OF CIVIL ENGINEERING")
    run.bold = True
    run.font.size = Pt(16)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Class Timetable")
    run.bold = True
    run.font.size = Pt(14)

    groups = sorted({(x["batch"], x["section"]) for x in schedule})

    for batch, section in groups:
        doc.add_heading(f"{batch} — Section {section}", level=2)
        grid = timetable_grid(schedule, batch, section)

        table = doc.add_table(rows=1, cols=len(grid.columns))
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.style = "Table Grid"

        for i, col in enumerate(grid.columns):
            cell = table.rows[0].cells[i]
            cell.text = col
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            for run in cell.paragraphs[0].runs:
                run.bold = True

        for _, r in grid.iterrows():
            cells = table.add_row().cells
            for i, col in enumerate(grid.columns):
                cells[i].text = str(r[col])
                cells[i].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

        doc.add_paragraph("")

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()


# -------------------------------------------------------------------
# UI
# -------------------------------------------------------------------
st.title("📅 Civil Engineering Department, MUET Jamshoro")
st.subheader("Timetable Maker")
st.caption(
    "Manual subject allocation → optional fixed slots → automatic clash-free timetable → Excel/Word"
)

if "allocations" not in st.session_state:
    st.session_state.allocations = empty_allocations()

if "locks" not in st.session_state:
    st.session_state.locks = empty_locks()

if "schedule" not in st.session_state:
    st.session_state.schedule = None

if "failed" not in st.session_state:
    st.session_state.failed = []

tab1, tab2, tab3, tab4 = st.tabs([
    "1️⃣ Subject Allocation",
    "2️⃣ Fixed Time Slots",
    "3️⃣ Generate Timetable",
    "4️⃣ Results & Download",
])

with tab1:
    st.markdown("### Enter Subject Allocation")

    st.info(
        "Enter one row for each teaching assignment. "
        "For a subject with theory and practical, use separate rows. "
        "Batch and section are open-entry fields. Type the exact batch/section used by your department (e.g. 22CE-D)."
    )

    c1, c2, c3 = st.columns(3)
    batch = c1.text_input(
        "Batch",
        placeholder="e.g. 22CE, 23CE, 27CE, MSc-1, etc.",
        help="Open entry: type any batch name/code. It is not limited to a predefined list."
    )
    section = c2.text_input(
        "Section",
        placeholder="e.g. A, B, C, D, E",
        help="Open entry: type any section label required by your department."
    )
    typ = c3.selectbox("Type", ["Theory", "Practical"])

    st.caption(
        "Common examples: " + ", ".join(BATCH_SUGGESTIONS) +
        " — these are only examples; you can enter any batch."
    )

    c1, c2, c3, c4 = st.columns(4)
    subject = c1.text_input("Subject")
    code = c2.text_input("Subject Code")
    teacher = c3.text_input("Subject Teacher")
    ch = c4.number_input(
        "Credit Hours",
        min_value=1,
        max_value=3,
        value=3 if typ == "Theory" else 1,
        step=1,
    )

    if st.button("➕ Add Allocation", type="primary"):
        new_row = pd.DataFrame([{
            "Batch": batch,
            "Section": section,
            "Subject": subject.strip(),
            "Code": code.strip().upper(),
            "Teacher": teacher.strip(),
            "Type": typ,
            "CH": int(ch),
        }])

        if not subject.strip() or not teacher.strip():
            st.error("Subject and teacher are required.")
        elif typ == "Practical" and int(ch) != 1:
            st.error("Practical credit hours must be 1.")
        else:
            st.session_state.allocations = pd.concat(
                [st.session_state.allocations, new_row],
                ignore_index=True,
            )
            st.success("Allocation added.")

    st.markdown("#### Current Saved Allocation")
    edited = st.data_editor(
        st.session_state.allocations,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        column_config={
            "Batch": st.column_config.TextColumn("Batch"),
            "Section": st.column_config.TextColumn("Section"),
            "Type": st.column_config.SelectboxColumn(
                "Type", options=["Theory", "Practical"]
            ),
            "CH": st.column_config.NumberColumn("CH", min_value=1, max_value=3),
        },
        key="allocation_editor",
    )

    if st.button("💾 Save Allocation Table"):
        clean_df = normalize_allocations(edited)
        errors = validate_allocations(edited)
        if errors:
            st.error("Please correct these entries:")
            for e in errors:
                st.write("•", e)
        else:
            st.session_state.allocations = clean_df
            st.success(
                f"Saved {len(clean_df)} allocation rows. "
                "The saved table is now the source used by the scheduler."
            )

    if not st.session_state.allocations.empty:
        if st.button("🗑️ Clear All Allocation"):
            st.session_state.allocations = empty_allocations()
            st.session_state.schedule = None
            st.rerun()

with tab2:
    st.markdown("### Fixed / Manual Time Slots")
    st.write(
        "Optional: enter classes that the coordinator wants to place at a specific "
        "day/time. These slots are locked and the automatic scheduler works around them."
    )

    fixed_batch = st.text_input(
        "Batch for fixed slot",
        placeholder="e.g. 22CE",
        key="fixed_batch",
    )
    fixed_section = st.text_input(
        "Section for fixed slot",
        placeholder="e.g. A",
        key="fixed_section",
    )
    if not fixed_batch:
        st.caption("Common examples: " + ", ".join(BATCH_SUGGESTIONS) + " — custom batches are also allowed.")

    c1, c2, c3 = st.columns(3)
    fixed_day = c1.selectbox("Day", DAYS)
    fixed_start = c2.selectbox(
        "Start Time",
        DAY_HOURS[fixed_day],
        format_func=lambda x: hour_label(x),
    )
    fixed_type = c3.selectbox("Type", ["Theory", "Practical"], key="fixed_type")

    c1, c2, c3 = st.columns(3)
    fixed_subject = c1.text_input("Subject", key="fixed_subject")
    fixed_code = c2.text_input("Code", key="fixed_code")
    fixed_teacher = c3.text_input("Teacher", key="fixed_teacher")

    if st.button("🔒 Add Fixed Slot"):
        length = 3 if fixed_type == "Practical" else 1
        allowed = DAY_HOURS[fixed_day]
        if not all(h in allowed for h in range(fixed_start, fixed_start + length)):
            st.error("The practical block does not fit completely in the selected day.")
        elif not fixed_teacher.strip():
            st.error("Teacher is required.")
        else:
            row = pd.DataFrame([{
                "Batch": fixed_batch,
                "Section": fixed_section,
                "Day": fixed_day,
                "Start": fixed_start,
                "Subject": fixed_subject,
                "Code": fixed_code.upper(),
                "Teacher": fixed_teacher,
                "Type": fixed_type,
            }])
            st.session_state.locks = pd.concat(
                [st.session_state.locks, row],
                ignore_index=True,
            )
            st.success("Fixed slot added.")

    st.data_editor(
        st.session_state.locks,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        key="lock_editor",
    )

    if st.button("💾 Save Fixed Slots"):
        st.session_state.locks = st.session_state.lock_editor
        st.success("Fixed slots saved.")

with tab3:
    st.markdown("### Generate Timetable")

    if st.session_state.allocations.empty:
        st.warning("First enter and save the subject allocation.")
    else:
        st.write(
            f"**Saved allocation rows:** {len(st.session_state.allocations)}"
        )
        st.write(
            "The scheduler will use the saved allocation and fixed slots. "
            "Saturday and Sunday are automatically excluded."
        )

        if st.button("🚀 Generate Timetable", type="primary"):
            alloc = normalize_allocations(st.session_state.allocations)
            errors = validate_allocations(alloc)

            lock_source = st.session_state.locks.copy()
            locks, lock_errors = create_manual_locks(lock_source)
            lock_errors += validate_locks(locks)

            if errors or lock_errors:
                if errors:
                    st.error("Allocation errors:")
                    for e in errors:
                        st.write("•", e)
                if lock_errors:
                    st.error("Fixed-slot errors:")
                    for e in lock_errors:
                        st.write("•", e)
            else:
                records = records_from_allocations(alloc)

                # Remove records whose teacher is missing after editing.
                records = [r for r in records if r["teacher"].strip()]

                schedule, failed = generate(records, locks)
                st.session_state.schedule = schedule
                st.session_state.failed = failed

                st.success(
                    f"Timetable generated: {len(schedule)} scheduled periods."
                )
                if failed:
                    st.warning(
                        f"{len(failed)} required allocation(s)/period(s) could not be placed. "
                        "See the Results tab for details."
                    )

with tab4:
    st.markdown("### Results")

    schedule = st.session_state.schedule

    if schedule is None:
        st.info("Generate the timetable first.")
    else:
        df = schedule_dataframe(schedule)

        clashes = clash_report(schedule)
        if clashes:
            st.error(f"{len(clashes)} clash(es) detected.")
            for typ, key, items in clashes[:20]:
                st.write(f"• **{typ}:** {key}")
        else:
            st.success("✅ No teacher or section clashes detected.")

        if st.session_state.failed:
            st.warning(
                f"{len(st.session_state.failed)} required class period(s) could not be scheduled."
            )
            failed_df = pd.DataFrame([
                {
                    "Batch": x["batch"],
                    "Section": x["section"],
                    "Subject": x["subject"],
                    "Code": x["code"],
                    "Teacher": x["teacher"],
                    "Type": x["type"],
                }
                for x in st.session_state.failed
            ])
            st.dataframe(failed_df, use_container_width=True)

        st.markdown("### Complete Schedule")
        st.dataframe(
            df.sort_values(["Batch", "Section", "Day", "Start"]),
            use_container_width=True,
            hide_index=True,
        )

        st.markdown("### Timetable by Batch / Section")
        result_batches = sorted({x["batch"] for x in schedule})
        for batch in result_batches:
            result_sections = sorted({x["section"] for x in schedule if x["batch"] == batch})
            for section in result_sections:
                st.markdown(f"#### {batch} — Section {section}")
                st.dataframe(
                    timetable_grid(schedule, batch, section),
                    use_container_width=True,
                    hide_index=True,
                )

        st.markdown("### Downloads")
        excel_data = make_excel(schedule)
        word_data = make_word(schedule)

        c1, c2 = st.columns(2)
        c1.download_button(
            "⬇️ Download Excel",
            data=excel_data,
            file_name="Civil_Engineering_Timetable.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        c2.download_button(
            "⬇️ Download Word",
            data=word_data,
            file_name="Civil_Engineering_Timetable.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
