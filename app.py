import io
import re
from collections import defaultdict

import pandas as pd
import streamlit as st

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    from docx import Document
except ImportError:
    Document = None


st.set_page_config(
    page_title="Civil Engineering Timetable Planner",
    page_icon="🗓️",
    layout="wide",
)

DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
HOURS = {
    "Monday": list(range(8, 15)),
    "Tuesday": list(range(8, 15)),
    "Wednesday": list(range(8, 15)),
    "Thursday": list(range(8, 15)),
    "Friday": list(range(8, 13)),
}
SECTIONS = ["A", "B", "C", "D"]

# Section structure used by the Civil Engineering allocation:
# 25CE is the only batch with four sections (A-D).
# All other CE batches have only three sections (A-C).
def sections_for_batch(batch):
    return ["A", "B", "C", "D"] if canon_batch(batch) == "25CE" else ["A", "B", "C"]


# -------------------------------------------------------------------
# General helpers
# -------------------------------------------------------------------
def clean(x):
    if x is None:
        return ""
    return re.sub(r"\s+", " ", str(x).replace("\xa0", " ").replace("\n", " ")).strip()


def canon_batch(x):
    return re.sub(r"[^0-9A-Z]", "", clean(x).upper())


def subject_code(x):
    m = re.search(r"\(([A-Z]{2,8}\d{2,5})\)", clean(x).upper())
    return m.group(1) if m else ""


def parse_subject_cell(x):
    s = clean(x)
    m = re.search(
        r"^(.*?)\s*\(([A-Z]{2,8}\d{2,5})\)\s*(\d+)\s*\+\s*(\d+)",
        s,
        re.I,
    )
    if not m:
        return None
    return {
        "subject": clean(m.group(1)),
        "code": m.group(2).upper(),
        "theory": int(m.group(3)),
        "practical": int(m.group(4)),
    }


def is_na_teacher(x):
    s = clean(x).lower()
    return s in {"", "n.a", "n.a.", "na", "-", "—", "none", "cell", "bsrs"}


def is_civil_row(row):
    # The supplied allocation document has a separate final block for
    # subjects taught in other departments. Those rows are never imported.
    if row.get("other_department"):
        return False
    return clean(row.get("code", "")).upper().startswith("CE")


def detect_batch(text):
    t = clean(text).upper()
    m = re.search(r"\b(2[2-6])\s*[- ]?\s*CE\s*(?:BATCH)?\b", t)
    return canon_batch(m.group(1) + "CE") if m else ""


def detect_semester(text):
    m = re.search(r"\b(\d+)\s*(?:ST|ND|RD|TH)\s+SEMESTER\b", clean(text).upper())
    return int(m.group(1)) if m else ""


# -------------------------------------------------------------------
# File readers
# -------------------------------------------------------------------
def read_pdf(upload):
    if pdfplumber is None:
        raise RuntimeError("pdfplumber is missing.")
    upload.seek(0)
    pages = []
    with pdfplumber.open(upload) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=2, y_tolerance=3) or ""
            tables = []
            for settings in (
                {"vertical_strategy": "lines", "horizontal_strategy": "lines"},
                {"vertical_strategy": "text", "horizontal_strategy": "text"},
            ):
                try:
                    tables = page.extract_tables(settings) or []
                except Exception:
                    tables = []
                if tables:
                    break
            pages.append({"text": text, "tables": tables})
    return pages


def read_docx(upload):
    if Document is None:
        raise RuntimeError("python-docx is missing.")
    upload.seek(0)
    doc = Document(upload)
    tables = []
    for table in doc.tables:
        tables.append([[clean(c.text) for c in row.cells] for row in table.rows])
    paragraphs = [clean(p.text) for p in doc.paragraphs if clean(p.text)]
    return tables, paragraphs


def read_excel(upload):
    upload.seek(0)
    return pd.read_excel(upload, sheet_name=None)


# -------------------------------------------------------------------
# Allocation parsing
# -------------------------------------------------------------------
def parse_structured_table(table, context_text=""):
    """Parse a normal Subject / Theory A-D / Practical table."""
    if not table:
        return []

    out = []
    header = [clean(v).lower() for v in table[0]]
    # Find likely subject column.
    subject_col = next(
        (i for i, h in enumerate(header) if "subject" in h),
        0,
    )

    batch = detect_batch(context_text)
    semester = detect_semester(context_text)

    for row in table[1:] if len(table) > 1 else []:
        vals = [clean(v) for v in row]
        joined = " | ".join(vals)

        if "SUBJECTS TO BE TAUGHT IN OTHER DEPARTMENTS" in joined.upper():
            break

        parsed = None
        parsed_idx = None
        for i, v in enumerate(vals):
            p = parse_subject_cell(v)
            if p:
                parsed = p
                parsed_idx = i
                break

        if not parsed:
            # Sometimes subject/code/CH are split across cells.
            p = parse_subject_cell(joined.replace(" | ", " "))
            if p:
                parsed = p
                parsed_idx = subject_col

        if not parsed:
            continue

        after = vals[(parsed_idx or 0) + 1 :]
        # Expected layout after subject: Theory A, B, C, optional D, Practical.
        while len(after) < 5:
            after.append("")

        practical = after[-1]
        theories = after[:-1][:4]

        out.append(
            {
                "batch": batch,
                "semester": semester,
                "subject": parsed["subject"],
                "code": parsed["code"],
                "theory_hours": parsed["theory"],
                "practical_hours": parsed["practical"],
                "theory_teachers": theories,
                "practical_teacher": practical,
                "other_department": False,
            }
        )
    return out


def parse_text_blocks(text):
    """
    Fallback parser for PDFs/DOCX where table extraction loses the merged
    cells. It uses batch headings and subject-code lines, then reads the
    following teacher lines conservatively.
    """
    lines = [clean(x) for x in text.splitlines() if clean(x)]
    rows = []
    current_batch = ""
    current_sem = ""
    current = None
    teacher_lines = []
    other_department = False

    def finish():
        nonlocal current, teacher_lines
        if not current:
            return
        t = [clean(x) for x in teacher_lines if clean(x)]
        while len(t) < 5:
            t.append("")
        rows.append(
            {
                "batch": current_batch,
                "semester": current_sem,
                "subject": current["subject"],
                "code": current["code"],
                "theory_hours": current["theory"],
                "practical_hours": current["practical"],
                "theory_teachers": t[:4],
                "practical_teacher": t[4],
                "other_department": other_department,
            }
        )
        current = None
        teacher_lines = []

    for line in lines:
        if "SUBJECTS TO BE TAUGHT IN OTHER DEPARTMENTS" in line.upper():
            finish()
            other_department = True
            continue
        if other_department:
            continue

        b = detect_batch(line)
        if b:
            finish()
            current_batch = b
            continue

        s = detect_semester(line)
        if s:
            current_sem = s
            continue

        p = parse_subject_cell(line)
        if p:
            finish()
            current = p
            teacher_lines = []
            continue

        if current:
            # Accept typical teacher-name lines, but ignore table labels.
            if re.search(r"\b(Prof|Dr|Engr|Mr|Ms)\b", line, re.I):
                teacher_lines.append(line)

    finish()
    return rows


def dedupe(rows):
    seen = set()
    result = []
    for r in rows:
        key = (
            canon_batch(r.get("batch")),
            clean(r.get("code")).upper(),
            clean(r.get("subject")).lower(),
            str(r.get("semester")),
        )
        if not key[1] or key in seen:
            continue
        seen.add(key)
        result.append(r)
    return result


def parse_pdf(upload):
    pages = read_pdf(upload)
    all_rows = []

    # First use page-level table extraction because the allocation document
    # has explicit Subject / Theory A-D / Practical columns.
    for page in pages:
        page_text = page["text"]
        for table in page["tables"]:
            all_rows.extend(parse_structured_table(table, page_text))

    all_rows = dedupe(all_rows)

    # If table extraction failed, use the text fallback.
    if not all_rows:
        all_rows = parse_text_blocks("\n".join(p["text"] for p in pages))

    return dedupe(all_rows)


def parse_docx(upload):
    tables, paragraphs = read_docx(upload)
    rows = []
    paragraph_text = "\n".join(paragraphs)

    for table in tables:
        rows.extend(parse_structured_table(table, paragraph_text))

    if not rows:
        rows = parse_text_blocks(paragraph_text)

    return dedupe(rows)


def parse_excel(upload):
    sheets = read_excel(upload)
    rows = []

    for sheet_name, df in sheets.items():
        df = df.fillna("")
        text = "\n".join(" ".join(clean(v) for v in row.tolist()) for _, row in df.iterrows())

        if "SUBJECTS TO BE TAUGHT IN OTHER DEPARTMENTS" in text.upper():
            text = text.split("SUBJECTS TO BE TAUGHT IN OTHER DEPARTMENTS", 1)[0]

        # Excel files vary considerably, so use row-level parsing.
        current_batch = ""
        current_sem = ""
        for _, row in df.iterrows():
            vals = [clean(v) for v in row.tolist()]
            joined = " | ".join(vals)

            b = detect_batch(joined)
            if b:
                current_batch = b
            s = detect_semester(joined)
            if s:
                current_sem = s

            parsed = None
            idx = None
            for i, v in enumerate(vals):
                p = parse_subject_cell(v)
                if p:
                    parsed, idx = p, i
                    break

            if not parsed:
                p = parse_subject_cell(joined.replace(" | ", " "))
                if p:
                    parsed, idx = p, 0

            if not parsed:
                continue

            after = vals[(idx or 0) + 1 :]
            while len(after) < 5:
                after.append("")
            rows.append(
                {
                    "batch": current_batch,
                    "semester": current_sem,
                    "subject": parsed["subject"],
                    "code": parsed["code"],
                    "theory_hours": parsed["theory"],
                    "practical_hours": parsed["practical"],
                    "theory_teachers": after[:-1][:4],
                    "practical_teacher": after[-1],
                    "other_department": False,
                }
            )

    return dedupe(rows)


def parse_upload(upload):
    name = upload.name.lower()
    if name.endswith(".pdf"):
        return parse_pdf(upload)
    if name.endswith(".docx"):
        return parse_docx(upload)
    if name.endswith((".xlsx", ".xls", ".xlsm")):
        return parse_excel(upload)
    raise ValueError("Use PDF, DOCX, XLSX, XLS or XLSM.")


# -------------------------------------------------------------------
# Convert allocation into section-specific teaching requirements
# -------------------------------------------------------------------
def expand_records(rows, selected_sections):
    records = []

    for r in rows:
        if not is_civil_row(r):
            continue

        batch = canon_batch(r["batch"])
        # IMPORTANT: The allocation's final column is the practical teacher;
        # it is NOT a fourth section. Only 25CE has Section D.
        valid_sections = sections_for_batch(batch)
        sections = [s for s in selected_sections if s in valid_sections]

        for sec in sections:
            idx = ["A", "B", "C", "D"].index(sec)
            theory_teachers = r.get("theory_teachers", [])
            teacher = clean(theory_teachers[idx]) if idx < len(theory_teachers) else ""
            practical_teacher = clean(r.get("practical_teacher", ""))

            # Theory: one-hour periods. A 3+1 subject needs 3 theory periods/week.
            # A blank Theory teacher means no theory class is created for that section.
            for n in range(int(r.get("theory_hours", 0) or 0)):
                if is_na_teacher(teacher):
                    continue
                records.append(
                    {
                        "batch": batch,
                        "section": sec,
                        "subject": r["subject"],
                        "code": r["code"],
                        "teacher": teacher,
                        "type": "Theory",
                        "period_no": n + 1,
                    }
                )

            # Practical: the teacher comes from the LAST/PRACTICAL column.
            # It is applied to each real section of the batch, but never to a
            # non-existent Section D for 22CE/23CE/24CE/26CE/etc.
            if int(r.get("practical_hours", 0) or 0) > 0 and not is_na_teacher(practical_teacher):
                records.append(
                    {
                        "batch": batch,
                        "section": sec,
                        "subject": r["subject"],
                        "code": r["code"],
                        "teacher": practical_teacher,
                        "type": "Practical",
                        "period_no": 1,
                    }
                )

    return records


# -------------------------------------------------------------------
# Timetable solver
# -------------------------------------------------------------------
def teacher_busy(schedule, teacher, day, start):
    if is_na_teacher(teacher):
        return False
    return any(
        x["teacher"] == teacher and x["day"] == day and x["start"] == start
        for x in schedule
    )


def section_busy(schedule, batch, section, day, start):
    return any(
        x["batch"] == batch
        and x["section"] == section
        and x["day"] == day
        and x["start"] == start
        for x in schedule
    )


def same_subject_same_day(schedule, batch, section, code, day):
    return any(
        x["batch"] == batch
        and x["section"] == section
        and x["code"] == code
        and x["day"] == day
        and x["type"] == "Theory"
        for x in schedule
    )


def day_load(schedule, batch, section, day):
    return sum(
        1
        for x in schedule
        if x["batch"] == batch and x["section"] == section and x["day"] == day
    )


def place_practical(schedule, r):
    candidates = []

    for day in DAYS:
        hs = HOURS[day]
        for i in range(len(hs) - 2):
            block = [hs[i], hs[i + 1], hs[i + 2]]

            # Friday has only 08-13.
            if any(section_busy(schedule, r["batch"], r["section"], day, h) for h in block):
                continue
            if any(teacher_busy(schedule, r["teacher"], day, h) for h in block):
                continue

            score = block[0]
            if day == "Friday":
                score += 80

            # Prefer a block adjacent to existing classes rather than
            # creating an isolated late period.
            existing = sorted(
                x["start"]
                for x in schedule
                if x["batch"] == r["batch"]
                and x["section"] == r["section"]
                and x["day"] == day
            )
            if existing:
                if block[0] == max(existing) + 1:
                    score -= 40
                if block[-1] == min(existing) - 1:
                    score -= 30

            candidates.append((score, day, block[0]))

    if not candidates:
        return False

    _, day, start = min(candidates)
    for h in [start, start + 1, start + 2]:
        schedule.append(
            {
                **r,
                "day": day,
                "start": h,
                "end": h + 1,
            }
        )
    return True


def place_theory(schedule, r):
    candidates = []

    for day in DAYS:
        for start in HOURS[day]:
            if section_busy(schedule, r["batch"], r["section"], day, start):
                continue
            if teacher_busy(schedule, r["teacher"], day, start):
                continue

            # Strong rule: no two theory periods of the same subject on one day.
            if same_subject_same_day(
                schedule, r["batch"], r["section"], r["code"], day
            ):
                continue

            score = start

            # Prefer not to use Friday unless necessary.
            if day == "Friday":
                score += 100

            # Prefer compact daily schedules and avoid holes.
            starts = sorted(
                x["start"]
                for x in schedule
                if x["batch"] == r["batch"]
                and x["section"] == r["section"]
                and x["day"] == day
            )
            if starts:
                lo, hi = min(starts), max(starts)
                if start == hi + 1 or start == lo - 1:
                    score -= 50
                elif lo < start < hi:
                    score += 300

            # Spread a subject across different days.
            used_days = {
                x["day"]
                for x in schedule
                if x["batch"] == r["batch"]
                and x["section"] == r["section"]
                and x["code"] == r["code"]
                and x["type"] == "Theory"
            }
            if day in used_days:
                score += 200

            candidates.append((score, day, start))

    if not candidates:
        return False

    _, day, start = min(candidates)
    schedule.append(
        {
            **r,
            "day": day,
            "start": start,
            "end": start + 1,
        }
    )
    return True


def generate_timetable(records, locks):
    schedule = []

    # Fixed manual entries are inserted first and remain unchanged.
    for lock in locks:
        schedule.append(lock)

    practicals = [r for r in records if r["type"] == "Practical"]
    theories = [r for r in records if r["type"] == "Theory"]

    # Practical first because the 3-hour continuity constraint is strongest.
    failed = []
    for r in practicals:
        if not place_practical(schedule, r):
            failed.append(r)

    # More constrained teachers first.
    theories.sort(key=lambda x: (is_na_teacher(x["teacher"]), x["batch"], x["section"], x["code"]))

    for r in theories:
        if not place_theory(schedule, r):
            failed.append(r)

    return schedule, failed


# -------------------------------------------------------------------
# Validation and output
# -------------------------------------------------------------------
def conflicts(schedule):
    out = []
    teacher_map = defaultdict(list)
    section_map = defaultdict(list)

    for x in schedule:
        if not is_na_teacher(x["teacher"]):
            teacher_map[(x["teacher"], x["day"], x["start"])].append(x)
        section_map[(x["batch"], x["section"], x["day"], x["start"])].append(x)

    for key, items in teacher_map.items():
        if len(items) > 1:
            out.append(("Teacher clash", key, items))

    for key, items in section_map.items():
        if len(items) > 1:
            out.append(("Section clash", key, items))

    return out


def time_label(start):
    return f"{start:02d}:00-{start + 1:02d}:00"


def timetable_df(schedule):
    return pd.DataFrame(
        [
            {
                "Batch": x["batch"],
                "Section": x["section"],
                "Day": x["day"],
                "Start": x["start"],
                "Time": time_label(x["start"]),
                "Type": x["type"],
                "Subject": f'{x["subject"]} ({x["code"]})',
                "Teacher": x["teacher"] or "Not specified",
            }
            for x in schedule
        ]
    )


def make_grid(df, batch, section):
    data = {"Time": [time_label(h) for h in range(8, 15)]}
    for day in DAYS:
        values = []
        for h in range(8, 15):
            if day == "Friday" and h >= 13:
                values.append("")
                continue

            z = df[
                (df["Batch"] == batch)
                & (df["Section"] == section)
                & (df["Day"] == day)
                & (df["Start"] == h)
            ]

            if z.empty:
                values.append("")
            else:
                r = z.iloc[0]
                prefix = "🔬 " if r["Type"] == "Practical" else ""
                values.append(f'{prefix}{r["Subject"]}\n{r["Teacher"]}')
        data[day] = values
    return pd.DataFrame(data)


def export_excel(df):
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="All Classes")
        for (batch, section), _ in df.groupby(["Batch", "Section"]):
            sheet = f"{batch}-{section}"[:31]
            make_grid(df, batch, section).to_excel(writer, index=False, sheet_name=sheet)
    return bio.getvalue()


# -------------------------------------------------------------------
# Streamlit UI
# -------------------------------------------------------------------
st.title("🗓️ Civil Engineering Timetable Planner")
st.caption(
    "Upload the subject-allocation sheet in PDF, Word or Excel. "
    "The app extracts Civil Engineering subjects, maps Theory A-D to Sections A-D, "
    "and generates a timetable with teacher-clash checking."
)

with st.sidebar:
    st.header("Rules")
    st.write("• Saturday & Sunday: OFF")
    st.write("• Monday–Thursday: 08:00–15:00")
    st.write("• Friday: 08:00–13:00")
    st.write("• Theory = 1-hour periods")
    st.write("• Practical = one continuous 3-hour block/week")
    st.write("• No two theory periods of the same subject on one day")
    st.write("• Manual locked periods are preserved")
    st.write("• Other-department subjects are excluded")
    st.write("• Section limit: A, B, C, D")

upload = st.file_uploader(
    "Upload Subject Allocation File",
    type=["pdf", "docx", "xlsx", "xls", "xlsm"],
)

if upload:
    with st.spinner("Reading and reconstructing the allocation table..."):
        try:
            raw = parse_upload(upload)
            raw = [r for r in raw if is_civil_row(r)]
        except Exception as e:
            raw = []
            st.error(f"Reading error: {e}")

    if not raw:
        st.error(
            "No Civil Engineering subject rows were detected. "
            "Check that the file contains subject codes such as CE411, CE407, etc."
        )
        st.stop()

    st.success(f"{len(raw)} Civil Engineering subject rows detected.")

    # Review/edit stage is important for complex PDF/Word layouts.
    review = pd.DataFrame(
        [
            {
                "Batch": r["batch"],
                "Semester": r["semester"],
                "Subject": r["subject"],
                "Code": r["code"],
                "Theory CH": r["theory_hours"],
                "Practical CH": r["practical_hours"],
                "Theory A": r["theory_teachers"][0] if len(r["theory_teachers"]) > 0 else "",
                "Theory B": r["theory_teachers"][1] if len(r["theory_teachers"]) > 1 else "",
                "Theory C": r["theory_teachers"][2] if len(r["theory_teachers"]) > 2 else "",
                "Theory D": r["theory_teachers"][3] if len(r["theory_teachers"]) > 3 else "",
                "Practical": r["practical_teacher"],
            }
            for r in raw
        ]
    )

    st.subheader("1. Extracted allocation — review before scheduling")
    st.info(
        "This review table is intentional. University allocation files use merged cells, "
        "multi-line teacher names and different Theory A/B/C/D layouts. You can correct "
        "any extraction issue here before generating the timetable."
    )
    review = st.data_editor(
        review,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        key="allocation_review",
    )

    batches = sorted(
        {
            canon_batch(x)
            for x in review["Batch"].tolist()
            if clean(x)
        }
    )

    st.subheader("2. Select batches and sections")
    c1, c2 = st.columns(2)
    selected_batches = c1.multiselect("Batches", batches, default=batches)

    available_sections = sorted(
        {sec for batch in selected_batches for sec in sections_for_batch(batch)},
        key=lambda x: ["A", "B", "C", "D"].index(x),
    )
    selected_sections = c2.multiselect(
        "Sections",
        available_sections,
        default=available_sections,
        help="A-C are available for all CE batches. Section D is available only for 25CE.",
    )

    st.subheader("3. Optional manual teacher/time locks")
    st.caption(
        "Use this when the timetable coordinator wants a particular teacher/class "
        "at a fixed time. Automatic scheduling will fill the remaining periods."
    )

    empty_locks = pd.DataFrame(
        columns=["Batch", "Section", "Day", "Start", "Subject", "Code", "Teacher", "Type"]
    )
    locks_df = st.data_editor(
        empty_locks,
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "Batch": st.column_config.SelectboxColumn("Batch", options=selected_batches),
            "Section": st.column_config.SelectboxColumn("Section", options=available_sections),
            "Day": st.column_config.SelectboxColumn("Day", options=DAYS),
            "Start": st.column_config.NumberColumn("Start hour", min_value=8, max_value=14, step=1),
            "Type": st.column_config.SelectboxColumn("Type", options=["Theory", "Practical"]),
        },
        key="manual_locks",
    )

    if st.button("⚙️ Generate Timetable", type="primary"):
        edited_rows = []
        for _, r in review.iterrows():
            edited_rows.append(
                {
                    "batch": canon_batch(r["Batch"]),
                    "semester": r["Semester"],
                    "subject": clean(r["Subject"]),
                    "code": clean(r["Code"]).upper(),
                    "theory_hours": int(r["Theory CH"] or 0),
                    "practical_hours": int(r["Practical CH"] or 0),
                    "theory_teachers": [
                        clean(r["Theory A"]),
                        clean(r["Theory B"]),
                        clean(r["Theory C"]),
                        clean(r["Theory D"]),
                    ],
                    "practical_teacher": clean(r["Practical"]),
                    "other_department": False,
                }
            )

        edited_rows = [
            r for r in edited_rows
            if r["batch"] in selected_batches and r["code"].startswith("CE")
        ]

        records = expand_records(edited_rows, selected_sections)

        locks = []
        for _, r in locks_df.iterrows():
            batch = canon_batch(r.get("Batch", ""))
            section = clean(r.get("Section", "")).upper()
            day = clean(r.get("Day", ""))
            subject = clean(r.get("Subject", "")) or "Manual class"
            code = clean(r.get("Code", "")).upper()
            teacher = clean(r.get("Teacher", ""))
            typ = clean(r.get("Type", "Theory")) or "Theory"
            if not batch or not section or not day:
                continue

            start = int(r.get("Start", 8) or 8)
            length = 3 if typ == "Practical" else 1
            if start + length > (13 if day == "Friday" else 15):
                st.warning(f"Manual lock {batch}-{section} {day} {start}:00 is outside the allowed time.")
                continue

            for j in range(length):
                locks.append(
                    {
                        "batch": batch,
                        "section": section,
                        "subject": subject,
                        "code": code,
                        "teacher": teacher,
                        "type": typ,
                        "period_no": 1,
                        "day": day,
                        "start": start + j,
                        "end": start + j + 1,
                    }
                )

        schedule, failed = generate_timetable(records, locks)
        st.session_state["schedule"] = schedule
        st.session_state["failed"] = failed

if "schedule" in st.session_state:
    schedule = st.session_state["schedule"]
    failed = st.session_state.get("failed", [])

    st.subheader("4. Generated timetable")
    df = timetable_df(schedule)

    if failed:
        st.warning(
            f"{len(failed)} required period(s) could not be placed. "
            "This means the current constraints leave no legal slot for those classes."
        )
        with st.expander("Show unplaced classes"):
            st.dataframe(
                pd.DataFrame(
                    [
                        {
                            "Batch": x["batch"],
                            "Section": x["section"],
                            "Subject": x["subject"],
                            "Code": x["code"],
                            "Type": x["type"],
                            "Teacher": x["teacher"],
                        }
                        for x in failed
                    ]
                ),
                use_container_width=True,
            )

    if not df.empty:
        for batch in sorted(df["Batch"].unique()):
            for section in sorted(df[df["Batch"] == batch]["Section"].unique()):
                st.markdown(f"### {batch} — Section {section}")
                st.dataframe(
                    make_grid(df, batch, section),
                    use_container_width=True,
                    hide_index=True,
                )

        st.download_button(
            "⬇️ Download Complete Timetable (Excel)",
            data=export_excel(df),
            file_name="Civil_Engineering_Timetable.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        bad = conflicts(schedule)
        if bad:
            st.error(f"{len(bad)} clash(es) detected.")
            st.dataframe(
                pd.DataFrame(
                    [
                        {"Type": a, "Key": str(b), "Classes": str(c)}
                        for a, b, c in bad
                    ]
                ),
                use_container_width=True,
            )
        else:
            st.success("✅ No teacher or batch/section time clashes detected.")
