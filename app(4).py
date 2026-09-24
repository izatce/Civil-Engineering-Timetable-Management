import io
import re
from itertools import combinations

import pandas as pd
import streamlit as st

# Optional document readers
try:
    import pdfplumber
except Exception:
    pdfplumber = None

try:
    from docx import Document
except Exception:
    Document = None


st.set_page_config(
    page_title="Civil Engineering Timetable Generator",
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
TIME_LABEL = {h: f"{h:02d}:00–{h+1:02d}:00" for h in range(8, 15)}


# ============================================================
# BASIC HELPERS
# ============================================================

def clean(x):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    return str(x).strip()


def norm(x):
    return re.sub(r"\s+", " ", clean(x)).strip().lower()


def is_na(x):
    return norm(x) in {"", "n.a", "n/a", "na", "none", "-", "—"}


def split_teachers(x):
    s = clean(x)
    if is_na(s) or norm(s) in {"bsrs", "cell"}:
        return []
    parts = re.split(r"\s*/\s*|\s+\band\b\s+", s, flags=re.I)
    return [p.strip() for p in parts if p.strip()]


def parse_ch(x):
    s = clean(x).replace(" ", "")
    m = re.search(r"(\d+(?:\.\d+)?)\+(\d+(?:\.\d+)?)", s)
    if m:
        return int(float(m.group(1))), int(float(m.group(2)))
    m = re.search(r"(\d+(?:\.\d+)?)", s)
    return (int(float(m.group(1))), 0) if m else (0, 0)


def find_col(columns, names):
    for wanted in names:
        for col in columns:
            if norm(col) == norm(wanted):
                return col
    for col in columns:
        for wanted in names:
            if norm(wanted) in norm(col):
                return col
    return None


def looks_like_teacher(text):
    s = norm(text)
    return any(
        key in s
        for key in ["prof.", "prof ", "dr.", "dr ", "engr.", "engr ", "eng."]
    )


# ============================================================
# EXCEL
# ============================================================

def parse_excel(uploaded):
    xls = pd.ExcelFile(uploaded)
    records = []

    for sheet in xls.sheet_names:
        raw = pd.read_excel(xls, sheet_name=sheet, header=None)

        cutoff = len(raw)
        for i in range(len(raw)):
            row_text = " ".join(norm(v) for v in raw.iloc[i].tolist())
            if "subjects to be taught in other departments" in row_text:
                cutoff = i
                break

        raw = raw.iloc[:cutoff]

        current_batch = ""
        current_semester = ""

        for i in range(len(raw)):
            vals = [clean(v) for v in raw.iloc[i].tolist()]
            text = " ".join(v for v in vals if v)

            # Examples: 22CE, 23CE, 24CE, 25CE, 26CE
            b = re.search(
                r"\b(2[0-9]\s*CE)\b|\b(2[0-9])[- ]?BATCH\b",
                text,
                re.I,
            )
            if b:
                current_batch = (
                    (b.group(1) or f"{b.group(2)}CE")
                    .replace(" ", "")
                    .upper()
                )
                continue

            sem = re.search(
                r"(\d+(?:st|nd|rd|th)\s+Semester)",
                text,
                re.I,
            )
            if sem:
                current_semester = sem.group(1)
                continue

            code_match = re.search(
                r"\(([A-Z]{2,8}\s*\d{3})\)",
                text,
            )
            if not code_match:
                continue

            code = re.sub(r"\s+", "", code_match.group(1)).upper()

            # Only Civil Engineering course codes are considered.
            if not code.startswith("CE"):
                continue

            subject = text[:code_match.start()].strip(" -:")
            subject = re.sub(r"\s+", " ", subject)

            if not subject:
                continue

            after = text[code_match.end():]
            ch_match = re.search(
                r"(\d+(?:\.\d+)?\s*\+\s*\d+(?:\.\d+)?)",
                after,
            )
            ch = ch_match.group(1).replace(" ", "") if ch_match else ""

            # Teacher information may be in following rows in the allocation.
            following = []
            for j in range(i + 1, min(i + 9, len(raw))):
                next_text = " ".join(
                    clean(v) for v in raw.iloc[j].tolist() if clean(v)
                )

                if re.search(r"\([A-Z]{2,8}\s*\d{3}\)", next_text):
                    break

                if next_text:
                    following.append(next_text)

            teacher_values = [
                v for v in vals
                if looks_like_teacher(v)
            ]

            teacher_values = teacher_values or following

            teacher_values = [
                v for v in teacher_values
                if not any(
                    k in norm(v)
                    for k in [
                        "subject",
                        "theory",
                        "practical",
                        "semester",
                        "department",
                        "batch",
                    ]
                )
            ]

            while len(teacher_values) < 5:
                teacher_values.append("")

            records.append(
                {
                    "Batch": current_batch,
                    "Semester": current_semester,
                    "Subject": subject,
                    "Code": code,
                    "CH": ch,
                    "Theory A": teacher_values[0],
                    "Theory B": teacher_values[1],
                    "Theory C": teacher_values[2],
                    "Theory D": teacher_values[3],
                    "Practical": teacher_values[4],
                    "Source": sheet,
                }
            )

    if not records:
        raise ValueError(
            "No Civil Engineering allocation was detected in the Excel file."
        )

    return pd.DataFrame(records).drop_duplicates(
        ["Batch", "Code", "Subject"]
    )


# ============================================================
# PDF
# ============================================================

def extract_pdf_text(uploaded):
    if pdfplumber is None:
        raise RuntimeError(
            "pdfplumber is not installed. Add it to requirements.txt."
        )

    uploaded.seek(0)
    pages = []

    with pdfplumber.open(uploaded) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if text.strip():
                pages.append(text)

    return "\n".join(pages)


# ============================================================
# WORD
# ============================================================

def extract_docx_text(uploaded):
    if Document is None:
        raise RuntimeError(
            "python-docx is not installed. Add it to requirements.txt."
        )

    uploaded.seek(0)
    document = Document(uploaded)

    chunks = []

    for paragraph in document.paragraphs:
        if paragraph.text.strip():
            chunks.append(paragraph.text)

    for table in document.tables:
        for row in table.rows:
            chunks.append(
                " | ".join(cell.text.strip() for cell in row.cells)
            )

    return "\n".join(chunks)


# ============================================================
# GENERIC PDF/WORD TEXT PARSER
# ============================================================

def parse_text_allocation(text):
    """
    Generic parser for text extracted from PDF/DOCX.

    It looks for Civil Engineering course codes such as CE411,
    course names, CH values such as 3+1, and nearby teacher names.

    PDF/Word layouts can vary considerably, so the extracted
    records are displayed for coordinator verification before
    timetable generation.
    """

    lines = [
        re.sub(r"\s+", " ", line).strip()
        for line in text.splitlines()
        if line.strip()
    ]

    # Ignore explicit other-department section.
    filtered = []
    other_dept = False

    for line in lines:
        if "subjects to be taught in other departments" in norm(line):
            other_dept = True
            continue

        if other_dept:
            continue

        filtered.append(line)

    records = []
    batch = ""
    semester = ""

    for i, line in enumerate(filtered):
        b = re.search(
            r"\b(2[0-9]\s*CE)\b|\b(2[0-9])[- ]?BATCH\b",
            line,
            re.I,
        )
        if b:
            batch = (
                (b.group(1) or f"{b.group(2)}CE")
                .replace(" ", "")
                .upper()
            )

        s = re.search(
            r"(\d+(?:st|nd|rd|th)\s+Semester)",
            line,
            re.I,
        )
        if s:
            semester = s.group(1)

        cm = re.search(r"\b(CE\d{3})\b", line, re.I)
        if not cm:
            continue

        code = cm.group(1).upper()

        # Avoid treating headings as subjects.
        before = line[:cm.start()].strip(" |-:")
        subject = re.sub(r"\s+", " ", before)

        if not subject or norm(subject) in {
            "subject",
            "course",
            "course title",
        }:
            continue

        ch_match = re.search(
            r"(\d+(?:\.\d+)?\s*\+\s*\d+(?:\.\d+)?)",
            line[cm.end():],
        )
        ch = ch_match.group(1).replace(" ", "") if ch_match else ""

        nearby = []
        for j in range(i + 1, min(i + 8, len(filtered))):
            candidate = filtered[j]

            if re.search(r"\bCE\d{3}\b", candidate, re.I):
                break

            if looks_like_teacher(candidate):
                nearby.append(candidate)

        while len(nearby) < 5:
            nearby.append("")

        records.append(
            {
                "Batch": batch,
                "Semester": semester,
                "Subject": subject,
                "Code": code,
                "CH": ch,
                "Theory A": nearby[0],
                "Theory B": nearby[1],
                "Theory C": nearby[2],
                "Theory D": nearby[3],
                "Practical": nearby[4],
                "Source": "PDF/Word",
            }
        )

    if not records:
        raise ValueError(
            "No CE course records were detected in the PDF/Word file. "
            "If it is a scanned image, OCR support will be needed."
        )

    return pd.DataFrame(records).drop_duplicates(
        ["Batch", "Code", "Subject"]
    )


# ============================================================
# FILE DISPATCHER
# ============================================================

def parse_uploaded_file(uploaded):
    name = uploaded.name.lower()

    if name.endswith((".xlsx", ".xls")):
        return parse_excel(uploaded), "Excel"

    if name.endswith(".pdf"):
        text = extract_pdf_text(uploaded)
        if not text.strip():
            raise ValueError(
                "No selectable text was found in this PDF. "
                "It may be a scanned/image PDF and needs OCR."
            )
        return parse_text_allocation(text), "PDF"

    if name.endswith(".docx"):
        text = extract_docx_text(uploaded)
        if not text.strip():
            raise ValueError("No readable text/table was found in the Word file.")
        return parse_text_allocation(text), "Word"

    if name.endswith(".doc"):
        raise ValueError(
            "Old .doc format is not directly supported. "
            "Please save it as .docx and upload again."
        )

    raise ValueError("Supported formats: .xlsx, .xls, .pdf and .docx")


# ============================================================
# SCHEDULING
# ============================================================

def expand_courses(df):
    units = []

    for _, r in df.iterrows():
        theory_ch, practical_ch = parse_ch(r.get("CH"))

        for col in ["Theory A", "Theory B", "Theory C", "Theory D"]:
            for teacher in split_teachers(r.get(col)):
                if theory_ch > 0:
                    units.append(
                        {
                            "Batch": clean(r.get("Batch")),
                            "Section": "",
                            "Subject": clean(r.get("Subject")),
                            "Code": clean(r.get("Code")),
                            "Type": "Theory",
                            "Teacher": teacher,
                            "Hours": theory_ch,
                            "Duration": 1,
                            "Group": col,
                        }
                    )

        for teacher in split_teachers(r.get("Practical")):
            if practical_ch > 0:
                units.append(
                    {
                        "Batch": clean(r.get("Batch")),
                        "Section": "",
                        "Subject": clean(r.get("Subject")),
                        "Code": clean(r.get("Code")),
                        "Type": "Practical",
                        "Teacher": teacher,
                        "Hours": 1,
                        "Duration": 3,
                        "Group": "Practical",
                    }
                )

    return pd.DataFrame(units)


def make_slots(day, start, duration):
    candidate = list(range(start, start + duration))
    if all(h in DAY_HOURS[day] for h in candidate):
        return [(day, h) for h in candidate]
    return []


def is_free(assignments, unit, candidate_slots):
    for a in assignments:
        if set(candidate_slots).isdisjoint(a["Slots"]):
            continue

        if norm(a["Teacher"]) == norm(unit["Teacher"]):
            return False

        if (
            norm(a["Batch"]) == norm(unit["Batch"])
            and norm(a["Section"]) == norm(unit["Section"])
        ):
            return False

    return True


def same_subject_day(assignments, unit, day):
    return sum(
        a["Day"] == day
        and norm(a["Batch"]) == norm(unit["Batch"])
        and norm(a["Section"]) == norm(unit["Section"])
        and norm(a["Subject"]) == norm(unit["Subject"])
        and a["Type"] == "Theory"
        for a in assignments
    )


def candidate_positions(unit, assignments):
    candidates = []

    for day in DAYS:
        for start in DAY_HOURS[day]:
            ss = make_slots(day, start, unit["Duration"])

            if len(ss) != unit["Duration"]:
                continue

            if not is_free(assignments, unit, ss):
                continue

            # A theory subject should not be scheduled twice
            # for the same batch/section on the same day.
            if unit["Type"] == "Theory" and same_subject_day(
                assignments, unit, day
            ):
                continue

            batch_load = sum(
                len(a["Slots"])
                for a in assignments
                if a["Day"] == day
                and norm(a["Batch"]) == norm(unit["Batch"])
                and norm(a["Section"]) == norm(unit["Section"])
            )

            teacher_load = sum(
                len(a["Slots"])
                for a in assignments
                if a["Day"] == day
                and norm(a["Teacher"]) == norm(unit["Teacher"])
            )

            score = batch_load * 5 + teacher_load * 3 + start

            # Avoid Friday practicals where possible.
            if day == "Friday" and unit["Type"] == "Practical":
                score += 30

            candidates.append((score, day, start, ss))

    return sorted(candidates)


def place_manuals(manual_df):
    assignments = []
    errors = []

    for _, r in manual_df.iterrows():
        unit = r.to_dict()
        ss = make_slots(
            r["Day"],
            int(r["Start Hour"]),
            int(r["Duration"]),
        )

        if len(ss) != int(r["Duration"]):
            errors.append(
                f'Invalid slot: {r["Batch"]}-{r["Section"]} '
                f'{r["Subject"]} on {r["Day"]}.'
            )
            continue

        if not is_free(assignments, unit, ss):
            errors.append(
                f'Conflict: {r["Batch"]}-{r["Section"]} '
                f'{r["Subject"]} / {r["Teacher"]}.'
            )
            continue

        assignments.append(
            {
                **unit,
                "Slots": ss,
                "Locked": True,
            }
        )

    return assignments, errors


def generate_timetable(units, manual):
    assignments = list(manual)
    unscheduled = []

    work = units.copy()
    work["_priority"] = work["Type"].map(
        {"Practical": 0, "Theory": 1}
    )

    work = work.sort_values(
        ["_priority", "Duration"],
        ascending=[True, False],
    )

    for _, row in work.iterrows():
        unit = row.to_dict()

        # Don't duplicate a manually fixed matching allocation.
        if any(
            norm(a["Batch"]) == norm(unit["Batch"])
            and norm(a["Section"]) == norm(unit["Section"])
            and norm(a["Subject"]) == norm(unit["Subject"])
            and norm(a["Teacher"]) == norm(unit["Teacher"])
            and a["Type"] == unit["Type"]
            for a in assignments
        ):
            continue

        repetitions = (
            int(unit["Hours"])
            if unit["Type"] == "Theory"
            else 1
        )

        for occurrence in range(repetitions):
            candidates = candidate_positions(unit, assignments)

            if not candidates:
                unscheduled.append(
                    {
                        **unit,
                        "Occurrence": occurrence + 1,
                        "Reason": (
                            "No clash-free slot available during "
                            "normal working hours."
                        ),
                    }
                )
                continue

            _, day, start, ss = candidates[0]

            assignments.append(
                {
                    **unit,
                    "Day": day,
                    "Start": start,
                    "Slots": ss,
                    "Locked": False,
                    "Occurrence": occurrence + 1,
                }
            )

    return assignments, pd.DataFrame(unscheduled)


def clash_report(assignments):
    report = []

    for a, b in combinations(assignments, 2):
        overlap = set(a["Slots"]) & set(b["Slots"])

        if not overlap:
            continue

        if norm(a["Teacher"]) == norm(b["Teacher"]):
            report.append(
                {
                    "Type": "Teacher clash",
                    "Teacher": a["Teacher"],
                    "Class A": a["Subject"],
                    "Class B": b["Subject"],
                }
            )

        if (
            norm(a["Batch"]) == norm(b["Batch"])
            and norm(a["Section"]) == norm(b["Section"])
        ):
            report.append(
                {
                    "Type": "Batch/section clash",
                    "Teacher": f'{a["Batch"]}-{a["Section"]}',
                    "Class A": a["Subject"],
                    "Class B": b["Subject"],
                }
            )

    return pd.DataFrame(report).drop_duplicates()


def export_excel(assignments):
    rows = []

    for a in assignments:
        rows.append(
            {
                "Batch": a["Batch"],
                "Section": a["Section"],
                "Subject": a["Subject"],
                "Code": a["Code"],
                "Type": a["Type"],
                "Teacher": a["Teacher"],
                "Day": a["Day"],
                "Start": TIME_LABEL[a["Start"]],
                "Duration": a["Duration"],
                "Locked": "Yes" if a["Locked"] else "No",
            }
        )

    bio = io.BytesIO()

    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        pd.DataFrame(rows).to_excel(
            writer,
            index=False,
            sheet_name="Assignments",
        )

    return bio.getvalue()


# ============================================================
# STREAMLIT UI
# ============================================================

st.title("📅 Civil Engineering Timetable Generator")

st.caption(
    "Upload the Civil Engineering subject allocation in Excel, PDF or Word format."
)

with st.sidebar:
    st.header("Working Rules")
    st.write("**Monday–Thursday:** 08:00–15:00")
    st.write("**Friday:** 08:00–13:00")
    st.write("**Saturday/Sunday:** OFF")
    st.write("**Theory:** 1-hour periods")
    st.write("**Practical:** 3 continuous hours")
    st.write("**Theory subject:** maximum one class/day")
    st.write("**Other departments:** excluded")

if "allocation" not in st.session_state:
    st.session_state.allocation = None

if "units" not in st.session_state:
    st.session_state.units = None

if "manual" not in st.session_state:
    st.session_state.manual = pd.DataFrame(
        columns=[
            "Batch",
            "Section",
            "Subject",
            "Code",
            "Teacher",
            "Type",
            "Day",
            "Start Hour",
            "Duration",
        ]
    )

if "assignments" not in st.session_state:
    st.session_state.assignments = []

if "unscheduled" not in st.session_state:
    st.session_state.unscheduled = pd.DataFrame()

if "clashes" not in st.session_state:
    st.session_state.clashes = pd.DataFrame()


tab1, tab2, tab3, tab4 = st.tabs(
    [
        "1. Upload Allocation",
        "2. Manual Fixed Slots",
        "3. Generate",
        "4. Results",
    ]
)


# ============================================================
# TAB 1
# ============================================================

with tab1:
    st.subheader("Upload Subject Allocation")

    uploaded = st.file_uploader(
        "Supported formats: Excel, PDF, Word",
        type=["xlsx", "xls", "pdf", "docx"],
        help=(
            "Excel is preferred. PDF and Word are also supported. "
            "Scanned/image PDFs require OCR and may not be readable."
        ),
    )

    if uploaded and st.button(
        "Read Civil Engineering Allocation",
        type="primary",
    ):
        try:
            allocation, source_type = parse_uploaded_file(uploaded)

            st.session_state.allocation = allocation

            st.success(
                f"{source_type} file read successfully. "
                f"{len(allocation)} Civil Engineering course records detected."
            )

        except Exception as e:
            st.error(str(e))

    if st.session_state.allocation is not None:
        st.write("### Detected Civil Engineering Allocation")

        st.dataframe(
            st.session_state.allocation,
            use_container_width=True,
            hide_index=True,
        )

        st.info(
            "Please verify this table before generating the timetable. "
            "PDF/Word layouts vary, so verification is especially important "
            "for those formats."
        )

        if st.button("Prepare Scheduling Units"):
            try:
                units = expand_courses(
                    st.session_state.allocation
                )

                st.session_state.units = units

                st.success(
                    f"{len(units)} scheduling units prepared."
                )

            except Exception as e:
                st.error(str(e))

    if st.session_state.units is not None:
        st.write("### Scheduling Units")

        st.dataframe(
            st.session_state.units,
            use_container_width=True,
            hide_index=True,
        )


# ============================================================
# TAB 2
# ============================================================

with tab2:
    st.subheader("Coordinator – Manual / Locked Slots")

    st.info(
        "A manually entered slot is locked first. "
        "The automatic scheduler will avoid that teacher and "
        "batch/section at that time."
    )

    allocation = st.session_state.allocation

    if allocation is not None:
        batches = sorted(
            x for x in allocation["Batch"].dropna().unique()
            if clean(x)
        )

        subjects = sorted(
            x for x in allocation["Subject"].dropna().unique()
            if clean(x)
        )

    else:
        batches = []
        subjects = []

    c1, c2, c3 = st.columns(3)

    with c1:
        batch = st.selectbox(
            "Batch",
            batches or [""],
            key="manual_batch",
        )

        section = st.text_input(
            "Section",
            value="A",
            key="manual_section",
        )

        subject = st.selectbox(
            "Subject",
            subjects or [""],
            key="manual_subject",
        )

    with c2:
        class_type = st.selectbox(
            "Type",
            ["Theory", "Practical"],
            key="manual_type",
        )

        teacher = st.text_input(
            "Teacher",
            key="manual_teacher",
        )

        day = st.selectbox(
            "Day",
            DAYS,
            key="manual_day",
        )

    with c3:
        start = st.selectbox(
            "Start Time",
            DAY_HOURS[day],
            format_func=lambda h: TIME_LABEL[h],
            key="manual_start",
        )

        duration = (
            3 if class_type == "Practical" else 1
        )

        st.write(
            f"Duration: **{duration} hour(s)**"
        )

    if st.button(
        "Add Locked Slot",
        type="primary",
    ):
        if not batch or not subject or not teacher:
            st.error(
                "Batch, Subject and Teacher are required."
            )

        else:
            new_row = pd.DataFrame(
                [
                    {
                        "Batch": batch,
                        "Section": section,
                        "Subject": subject,
                        "Code": "",
                        "Teacher": teacher,
                        "Type": class_type,
                        "Day": day,
                        "Start Hour": start,
                        "Duration": duration,
                    }
                ]
            )

            st.session_state.manual = pd.concat(
                [
                    st.session_state.manual,
                    new_row,
                ],
                ignore_index=True,
            )

            st.success("Locked slot added.")

    st.write("### Current Locked Slots")

    st.dataframe(
        st.session_state.manual,
        use_container_width=True,
        hide_index=True,
    )

    if st.button("Clear All Locked Slots"):
        st.session_state.manual = (
            st.session_state.manual.iloc[0:0]
        )
        st.rerun()


# ============================================================
# TAB 3
# ============================================================

with tab3:
    st.subheader("Generate Timetable")

    if st.session_state.units is None:
        st.warning(
            "Upload and prepare the allocation first."
        )

    else:
        st.write(
            f"Scheduling units: **{len(st.session_state.units)}**"
        )

        st.write(
            f"Locked slots: **{len(st.session_state.manual)}**"
        )

        if st.button(
            "🚀 Generate Timetable",
            type="primary",
        ):
            manual, manual_errors = place_manuals(
                st.session_state.manual
            )

            for error in manual_errors:
                st.error(error)

            assignments, unscheduled = generate_timetable(
                st.session_state.units,
                manual,
            )

            st.session_state.assignments = assignments
            st.session_state.unscheduled = unscheduled
            st.session_state.clashes = clash_report(
                assignments
            )

            if len(st.session_state.clashes) == 0:
                st.success(
                    "No teacher or batch/section clashes detected."
                )
            else:
                st.warning(
                    f"{len(st.session_state.clashes)} clash(es) detected."
                )

            if len(unscheduled):
                st.warning(
                    f"{len(unscheduled)} scheduling unit(s) "
                    "could not be placed."
                )


# ============================================================
# TAB 4
# ============================================================

with tab4:
    st.subheader("Generated Timetable")

    assignments = st.session_state.assignments

    if not assignments:
        st.info(
            "Generate the timetable first."
        )

    else:
        batches = sorted(
            set(
                clean(a["Batch"])
                for a in assignments
                if clean(a["Batch"])
            )
        )

        selected_batch = st.selectbox(
            "Batch",
            batches,
            key="result_batch",
        )

        sections = sorted(
            set(
                clean(a["Section"])
                for a in assignments
                if clean(a["Batch"]) == selected_batch
            )
        )

        selected_section = st.selectbox(
            "Section",
            sections or [""],
            key="result_section",
        )

        timetable_rows = []

        for hour in range(8, 15):
            row = {
                "Time": TIME_LABEL[hour]
            }

            for day in DAYS:
                values = []

                for a in assignments:
                    if (
                        a["Batch"] == selected_batch
                        and a["Section"] == selected_section
                        and a["Day"] == day
                        and hour in [
                            x[1]
                            for x in a["Slots"]
                        ]
                    ):
                        lock = " 🔒" if a["Locked"] else ""

                        values.append(
                            f'{a["Subject"]}{lock}'
                        )

                row[day] = " | ".join(values)

            timetable_rows.append(row)

        st.dataframe(
            pd.DataFrame(timetable_rows),
            use_container_width=True,
            hide_index=True,
        )

        st.subheader("Clash Report")

        if len(st.session_state.clashes):
            st.error(
                "Clashes detected."
            )

            st.dataframe(
                st.session_state.clashes,
                use_container_width=True,
                hide_index=True,
            )

        else:
            st.success(
                "No teacher or batch/section clashes detected."
            )

        st.subheader("Unscheduled Items")

        if len(st.session_state.unscheduled):
            st.warning(
                "These items could not be placed within "
                "the normal working hours."
            )

            st.dataframe(
                st.session_state.unscheduled,
                use_container_width=True,
                hide_index=True,
            )

        else:
            st.success(
                "All scheduling units were placed."
            )

        st.download_button(
            "⬇️ Download Timetable Excel",
            data=export_excel(assignments),
            file_name="Civil_Engineering_Timetable.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )
