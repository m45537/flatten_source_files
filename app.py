import os
# Force Streamlit to use polling watcher to avoid inotify limit on Streamlit Cloud
os.environ["STREAMLIT_SERVER_FILE_WATCHER_TYPE"] = "poll"

import io, re
import pandas as pd
import streamlit as st
from datetime import datetime
import pytz


# ------------------------------
# App config
# ------------------------------
st.set_page_config(page_title="Roster → Master (LENIENT)", page_icon="📘", layout="centered")
st.title("📘 Master Students Builder — Lenient")
st.caption("Upload Blackbaud, Rediker, Student Records → get a styled Excel with Master + Summary tabs.")

# ------------------------------
# Helpers (normalization & parsing)
# ------------------------------
def norm_piece(s: str) -> str:
    return re.sub(r"[^A-Z0-9 ]+", "", str(s).upper()).strip()

def grade_norm(s: str) -> str:
    x = norm_piece(s)
    x = re.sub(r"\s+", "", x)
    aliases = {
        "P4": "PK4", "PK": "PK4", "PREK": "PK4", "PREK4": "PK4", "PRE-K": "PK4", "PRE-K4": "PK4",
        "P3": "PK3", "PREK3": "PK3", "PRE-K3": "PK3",
        "KINDERGARTEN": "K", "KINDER": "K", "KG": "K"
    }
    if x in aliases:
        return aliases[x]
    m = re.fullmatch(r"(GRADE|GR|G)?(\d{1,2})", x)
    if m:
        return str(int(m.group(2)))
    return x

def surname_first_token(last: str) -> str:
    tokens = [t for t in norm_piece(last).split() if t]
    return tokens[0] if tokens else ""

def firstname_first_token(first: str, last: str) -> str:
    ftoks = [t for t in norm_piece(first).split() if t]
    if ftoks:
        return ftoks[0]
    ltoks = [t for t in norm_piece(last).split() if t]
    return ltoks[1] if len(ltoks) >= 2 else (ltoks[0] if ltoks else "")

def make_unique_key_lenient(first: str, last: str, grade: str) -> str:
    return f"{surname_first_token(last)}|{firstname_first_token(first, last)}|{grade_norm(grade)}"

# ---------- Blackbaud parser ----------
def parse_blackbaud(file) -> pd.DataFrame:
    df = pd.read_excel(file, usecols=list(range(5)))
    df.columns = [str(c).strip().upper() for c in df.columns]
    col_fam  = next((c for c in df.columns if "FAMILY" in c and "ID" in c), None)
    col_pf   = next((c for c in df.columns if "FIRST" in c and "PARENT" in c), None)
    col_pl   = next((c for c in df.columns if "LAST" in c and "PARENT" in c), None)
    col_stu  = next((c for c in df.columns if "STUDENT" in c and "GRADE" in c), None)
    if not all([col_fam, col_pf, col_pl, col_stu]):
        raise ValueError("Blackbaud: missing expected columns A–E")

    def split_students(cell: str):
        if pd.isna(cell) or str(cell).strip() == "":
            return []
        text = str(cell)
        text = re.sub(r"\s*\)\s*[,/;|]?\s*", ")|", text)
        parts = [p.strip().rstrip(",;/|") for p in text.split("|") if p.strip()]
        return parts

    def parse_student_entry(entry: str):
        m = re.search(r"\(([^)]+)\)\s*$", entry)
        grade = m.group(1).strip() if m else ""
        name = re.sub(r"\([^)]+\)\s*$", "", entry).strip()
        if ";" in name:
            last, first = [t.strip() for t in name.split(";", 1)]
        elif "," in name:
            last, first = [t.strip() for t in name.split(",", 1)]
        else:
            toks = name.split()
            last, first = (" ".join(toks[:-1]), toks[-1]) if len(toks) >= 2 else (name, "")
        return last, first, grade

    rows = []
    for _, r in df.iterrows():
        fam = str(r.get(col_fam, "")).replace(".0","").strip()
        pf  = str(r.get(col_pf, "")).strip()
        pl  = str(r.get(col_pl, "")).strip()
        for entry in split_students(r.get(col_stu, "")):
            stu_last, stu_first, grade = parse_student_entry(entry)
            rows.append({
                "ID": "",
                "FAMILY ID": fam,
                "PARENT FIRST NAME": pf,
                "PARENT LAST NAME": pl,
                "STUDENT FIRST NAME": stu_first,
                "STUDENT LAST NAME": stu_last,
                "GRADE": grade,
                "REDIKER ID": "",
                "SOURCE": "BB",
                "UNIQUE_KEY": make_unique_key_lenient(stu_first, stu_last, grade),
            })
    return pd.DataFrame(rows)

# ---------- Rediker parser ----------
def parse_rediker(file) -> pd.DataFrame:
    df = pd.read_excel(file, usecols="A:K").fillna("")
    df.columns = [str(c).strip().upper() for c in df.columns]
    col_first = next((c for c in df.columns if "FIRST" in c and "STUDENT" not in c), None)
    col_last  = next((c for c in df.columns if "LAST" in c and "STUDENT" not in c), None)
    col_grade = next((c for c in df.columns if "GRADE" in c), None)
    col_fam   = next((c for c in df.columns if "FAMILY" in c and "ID" in c), None)
    col_red   = next((c for c in df.columns if "APID" in c or "UNIQUE" in c or "REDIKER" in c or c=="ID"), None)

    rows = []
    for _, r in df.iterrows():
        fam   = str(r.get(col_fam, "")).replace(".0","").strip() if col_fam else ""
        rid   = str(r.get(col_red, "")).replace(".0","").strip() if col_red else ""
        first = str(r.get(col_first, "")).strip() if col_first else ""
        last  = str(r.get(col_last, "")).strip() if col_last else ""
        grade = str(r.get(col_grade, "")).strip() if col_grade else ""
        rows.append({
            "ID": "",
            "FAMILY ID": fam,
            "PARENT FIRST NAME": "",
            "PARENT LAST NAME": "",
            "STUDENT FIRST NAME": first,
            "STUDENT LAST NAME": last,
            "GRADE": grade,
            "REDIKER ID": rid,
            "SOURCE": "RED",
            "UNIQUE_KEY": make_unique_key_lenient(first, last, grade),
        })
    return pd.DataFrame(rows)

# ---------- Student Records parser ----------
def parse_student_records(file) -> pd.DataFrame:
    df = pd.read_excel(file).fillna("")
    df.columns = [str(c).strip().upper() for c in df.columns]
    col_id = df.columns[0]
    col_fam = next((c for c in df.columns if "FAMILY" in c and "ID" in c), None)
    col_red = next((c for c in df.columns if "REDIKER" in c), None)
    col_pf  = next((c for c in df.columns if "PARENT" in c and "FIRST" in c), None)
    col_pl  = next((c for c in df.columns if "PARENT" in c and "LAST" in c), None)
    col_sf  = next((c for c in df.columns if "STUDENT" in c and "FIRST" in c), None)
    col_sl  = next((c for c in df.columns if "STUDENT" in c and "LAST" in c), None)
    col_grade = next((c for c in df.columns if "GRADE" in c), None)

    out = pd.DataFrame({
        "ID": df[col_id].astype(str).str.replace(r"\.0$", "", regex=True).str.strip(),
        "FAMILY ID": df[col_fam].astype(str).str.replace(r"\.0$", "", regex=True).str.strip() if col_fam else "",
        "PARENT FIRST NAME": df[col_pf].astype(str).str.strip() if col_pf else "",
        "PARENT LAST NAME":  df[col_pl].astype(str).str.strip() if col_pl else "",
        "STUDENT FIRST NAME": df[col_sf].astype(str).str.strip() if col_sf else "",
        "STUDENT LAST NAME":  df[col_sl].astype(str).str.strip() if col_sl else "",
        "GRADE": df[col_grade].astype(str).str.strip() if col_grade else "",
        "REDIKER ID": df[col_red].astype(str).str.replace(r"\.0$", "", regex=True).str.strip() if col_red else "",
        "SOURCE": "SR",
    })
    out["UNIQUE_KEY"] = [make_unique_key_lenient(f, l, g) for f, l, g in zip(out["STUDENT FIRST NAME"], out["STUDENT LAST NAME"], out["GRADE"])]
    return out

# ------------------------------
# File upload interface
# ------------------------------
col1, col2, col3 = st.columns(3)
with col1:
    f_bb = st.file_uploader("Blackbaud", type=["xlsx","xls"])
with col2:
    f_red = st.file_uploader("Rediker", type=["xlsx","xls"])
with col3:
    f_sr = st.file_uploader("Student Records", type=["xlsx","xls"])

run = st.button("Build Master Excel", type="primary", disabled=not (f_bb and f_red and f_sr))

# ------------------------------
# Main process
# ------------------------------
if run:
    with st.spinner("Parsing & building Excel..."):
        bb_df = parse_blackbaud(f_bb)
        red_df = parse_rediker(f_red)
        sr_df = parse_student_records(f_sr)

        master = pd.concat([bb_df, red_df, sr_df], ignore_index=True)

        # Build lenient grouping
        master["__SURNAME"] = master["STUDENT LAST NAME"].apply(surname_first_token)
        master["__FIRSTTOK"] = master.apply(lambda r: firstname_first_token(r["STUDENT FIRST NAME"], r["STUDENT LAST NAME"]), axis=1)
        master["__GRADELEN"] = master["GRADE"].apply(grade_norm)
        master["UNIQUE_KEY"] = master["__SURNAME"] + "|" + master["__FIRSTTOK"] + "|" + master["__GRADELEN"]
        master["__GROUP_KEY"] = master["__SURNAME"] + "|" + master["__GRADELEN"]

        src_counts = master.groupby("__GROUP_KEY")["SOURCE"].nunique().to_dict()
        master["__SRC_PRESENT"] = master["__GROUP_KEY"].map(src_counts).fillna(0).astype(int)

        # Sort
        order = {"BB":0,"RED":1,"SR":2}
        master["_rank"] = master["SOURCE"].map(lambda x: order.get(x.upper(), 99))
        master = master.sort_values(["UNIQUE_KEY","_rank","STUDENT LAST NAME","STUDENT FIRST NAME"])

        # ---------- Summary ----------
        from collections import Counter
        summary = []
        for gkey, grp in master.groupby("__GROUP_KEY"):
            surname, grade = gkey.split("|",1)
            first_tokens = [t for t in grp["__FIRSTTOK"].tolist() if t]
            first_common = Counter(first_tokens).most_common(1)[0][0] if first_tokens else ""
            in_bb = any(grp["SOURCE"].str.upper()=="BB")
            in_red = any(grp["SOURCE"].str.upper()=="RED")
            in_sr = any(grp["SOURCE"].str.upper()=="SR")
            summary.append({
                "SURNAME": surname, "FIRST": first_common, "GRADE": grade,
                "BB": "✅" if in_bb else "❌",
                "RED": "✅" if in_red else "❌",
                "SR": "✅" if in_sr else "❌",
                "SOURCES_PRESENT": int(in_bb)+int(in_red)+int(in_sr)
            })
        summary = pd.DataFrame(summary)

        # ---------- Write Excel ----------
        import xlsxwriter
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            master.to_excel(writer, index=False, sheet_name="Master")
            summary.to_excel(writer, index=False, sheet_name="Summary")

        # ---------- Timestamped download ----------
        eastern = pytz.timezone("America/New_York")
        timestamp = datetime.now(eastern).strftime("%y%m%d_%H%M")
        file_name = f"{timestamp}_Master_Students_Combined_LENIENT_WITH_SUMMARY.xlsx"

        st.success("✅ Processing complete")
        st.download_button(
            label=f"⬇️ Download {file_name}",
            data=output.getvalue(),
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
