import io
import re
import pandas as pd
import streamlit as st

# ──────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG (must be first Streamlit call, and only once)
# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Roster → Master (LENIENT)", page_icon="📘", layout="centered")

st.title("📘 Master Students Builder — Lenient")
st.caption("Upload Blackbaud, Rediker, and Student Records → get a styled Excel with Master + Summary tabs.")

# ──────────────────────────────────────────────────────────────────────────────
# NORMALIZATION HELPERS
# ──────────────────────────────────────────────────────────────────────────────
def norm_piece(s: str) -> str:
    return re.sub(r"[^A-Z0-9 \-]+", "", str(s).upper()).strip()

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

def surname_last_token(last: str) -> str:
    s = norm_piece(last).replace("-", " ")
    toks = [t for t in s.split() if t]
    return toks[-1] if toks else ""

def firstname_first_token(first: str, last: str) -> str:
    ftoks = [t for t in norm_piece(first).split() if t]
    if ftoks:
        return ftoks[0]
    ltoks = [t for t in norm_piece(last).split() if t]
    return ltoks[0] if ltoks else ""

def make_unique_key_lenient(first: str, last: str, grade: str) -> str:
    return f"{surname_last_token(last)}|{firstname_first_token(first, last)}|{grade_norm(grade)}"

# ──────────────────────────────────────────────────────────────────────────────
# GENERIC COLUMN-FINDING UTILITIES
# ──────────────────────────────────────────────────────────────────────────────
def upper_cols(df):
    return {str(c).strip().upper(): c for c in df.columns}

def find_any(df, *need_tokens):
    """
    Return first column whose UPPER header contains ALL tokens in any of the
    provided token-tuples. Example:
      find_any(df, ("PARENT","FIRST"), ("PARENT 1","FIRST"), ("GUARDIAN","FIRST"))
    """
    U = upper_cols(df)
    for cand in df.columns:
        up = str(cand).strip().upper()
        for token_tuple in need_tokens:
            if all(tok in up for tok in token_tuple):
                return cand
    return None

def find_student_grade_blob_column(df):
    # Prefer columns that explicitly include both STUDENT and GRADE
    for c in df.columns:
        up = str(c).strip().upper()
        if "STUDENT" in up and "GRADE" in up:
            return c
    # Fallback: a column with many "(...)" endings (e.g., "LAST, FIRST (K)")
    scores = {c: df[c].astype(str).str.contains(r"\([^)]+\)\s*$", regex=True).sum() for c in df.columns}
    if not scores:
        return None
    best = max(scores, key=scores.get)
    return best if scores[best] >= 3 else None

# ──────────────────────────────────────────────────────────────────────────────
# BLACKBAUD PARSER (robust header detection; parent columns OPTIONAL)
# ──────────────────────────────────────────────────────────────────────────────
def parse_blackbaud(file) -> pd.DataFrame:
    # Detect header row in first 25 lines
    probe = pd.read_excel(file, header=None, nrows=25)
    want = ["FAMILY", "ID", "PARENT", "FIRST", "LAST", "STUDENT", "GRADE"]
    best_row, best_hits = 0, -1
    for i in range(len(probe)):
        row = " ".join(str(x).upper() for x in probe.iloc[i].tolist())
        hits = sum(w in row for w in want)
        if hits > best_hits:
            best_row, best_hits = i, hits

    df = pd.read_excel(file, header=best_row).fillna("")
    df.columns = [str(c).strip() for c in df.columns]

    # Flexible column finding
    fam_col = find_any(df, ("FAMILY","ID"))
    # Parent FIRST/LAST: accept broad variants and make OPTIONAL
    pf_col = find_any(
        df, ("PARENT","FIRST"),
        ("PARENT 1","FIRST"), ("P1","FIRST"),
        ("PRIMARY","PARENT","FIRST"),
        ("GUARDIAN","FIRST"),
        ("CONTACT 1","FIRST"), ("CONTACT1","FIRST")
    )
    pl_col = find_any(
        df, ("PARENT","LAST"),
        ("PARENT 1","LAST"), ("P1","LAST"),
        ("PRIMARY","PARENT","LAST"),
        ("GUARDIAN","LAST"),
        ("CONTACT 1","LAST"), ("CONTACT1","LAST")
    )
    stu_blob_col = find_student_grade_blob_column(df)

    # Only the student blob is truly required for rows; FamilyID helps but can be missing.
    if not stu_blob_col:
        st.error("Blackbaud: couldn’t find the student + (grade) column. Please check your export.")
        st.stop()

    # Prepare simple parsing helpers
    def split_students(cell: str):
        if pd.isna(cell) or str(cell).strip() == "":
            return []
        text = re.sub(r"\s*\)\s*[,/;|]?\s*", ")|", str(cell))
        return [p.strip().rstrip(",;/|") for p in text.split("|") if p.strip()]

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
            if len(toks) >= 3:
                last, first = toks[0], " ".join(toks[1:])
            elif len(toks) == 2:
                last, first = toks[0], toks[1]
            else:
                last, first = name, ""
        return last, first, grade

    rows = []
    for _, r in df.iterrows():
        fam = str(r.get(fam_col, "")).replace(".0","").strip() if fam_col else ""
        pf  = str(r.get(pf_col,  "")).strip() if pf_col else ""
        pl  = str(r.get(pl_col,  "")).strip() if pl_col else ""
        for entry in split_students(r.get(stu_blob_col, "")):
            l, f, g = parse_student_entry(entry)
            rows.append({
                "ID": "",
                "FAMILY ID": fam,
                "PARENT FIRST NAME": pf,
                "PARENT LAST NAME": pl,
                "STUDENT FIRST NAME": f,
                "STUDENT LAST NAME": l,
                "GRADE": g,
                "REDIKER ID": "",
                "SOURCE": "BB",
                "UNIQUE_KEY": make_unique_key_lenient(f, l, g),
            })

    # If parent columns were missing, inform but do not stop.
    if not pf_col or not pl_col:
        st.warning("Blackbaud: Parent First/Last columns not found. Proceeding with blanks for those fields.")

    return pd.DataFrame(rows)

# ──────────────────────────────────────────────────────────────────────────────
# REDIKER PARSER (robust; required cols guarded)
# ──────────────────────────────────────────────────────────────────────────────
def parse_rediker(file) -> pd.DataFrame:
    probe = pd.read_excel(file, header=None, nrows=12, usecols="A:K")
    tokens = {"APID","UNIQUE","STUDENT","FIRST","LAST","GRADE","LEVEL","GR","FAMILY","ID"}
    best_row, best_hits = 0, -1
    for i in range(len(probe)):
        row_vals = [str(x).strip().upper() for x in probe.iloc[i].tolist()]
        hits = sum(any(tok in cell for tok in tokens) for cell in row_vals)
        if hits > best_hits:
            best_row, best_hits = i, hits

    df = pd.read_excel(file, header=best_row, usecols="A:K").fillna("")
    df.columns = [str(c).strip() for c in df.columns]
    U = upper_cols(df)

    first_col = next((U[k] for k in U if k in ("FIRST","FIRST NAME","FIRST_NAME","STUDENT FIRST NAME")), None)
    last_col  = next((U[k] for k in U if k in ("LAST","LAST NAME","LAST_NAME","STUDENT LAST NAME")), None)
    name_col  = next((U[k] for k in U if k in ("STUDENT NAME","STUDENT_NAME","NAME")), None)
    grade_col = next((U[k] for k in U if k in ("GRADE","GRADE LEVEL","GR")), None)
    fam_col   = next((U[k] for k in U if "FAMILY" in k and "ID" in k), None)
    rid_col   = next((U[k] for k in U if k in ("APID","UNIQUE ID","UNIQUE_ID","REDIKER ID","REDIKERID","ID") and U[k] != fam_col), None)

    def split_student_name(val: str):
        if pd.isna(val) or str(val).strip() == "":
            return "", ""
        s = str(val).strip()
        if ";" in s:
            last, first = [t.strip() for t in s.split(";",1)]
        elif "," in s:
            last, first = [t.strip() for t in s.split(",",1)]
        else:
            parts = s.split()
            last, first = (parts[0], " ".join(parts[1:])) if len(parts) >= 2 else (s, "")
        return first, last

    if not (first_col and last_col) and name_col:
        split = df[name_col].apply(split_student_name).tolist()
        df["__First"], df["__Last"] = zip(*split) if split else ([], [])
        first_col, last_col = "__First", "__Last"

    required = {"FIRST name": first_col, "LAST name": last_col, "GRADE": grade_col}
    missing = [k for k, v in required.items() if not v]
    if missing:
        st.error(f"Rediker: couldn’t find required column(s): {', '.join(missing)}.")
        st.stop()

    rows = []
    for _, r in df.iterrows():
        fam   = str(r.get(fam_col, "")).replace(".0","").strip() if fam_col else ""
        rid   = str(r.get(rid_col, "")).replace(".0","").strip() if rid_col else ""
        first = str(r.get(first_col, "")).strip() if first_col else ""
        last  = str(r.get(last_col,  "")).strip() if last_col  else ""
        grade = str(r.get(grade_col, "")).strip() if grade_col else ""
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

# ──────────────────────────────────────────────────────────────────────────────
# STUDENT RECORDS PARSER (guards for required; flexible IDs/parents)
# ──────────────────────────────────────────────────────────────────────────────
def parse_student_records(file) -> pd.DataFrame:
    df = pd.read_excel(file).fillna("")
    U = upper_cols(df)

    col_id   = list(df.columns)[0] if len(df.columns) else None
    col_fam  = U.get("FAMILY ID") or U.get("FAMILYID") or U.get("FAMILY_ID")
    col_red  = U.get("REDIKER ID") or U.get("REDIKERID") or U.get("REDIKER_ID")
    col_pf   = U.get("PARENT FIRST NAME") or U.get("PARENT FIRST")
    col_pl   = U.get("PARENT LAST NAME")  or U.get("PARENT LAST")
    col_sf   = U.get("STUDENT FIRST NAME") or U.get("FIRST NAME") or U.get("FIRST")
    col_sl   = U.get("STUDENT LAST NAME")  or U.get("LAST NAME")  or U.get("LAST")
    col_grade= U.get("GRADE") or U.get("GRADE LEVEL") or U.get("GR")

    required = {"FIRST name": col_sf, "LAST name": col_sl, "GRADE": col_grade}
    missing = [k for k, v in required.items() if not v]
    if missing:
        st.error(f"Student Records: couldn’t find required column(s): {', '.join(missing)}.")
        st.stop()

    out = pd.DataFrame({
        "ID": df[col_id].astype(str).str.replace(r"\.0$", "", regex=True).str.strip() if col_id else "",
        "FAMILY ID": df[col_fam].astype(str).str.replace(r"\.0$", "", regex=True).str.strip() if col_fam else "",
        "PARENT FIRST NAME": df[col_pf].astype(str).str.strip() if col_pf else "",
        "PARENT LAST NAME":  df[col_pl].astype(str).str.strip() if col_pl else "",
        "STUDENT FIRST NAME": df[col_sf].astype(str).str.strip(),
        "STUDENT LAST NAME":  df[col_sl].astype(str).str.strip(),
        "GRADE": df[col_grade].astype(str).str.strip(),
        "REDIKER ID": df[col_red].astype(str).str.replace(r"\.0$", "", regex=True).str.strip() if col_red else "",
        "SOURCE": "SR",
    })
    out["UNIQUE_KEY"] = [make_unique_key_lenient(f, l, g) for f, l, g in zip(out["STUDENT FIRST NAME"], out["STUDENT LAST NAME"], out["GRADE"])]
    return out

# ──────────────────────────────────────────────────────────────────────────────
# UI — UPLOADS
# ──────────────────────────────────────────────────────────────────────────────
col1, col2, col3 = st.columns(3)
with col1:
    f_bb = st.file_uploader("Blackbaud", type=["xlsx","xls"])
with col2:
    f_red = st.file_uploader("Rediker", type=["xlsx","xls"])
with col3:
    f_sr = st.file_uploader("Student Records", type=["xlsx","xls"])

run = st.button("Build Master Excel", type="primary", disabled=not (f_bb and f_red and f_sr))

# ──────────────────────────────────────────────────────────────────────────────
# PROCESS
# ──────────────────────────────────────────────────────────────────────────────
if run:
    with st.spinner("Parsing & normalizing..."):
        bb_df  = parse_blackbaud(f_bb)
        red_df = parse_rediker(f_red)
        sr_df  = parse_student_records(f_sr)

    TARGET = [
        "ID","FAMILY ID","PARENT FIRST NAME","PARENT LAST NAME",
        "STUDENT FIRST NAME","STUDENT LAST NAME","GRADE","REDIKER ID","SOURCE","UNIQUE_KEY"
    ]
    master = pd.concat([bb_df[TARGET], red_df[TARGET], sr_df[TARGET]], ignore_index=True)

    # Build helpers for lenient grouping/presence (LAST surname token)
    master["__SURNAME_TOKEN"] = master["STUDENT LAST NAME"].apply(surname_last_token)
    master["__FIRSTTOK"]      = master.apply(lambda r: firstname_first_token(r["STUDENT FIRST NAME"], r["STUDENT LAST NAME"]), axis=1)
    master["__GRADELEN"]      = master["GRADE"].apply(grade_norm)
    master["__GROUP_KEY"]     = master["__SURNAME_TOKEN"] + "|" + master["__GRADELEN"]

    src_counts = master.groupby("__GROUP_KEY")["SOURCE"].nunique().to_dict()
    master["__SRC_PRESENT"] = master["__GROUP_KEY"].map(src_counts).fillna(0).astype(int)

    # Sort by key then source
    order = {"BB":0, "RED":1, "SR":2}
    master["_source_rank"] = master["SOURCE"].map(lambda x: order.get(str(x).upper(), 99))
    master["UNIQUE_KEY"] = master["__SURNAME_TOKEN"] + "|" + master["__FIRSTTOK"] + "|" + master["__GRADELEN"]
    master_sorted = master.sort_values(by=["UNIQUE_KEY","_source_rank","STUDENT LAST NAME","STUDENT FIRST NAME"], kind="mergesort").reset_index(drop=True)

    # SUMMARY
    from collections import Counter
    summary_rows = []
    for gkey, grp in master.groupby("__GROUP_KEY"):
        surname_token, grade = gkey.split("|", 1)
        first_tokens = [t for t in grp["__FIRSTTOK"].tolist() if t]
        first_common = Counter(first_tokens).most_common(1)[0][0] if first_tokens else ""
        in_bb  = "BB"  in grp["SOURCE"].str.upper().values
        in_red = "RED" in grp["SOURCE"].str.upper().values
        in_sr  = "SR"  in grp["SOURCE"].str.upper().values
        summary_rows.append({
            "SURNAME_TOKEN(LAST)": surname_token,
            "FIRST_TOKEN": first_common,
            "GRADE": grade,
            "BB": "✅" if in_bb else "❌",
            "RED": "✅" if in_red else "❌",
            "SR": "✅" if in_sr else "❌",
            "SOURCES_PRESENT": int(in_bb) + int(in_red) + int(in_sr),
        })
    summary = pd.DataFrame(summary_rows).sort_values(["SURNAME_TOKEN(LAST)","GRADE","FIRST_TOKEN"]).reset_index(drop=True)

    # DIAGNOSTICS
    with st.expander("Diagnostics (detected rows & sample keys)"):
        st.write("Blackbaud rows:", len(bb_df), "Rediker rows:", len(red_df), "Student Records rows:", len(sr_df))
        st.dataframe(master_sorted[["SOURCE","FAMILY ID","PARENT FIRST NAME","PARENT LAST NAME","STUDENT LAST NAME","STUDENT FIRST NAME","GRADE","UNIQUE_KEY"]].head(20))

    # WRITE EXCEL (+ styling)
    import xlsxwriter
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # Sheet 1: Master
        master_sorted.to_excel(writer, index=False, sheet_name="Master")
        wb = writer.book
        ws1 = writer.sheets["Master"]
        header_fmt = wb.add_format({"bold": True})
        fmt_bb = wb.add_format({"font_color": "#000000"})
        fmt_red = wb.add_format({"font_color": "#A10000"})
        fmt_sr = wb.add_format({"font_color": "#006400"})
        warn_fill = "#FFF59D"
        fmt_bb_warn  = wb.add_format({"font_color": "#000000", "bg_color": warn_fill, "bold": True})
        fmt_red_warn = wb.add_format({"font_color": "#A10000", "bg_color": warn_fill, "bold": True})
        fmt_sr_warn  = wb.add_format({"font_color": "#006400", "bg_color": warn_fill, "bold": True})
        # header
        for c_idx, col in enumerate(master_sorted.columns):
            ws1.write(0, c_idx, col, header_fmt)
        # autosize
        for i, col in enumerate(master_sorted.columns):
            vals = master_sorted[col].astype(str).head(2000).tolist()
            width = min(max([len(str(col))] + [len(v) for v in vals]) + 2, 40)
            ws1.set_column(i, i, width)
        idx = {c:i for i,c in enumerate(master_sorted.columns)}
        s_col = idx["SOURCE"]; present_col = idx["__SRC_PRESENT"]
        n_rows, n_cols = master_sorted.shape
        for r in range(n_rows):
            src = str(master_sorted.iat[r, s_col]).strip().upper()
            present_all = int(master_sorted.iat[r, present_col]) >= 3
            base_fmt, warn_fmt = (fmt_bb, fmt_bb_warn)
            if src == "RED": base_fmt, warn_fmt = (fmt_red, fmt_red_warn)
            elif src == "SR": base_fmt, warn_fmt = (fmt_sr, fmt_sr_warn)
            fmt = base_fmt if present_all else warn_fmt
            for c in range(n_cols):
                ws1.write(r + 1, c, master_sorted.iat[r, c], fmt)
        # hide helpers
        for helper in ["__SURNAME_TOKEN","__FIRSTTOK","__GRADELEN","__GROUP_KEY","__SRC_PRESENT","_source_rank"]:
            if helper in idx:
                ws1.set_column(idx[helper], idx[helper], None, None, {"hidden": True})

        # Sheet 2: Summary
        summary.to_excel(writer, index=False, sheet_name="Summary")
        ws2 = writer.sheets["Summary"]
        header_fmt2 = wb.add_format({"bold": True})
        ok_fmt  = wb.add_format({"bg_color": "#C6EFCE", "font_color": "#006100"})
        bad_fmt = wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})
        for c_idx, col in enumerate(summary.columns):
            ws2.write(0, c_idx, col, header_fmt2)
        for i, col in enumerate(summary.columns):
            vals = summary[col].astype(str).head(2000).tolist()
            width = min(max([len(str(col))] + [len(v) for v in vals]) + 2, 50)
            ws2.set_column(i, i, width)
        col_idx = {c:i for i,c in enumerate(summary.columns)}
        for r in range(len(summary)):
            for src_col in ["BB","RED","SR"]:
                val = summary.iat[r, col_idx[src_col]]
                ws2.write(r + 1, col_idx[src_col], val, ok_fmt if val == "✅" else bad_fmt)

    st.success("✅ Excel generated successfully")
    st.download_button(
        label="⬇️ Download Master_Students_Combined_LENIENT_WITH_SUMMARY.xlsx",
        data=output.getvalue(),
        file_name="Master_Students_Combined_LENIENT_WITH_SUMMARY.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
