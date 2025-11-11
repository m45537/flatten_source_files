import io, re
import pandas as pd
import streamlit as st

# --------------------------------------------------------------------
# APP CONFIG
# --------------------------------------------------------------------
st.set_page_config(page_title="Roster → Master (LENIENT)", page_icon="📘", layout="centered")
st.title("📘 Master Students Builder — Lenient")
st.caption("Upload Blackbaud, Rediker, Student Records → get a styled Excel with Master + Summary tabs.")

# --------------------------------------------------------------------
# NORMALIZATION HELPERS
# --------------------------------------------------------------------
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

# --------------------------------------------------------------------
# PARSERS
# --------------------------------------------------------------------
def parse_blackbaud(file) -> pd.DataFrame:
    df = pd.read_excel(file, usecols=list(range(5)))
    df.columns = [str(c).strip().upper() for c in df.columns]
    col_fam  = next((c for c in df.columns if "FAMILY" in c and "ID" in c), None)
    col_pf   = next((c for c in df.columns if "FIRST" in c and "PARENT" in c), None)
    col_pl   = next((c for c in df.columns if "LAST" in c and "PARENT" in c), None)
    col_stu  = next((c for c in df.columns if "STUDENT" in c and "GRADE" in c), None)

    def split_students(cell):
        if pd.isna(cell) or str(cell).strip() == "":
            return []
        text = re.sub(r"\s*\)\s*[,/;|]?\s*", ")|", str(cell))
        return [p.strip().rstrip(",;/|") for p in text.split("|") if p.strip()]

    def parse_student_entry(entry):
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
        fam = str(r.get(col_fam, "")).replace(".0", "").strip()
        pf = str(r.get(col_pf, "")).strip()
        pl = str(r.get(col_pl, "")).strip()
        for entry in split_students(r.get(col_stu, "")):
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
    return pd.DataFrame(rows)

def parse_rediker(file) -> pd.DataFrame:
    df = pd.read_excel(file, header=0, usecols="A:K").fillna("")
    df.columns = [str(c).strip() for c in df.columns]
    U = {c.upper(): c for c in df.columns}
    first_col = next((U[k] for k in U if "FIRST" in k and "NAME" in k), None)
    last_col = next((U[k] for k in U if "LAST" in k and "NAME" in k), None)
    grade_col = next((U[k] for k in U if "GRADE" in k), None)
    fam_col = next((U[k] for k in U if "FAMILY" in k and "ID" in k), None)
    rid_col = next((U[k] for k in U if "ID" in k and not "FAMILY" in k), None)

    rows = []
    for _, r in df.iterrows():
        fam = str(r.get(fam_col, "")).replace(".0", "").strip() if fam_col else ""
        rid = str(r.get(rid_col, "")).replace(".0", "").strip() if rid_col else ""
        first = str(r.get(first_col, "")).strip() if first_col else ""
        last = str(r.get(last_col, "")).strip() if last_col else ""
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

def parse_student_records(file) -> pd.DataFrame:
    df = pd.read_excel(file).fillna("")
    U = {str(c).strip().upper(): c for c in df.columns}
    col_id = list(df.columns)[0]
    col_fam = U.get("FAMILY ID") or U.get("FAMILYID")
    col_red = U.get("REDIKER ID") or U.get("REDIKERID")
    col_pf = U.get("PARENT FIRST NAME")
    col_pl = U.get("PARENT LAST NAME")
    col_sf = U.get("STUDENT FIRST NAME")
    col_sl = U.get("STUDENT LAST NAME")
    col_grade = U.get("GRADE")

    out = pd.DataFrame({
        "ID": df[col_id].astype(str).str.replace(r"\.0$", "", regex=True),
        "FAMILY ID": df[col_fam].astype(str).str.replace(r"\.0$", "", regex=True),
        "PARENT FIRST NAME": df[col_pf],
        "PARENT LAST NAME": df[col_pl],
        "STUDENT FIRST NAME": df[col_sf],
        "STUDENT LAST NAME": df[col_sl],
        "GRADE": df[col_grade],
        "REDIKER ID": df[col_red].astype(str).str.replace(r"\.0$", "", regex=True),
        "SOURCE": "SR",
    })
    out["UNIQUE_KEY"] = [make_unique_key_lenient(f, l, g) for f, l, g in zip(out["STUDENT FIRST NAME"], out["STUDENT LAST NAME"], out["GRADE"])]
    return out

# --------------------------------------------------------------------
# UI
# --------------------------------------------------------------------
col1, col2, col3 = st.columns(3)
with col1:
    f_bb = st.file_uploader("Blackbaud", type=["xlsx","xls"])
with col2:
    f_red = st.file_uploader("Rediker", type=["xlsx","xls"])
with col3:
    f_sr = st.file_uploader("Student Records", type=["xlsx","xls"])

run = st.button("Build Master Excel", type="primary", disabled=not (f_bb and f_red and f_sr))

# --------------------------------------------------------------------
# PROCESSING
# --------------------------------------------------------------------
if run:
    bb_df = parse_blackbaud(f_bb)
    red_df = parse_rediker(f_red)
    sr_df = parse_student_records(f_sr)
    TARGET = ["ID","FAMILY ID","PARENT FIRST NAME","PARENT LAST NAME",
              "STUDENT FIRST NAME","STUDENT LAST NAME","GRADE","REDIKER ID","SOURCE","UNIQUE_KEY"]
    master = pd.concat([bb_df[TARGET], red_df[TARGET], sr_df[TARGET]], ignore_index=True)

    master["__SURNAME_TOKEN"] = master["STUDENT LAST NAME"].apply(surname_last_token)
    master["__FIRSTTOK"] = master.apply(lambda r: firstname_first_token(r["STUDENT FIRST NAME"], r["STUDENT LAST NAME"]), axis=1)
    master["__GRADELEN"] = master["GRADE"].apply(grade_norm)
    master["__GROUP_KEY"] = master["__SURNAME_TOKEN"] + "|" + master["__GRADELEN"]

    src_counts = master.groupby("__GROUP_KEY")["SOURCE"].nunique().to_dict()
    master["__SRC_PRESENT"] = master["__GROUP_KEY"].map(src_counts).fillna(0).astype(int)

    # Sort by unique key then source
    order = {"BB":0, "RED":1, "SR":2}
    master["_source_rank"] = master["SOURCE"].map(lambda x: order.get(x, 99))
    master_sorted = master.sort_values(by=["UNIQUE_KEY","_source_rank"], kind="mergesort")

    # ----------------------------------------------------------------
    # SUMMARY + STYLING
    # ----------------------------------------------------------------
    from collections import Counter
    summary = []
    for gkey, grp in master.groupby("__GROUP_KEY"):
        surname_token, grade = gkey.split("|", 1)
        first_tokens = [t for t in grp["__FIRSTTOK"] if t]
        first_common = Counter(first_tokens).most_common(1)[0][0] if first_tokens else ""
        in_bb  = "BB" in grp["SOURCE"].values
        in_red = "RED" in grp["SOURCE"].values
        in_sr  = "SR" in grp["SOURCE"].values
        summary.append({
            "SURNAME_TOKEN": surname_token,
            "FIRST_TOKEN": first_common,
            "GRADE": grade,
            "BB": "✅" if in_bb else "❌",
            "RED": "✅" if in_red else "❌",
            "SR": "✅" if in_sr else "❌",
            "COUNT_PRESENT": sum([in_bb, in_red, in_sr]),
        })
    summary = pd.DataFrame(summary).sort_values(["SURNAME_TOKEN","GRADE","FIRST_TOKEN"])

    import xlsxwriter
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        master_sorted.to_excel(writer, index=False, sheet_name="Master")
        summary.to_excel(writer, index=False, sheet_name="Summary")

        wb = writer.book
        ws1 = writer.sheets["Master"]
        ws2 = writer.sheets["Summary"]
        bold = wb.add_format({"bold": True})
        fmt_bb = wb.add_format({"font_color": "#000000"})
        fmt_red = wb.add_format({"font_color": "#A10000"})
        fmt_sr = wb.add_format({"font_color": "#006400"})
        warn_fill = "#FFF59D"
        fmt_bb_warn = wb.add_format({"font_color": "#000000","bg_color":warn_fill,"bold":True})
        fmt_red_warn = wb.add_format({"font_color": "#A10000","bg_color":warn_fill,"bold":True})
        fmt_sr_warn = wb.add_format({"font_color": "#006400","bg_color":warn_fill,"bold":True})
        ok_fmt = wb.add_format({"bg_color":"#C6EFCE","font_color":"#006100"})
        bad_fmt = wb.add_format({"bg_color":"#FFC7CE","font_color":"#9C0006"})

        for i, c in enumerate(master_sorted.columns):
            ws1.write(0, i, c, bold)
            width = min(max([len(str(c))] + [len(str(v)) for v in master_sorted[c].astype(str).head(2000)]) + 2, 40)
            ws1.set_column(i, i, width)

        s_col = master_sorted.columns.get_loc("SOURCE")
        p_col = master_sorted.columns.get_loc("__SRC_PRESENT")
        n_rows, n_cols = master_sorted.shape
        for r in range(n_rows):
            src = str(master_sorted.iat[r, s_col]).upper()
            all3 = int(master_sorted.iat[r, p_col]) >= 3
            fmt = { "BB": fmt_bb, "RED": fmt_red, "SR": fmt_sr }.get(src, fmt_bb)
            warn = { "BB": fmt_bb_warn, "RED": fmt_red_warn, "SR": fmt_sr_warn }.get(src, fmt_bb_warn)
            style = fmt if all3 else warn
            for c in range(n_cols):
                ws1.write(r+1, c, master_sorted.iat[r, c], style)

        for i, c in enumerate(summary.columns):
            ws2.write(0, i, c, bold)
            width = min(max([len(str(c))] + [len(str(v)) for v in summary[c].astype(str).head(2000)]) + 2, 40)
            ws2.set_column(i, i, width)
        for r in range(len(summary)):
            for src_col in ["BB","RED","SR"]:
                val = summary.at[summary.index[r], src_col]
                ws2.write(r+1, summary.columns.get_loc(src_col), val, ok_fmt if val=="✅" else bad_fmt)

    st.success("✅ Excel generated successfully")
    st.download_button(
        "⬇️ Download Master_Students_Combined_LENIENT_WITH_SUMMARY.xlsx",
        data=output.getvalue(),
        file_name="Master_Students_Combined_LENIENT_WITH_SUMMARY.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
