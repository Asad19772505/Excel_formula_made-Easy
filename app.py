import io
from datetime import datetime
from typing import Dict, List, Optional

import numpy as np
import pandas as pd
import streamlit as st

# ---------------------------
# Page Config
# ---------------------------
st.set_page_config(
    page_title="Excel Formula Companion",
    page_icon=None,
    layout="wide"
)

# ---------------------------
# Helpers
# ---------------------------
@st.cache_data(show_spinner=False)
def load_workbook(file_bytes: bytes) -> Dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    sheets = {}
    for name in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=name)
            df.columns = [str(c).strip() for c in df.columns]
            sheets[name] = df
        except Exception:
            pass
    return sheets

def df_multifilter(
    df: pd.DataFrame,
    category_col: Optional[str],
    category_values: List[str],
    search_cols: List[str],
    search_term: str
) -> pd.DataFrame:
    out = df.copy()
    if category_col and category_col in out.columns and category_values:
        out = out[out[category_col].isin(category_values)]
    if search_term:
        mask = pd.Series(False, index=out.index)
        for col in search_cols:
            if col in out.columns:
                mask = mask | out[col].fillna("").astype(str).str.contains(search_term, case=False, na=False)
        out = out[mask]
    return out

def to_excel_bytes(dfs: Dict[str, pd.DataFrame]) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        for sheet_name, df in dfs.items():
            safe_name = sheet_name[:31] if sheet_name else "Sheet"
            df.to_excel(writer, index=False, sheet_name=safe_name)
            ws = writer.sheets[safe_name]
            for idx, col in enumerate(df.columns):
                width = max(10, min(60, int(df[col].astype(str).str.len().quantile(0.9)) + 4))
                ws.set_column(idx, idx, width)
    return buf.getvalue()

def to_pdf_bytes(title: str, sections: Dict[str, pd.DataFrame]) -> bytes:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import cm

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    W, H = A4

    def draw_header(page_title):
        c.setFont("Helvetica-Bold", 14)
        c.drawString(2*cm, H-2.0*cm, page_title)
        c.setFont("Helvetica", 9)
        c.drawString(2*cm, H-2.6*cm, "Generated: " + datetime.now().strftime("%Y-%m-%d %H:%M"))

    y = H - 3.0*cm
    draw_header(title)

    for section, df in sections.items():
        c.setFont("Helvetica-Bold", 11)
        c.drawString(2*cm, y, section)
        y -= 0.5*cm

        c.setFont("Helvetica", 8)
        if not df.empty:
            cols = list(df.columns)[:5]
            x = 2*cm
            for col in cols:
                c.drawString(x, y, str(col)[:28])
                x += 6*cm
            y -= 0.4*cm

            max_rows = 20
            for i in range(min(max_rows, len(df))):
                x = 2*cm
                for col in cols:
                    cell = str(df.iloc[i][col]) if col in df.columns else ""
                    c.drawString(x, y, cell[:28])
                    x += 6*cm
                y -= 0.35*cm
                if y < 2.5*cm:
                    c.showPage()
                    y = H - 3.0*cm
                    draw_header(title)
        else:
            c.drawString(2*cm, y, "(No rows)")
            y -= 0.4*cm

        y -= 0.4*cm
        if y < 2.5*cm:
            c.showPage()
            y = H - 3.0*cm
            draw_header(title)

    c.save()
    return buf.getvalue()

def chip(text: str):
    st.markdown(
        """
        <span style="
            display:inline-block;
            padding:4px 10px;
            border-radius:999px;
            border:1px solid #e5e7eb;
            background:#f8fafc;
            font-size:12px;
            margin-right:6px;">
            {t}
        </span>
        """.format(t=text),
        unsafe_allow_html=True
    )

# ---------- Formula Playground helpers ----------
def structured_ref(table: str, col: str) -> str:
    return f"{table}[{col}]"

def structured_ref_multi(table: str, cols: List[str]) -> str:
    inner = "],[".join(cols)
    return f"{table}[[{inner}]]"

def excel_literal(val):
    # Quote text for Excel, leave numerics unquoted
    try:
        if val is None:
            return '""'
        if isinstance(val, bool):
            return "TRUE" if val else "FALSE"
        f = float(val)
        if np.isfinite(f):
            return str(val)
    except Exception:
        pass
    s = str(val)
    s = s.replace('"', '""')
    return f'"{s}"'

def excel_criteria(op: str, value):
    lit = excel_literal(value)
    if op == "equals":
        return lit
    if op == "not equals":
        # Use string literal properly without smart quotes
        v = lit
        if v.startswith('"') and v.endswith('"'):
            inner = v[1:-1]
            return f'"<>{inner}"'
        return f'"<>{v}"'
    if op == "contains":
        # "*text*"
        if lit.startswith('"') and lit.endswith('"'):
            return f'"*{lit[1:-1]}*"'
        return f'"*{lit}*"'
    if op in {">", ">=", "<", "<="}:
        return f'"{op}"&{lit}'
    return lit

@st.cache_data(show_spinner=False)
def load_user_dataset(file) -> pd.DataFrame:
    name = file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(file)
    elif name.endswith((".xlsx", ".xlsm", ".xls")):
        return pd.read_excel(file)
    else:
        raise ValueError("Unsupported file type. Upload CSV or Excel.")

def sample_data() -> pd.DataFrame:
    np.random.seed(1)
    n = 50
    seg = np.random.choice(["Retail","Gov","SMB"], size=n)
    prod = np.random.choice(["Alpha","Beta","Gamma"], size=n)
    country = np.random.choice(["KSA","UAE","EGY"], size=n)
    units = np.random.randint(1, 200, size=n)
    price = np.random.uniform(5, 50, size=n).round(2)
    sales = (units * price).round(2)
    return pd.DataFrame({
        "Segment": seg,
        "Product": prod,
        "Country": country,
        "Units": units,
        "Price": price,
        "Sales": sales
    })

# ---------------------------
# UI
# ---------------------------
st.title("Excel Formula Companion")
st.caption("Search, Filter, Compare, Learn, Export, and Playground")

with st.sidebar:
    st.header("Step 1 - Upload Master Workbook")
    uploaded = st.file_uploader(
        "Upload your Excel master file",
        type=["xlsx", "xlsm", "xls"],
        accept_multiple_files=False
    )

    st.markdown("---")
    st.subheader("Global Search and Filters")

    search_term = st.text_input("Search text (function, description, syntax, hints)", "")
    selected_categories = []
    selected_functions = []

# ---------------------------
# Load Master Data
# ---------------------------
if not uploaded:
    st.info("Upload your Excel_Formulas_Master_Reference_With_Filters.xlsx to begin.")
    st.stop()

sheets = load_workbook(uploaded.read())

FULL_REF = "Full Reference"
FORM_LIST = "Formulas (List)"
CAT_INDEX = "Category Index"
CHEATS = "Usage Cheat-Sheet"
QUICK = "Quick Examples"
CONTROLS = "Controls"
FILTERED = "Filtered View"

full_ref = sheets.get(FULL_REF, pd.DataFrame())
form_list = sheets.get(FORM_LIST, pd.DataFrame())
cat_index = sheets.get(CAT_INDEX, pd.DataFrame())
cheats = sheets.get(CHEATS, pd.DataFrame())
quick = sheets.get(QUICK, pd.DataFrame())
controls = sheets.get(CONTROLS, pd.DataFrame())
filtered_view = sheets.get(FILTERED, pd.DataFrame())

if not full_ref.empty:
    categories = sorted([c for c in full_ref.get("Category", pd.Series()).dropna().unique()])
    functions = sorted([c for c in full_ref.get("Function", pd.Series()).dropna().unique()])
else:
    categories = sorted([c for c in form_list.get("Category", pd.Series()).dropna().unique()])
    functions = sorted([c for c in form_list.get("Function", pd.Series()).dropna().unique()])

with st.sidebar:
    if categories:
        selected_categories = st.multiselect("Filter by Category", categories, default=[])
    if functions:
        selected_functions = st.multiselect("Filter by Function", functions, default=[])

# ---------------------------
# Tabs
# ---------------------------
tabs = st.tabs([
    "Full Reference",
    "Filtered View",
    "Usage Cheat-Sheet",
    "Category Comparisons",
    "Quick Examples",
    "Flashcards",
    "Exports",
    "Formula Playground"
])

# ---------------------------
# Full Reference Tab
# ---------------------------
with tabs[0]:
    st.subheader("Full Reference")
    if full_ref.empty:
        st.warning("No data found in Full Reference.")
    else:
        base_df = full_ref.copy()
        df = df_multifilter(
            base_df,
            category_col="Category" if "Category" in base_df.columns else None,
            category_values=selected_categories,
            search_cols=[c for c in ["Function", "Short Description", "Syntax Template", "Example Formula (copy/paste)", "Hints"] if c in base_df.columns],
            search_term=search_term
        )
        if selected_functions and "Function" in df.columns:
            df = df[df["Function"].isin(selected_functions)]

        st.caption(f"{len(df):,} of {len(base_df):,} rows")
        st.dataframe(df, use_container_width=True)

        st.markdown("#### Details")
        detail_col1, detail_col2 = st.columns([1, 1])
        with detail_col1:
            pick = st.selectbox("Inspect a function", sorted(df["Function"].dropna().unique()) if "Function" in df.columns else [])
        with detail_col2:
            st.write("")
        if pick:
            sel = df[df["Function"] == pick].head(1)
            if not sel.empty:
                row = sel.iloc[0].to_dict()
                st.write("Category:", row.get("Category", ""))
                st.write("Short Description:", row.get("Short Description", ""))
                st.write("Syntax Template:")
                st.code(str(row.get("Syntax Template", "")))
                st.write("Example Formula (copy/paste):")
                st.code(str(row.get("Example Formula (copy/paste)", "")))
                st.write("Hints:", row.get("Hints", ""))
                chip(row.get("Category", ""))
                chip(pick)

# ---------------------------
# Filtered View Tab
# ---------------------------
with tabs[1]:
    st.subheader("Filtered View")
    if filtered_view.empty and full_ref.empty:
        st.warning("No data found in Filtered View (using Full Reference instead).")
        df_base = full_ref.copy()
    else:
        df_base = (filtered_view if not filtered_view.empty else full_ref).copy()

    df_f = df_multifilter(
        df_base,
        category_col="Category" if "Category" in df_base.columns else None,
        category_values=selected_categories,
        search_cols=[c for c in df_base.columns if df_base[c].dtype == "O"],
        search_term=search_term
    )
    if selected_functions and "Function" in df_f.columns:
        df_f = df_f[df_f["Function"].isin(selected_functions)]

    st.caption(f"{len(df_f):,} rows")
    st.dataframe(df_f, use_container_width=True)

# ---------------------------
# Usage Cheat-Sheet Tab
# ---------------------------
with tabs[2]:
    st.subheader("Usage Cheat-Sheet")
    if cheats.empty:
        st.info("No data in Usage Cheat-Sheet.")
    else:
        st.dataframe(cheats, use_container_width=True)
        topic_col = "Comparison / Topic"
        guide_col = "Usage Guidance"
        if topic_col in cheats.columns and guide_col in cheats.columns:
            st.markdown("---")
            st.markdown("#### Topics")
            for _, r in cheats.iterrows():
                with st.expander(str(r.get(topic_col, "Topic"))):
                    st.write(r.get(guide_col, ""))

# ---------------------------
# Category Comparisons Tab
# ---------------------------
with tabs[3]:
    st.subheader("Category Comparisons")
    if cat_index.empty:
        st.info("No data in Category Index.")
    else:
        st.dataframe(cat_index, use_container_width=True)
        st.markdown("---")
        st.caption("Tip: Use this section to choose the right function for a scenario.")

# ---------------------------
# Quick Examples Tab
# ---------------------------
with tabs[4]:
    st.subheader("Quick Examples")
    if quick.empty:
        st.info("No data in Quick Examples.")
    else:
        st.dataframe(quick, use_container_width=True)
        fn_col = "Function" if "Function" in quick.columns else None
        if fn_col:
            pick2 = st.selectbox("Show example for function", sorted(quick[fn_col].dropna().unique()))
            ex_df = quick[quick[fn_col] == pick2]
            st.write(ex_df)

# ---------------------------
# Flashcards (Learning Mode) - fixed slider
# ---------------------------
with tabs[5]:
    st.subheader("Flashcards (Learning Mode)")
    if full_ref.empty:
        st.info("Need Full Reference to use flashcards.")
    else:
        pool = df_multifilter(
            full_ref,
            category_col="Category" if "Category" in full_ref.columns else None,
            category_values=selected_categories,
            search_cols=[c for c in ["Function", "Short Description", "Syntax Template", "Hints"] if c in full_ref.columns],
            search_term=search_term
        )
        if selected_functions and "Function" in pool.columns:
            pool = pool[pool["Function"].isin(selected_functions)]

        if pool.empty:
            st.info("No rows match your current filters.")
        else:
            n = len(pool)
            if n == 1:
                st.info("Only 1 flashcard available with current filters.")
                idx = 1
            else:
                default_val = 1 if "flash_idx" not in st.session_state else min(st.session_state.get("flash_idx", 1), n)
                idx = st.slider("Flashcard index", min_value=1, max_value=n, value=default_val, key="flash_idx")
            row = pool.iloc[idx-1].to_dict()
            st.markdown("### " + str(row.get("Function", "Function")))
            st.write("Category:", row.get("Category", ""))
            st.write("Short Description:", row.get("Short Description", ""))
            with st.expander("Show Syntax"):
                st.code(str(row.get("Syntax Template", "")))
            with st.expander("Show Example"):
                st.code(str(row.get("Example Formula (copy/paste)", "")))
            with st.expander("Hints"):
                st.write(row.get("Hints", ""))

# ---------------------------
# Exports Tab
# ---------------------------
with tabs[6]:
    st.subheader("Exports")
    st.caption("Export your current, filtered views.")

    export_slices = {}

    if not full_ref.empty:
        df1 = df_multifilter(
            full_ref,
            "Category" if "Category" in full_ref.columns else None,
            selected_categories,
            [c for c in ["Function", "Short Description", "Syntax Template", "Example Formula (copy/paste)", "Hints"] if c in full_ref.columns],
            search_term
        )
        if selected_functions and "Function" in df1.columns:
            df1 = df1[df1["Function"].isin(selected_functions)]
        export_slices["Full Reference (filtered)"] = df1

    if not filtered_view.empty:
        df2 = df_multifilter(
            filtered_view,
            "Category" if "Category" in filtered_view.columns else None,
            selected_categories,
            [c for c in filtered_view.columns if filtered_view[c].dtype == "O"],
            search_term
        )
        if selected_functions and "Function" in df2.columns:
            df2 = df2[df2["Function"].isin(selected_functions)]
        export_slices["Filtered View (filtered)"] = df2

    if not cheats.empty:
        export_slices["Usage Cheat-Sheet"] = cheats
    if not cat_index.empty:
        export_slices["Category Index"] = cat_index
    if not quick.empty:
        export_slices["Quick Examples"] = quick

    excel_bytes = to_excel_bytes(export_slices)
    st.download_button(
        label="Download Excel (.xlsx)",
        data=excel_bytes,
        file_name="Excel_Formula_Companion_" + datetime.now().strftime("%Y%m%d_%H%M") + ".xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    try:
        pdf_bytes = to_pdf_bytes(
            "Excel Formula Companion - Export",
            export_slices
        )
        st.download_button(
            label="Download PDF (summary)",
            data=pdf_bytes,
            file_name="Excel_Formula_Companion_" + datetime.now().strftime("%Y%m%d_%H%M") + ".pdf",
            mime="application/pdf"
        )
    except Exception as e:
        st.warning("PDF export requires reportlab. If missing, add to requirements. (" + str(e) + ")")

# ---------------------------
# Formula Playground
# ---------------------------
with tabs[7]:
    st.subheader("Formula Playground")
    st.caption("Evaluate common Excel functions on a dataset and get copy-ready Excel formulas.")

    left, right = st.columns([2, 1])
    with left:
        data_file = st.file_uploader("Upload a dataset (CSV or Excel)", type=["csv", "xlsx", "xlsm", "xls"], key="play_data")
        if data_file:
            df_data = load_user_dataset(data_file)
        else:
            st.info("No dataset uploaded. Using a small sample dataset.")
            df_data = sample_data()
        st.dataframe(df_data.head(20), use_container_width=True, height=350)
    with right:
        table_name = st.text_input("Excel Table name (for formula output)", value="Data")
        func = st.selectbox(
            "Function",
            [
                "SUM", "AVERAGE", "COUNT",
                "SUMIF", "COUNTIF", "SUMIFS", "AVERAGEIF",
                "XLOOKUP", "INDEX+MATCH",
                "FILTER", "UNIQUE", "TEXTJOIN"
            ],
            index=0
        )

    cols = list(df_data.columns)

    def choose_numeric(label="Numeric column"):
        num_cols = [c for c in cols if pd.api.types.is_numeric_dtype(df_data[c])]
        return st.selectbox(label, num_cols) if num_cols else st.selectbox(label, cols)

    def pick_op(label="Operator", include_contains=True, key=None):
        ops = ["equals", "not equals", ">", ">=", "<", "<="]
        if include_contains:
            ops = ["equals", "contains", "not equals", ">", ">=", "<", "<="]
        return st.selectbox(label, ops, key=key)

    result_df = None
    result_scalar = None
    excel_formula = ""

    if func == "SUM":
        target = choose_numeric("Sum column")
        result_scalar = float(df_data[target].sum())
        excel_formula = f"=SUM({structured_ref(table_name, target)})"

    elif func == "AVERAGE":
        target = choose_numeric("Average column")
        result_scalar = float(df_data[target].mean())
        excel_formula = f"=AVERAGE({structured_ref(table_name, target)})"

    elif func == "COUNT":
        target = st.selectbox("Count in column (numbers counted)", cols)
        count_num = pd.to_numeric(df_data[target], errors="coerce").notna().sum()
        result_scalar = int(count_num)
        excel_formula = (
            f"=COUNT({structured_ref(table_name, target)})  "
            f"// or use COUNTA for non-empty: =COUNTA({structured_ref(table_name, target)})"
        )

    elif func == "SUMIF":
        sum_col = choose_numeric("Sum column")
        crit_col = st.selectbox("Criteria column", cols)
        op = pick_op()
        val = st.text_input("Criteria value")
        ser = df_data[crit_col]
        if op == "contains":
            mask = ser.astype(str).str.contains(str(val), case=False, na=False)
        elif op == "equals":
            mask = ser == pd.Series([val]*len(ser), index=ser.index)
        elif op == "not equals":
            mask = ser != pd.Series([val]*len(ser), index=ser.index)
        else:
            mask = pd.to_numeric(ser, errors="coerce").map(lambda x: eval(f"x {op} float(val)") if pd.notna(x) and str(val) != "" else False)
        result_scalar = float(df_data.loc[mask, sum_col].sum())
        excel_formula = f"=SUMIF({structured_ref(table_name, crit_col)}, {excel_criteria(op, val)}, {structured_ref(table_name, sum_col)})"

    elif func == "COUNTIF":
        crit_col = st.selectbox("Criteria column", cols)
        op = pick_op()
        val = st.text_input("Criteria value")
        ser = df_data[crit_col]
        if op == "contains":
            mask = ser.astype(str).str.contains(str(val), case=False, na=False)
        elif op == "equals":
            mask = ser == pd.Series([val]*len(ser), index=ser.index)
        elif op == "not equals":
            mask = ser != pd.Series([val]*len(ser), index=ser.index)
        else:
            mask = pd.to_numeric(ser, errors="coerce").map(lambda x: eval(f"x {op} float(val)") if pd.notna(x) and str(val) != "" else False)
        result_scalar = int(mask.sum())
        excel_formula = f"=COUNTIF({structured_ref(table_name, crit_col)}, {excel_criteria(op, val)})"

    elif func == "SUMIFS":
        sum_col = choose_numeric("Sum column")
        crit1_col = st.selectbox("Criteria 1 column", cols)
        op1 = pick_op()
        val1 = st.text_input("Criteria 1 value")
        add_second = st.checkbox("Add Criteria 2", value=False)
        mask1 = (
            df_data[crit1_col].astype(str).str.contains(str(val1), case=False, na=False)
            if op1 == "contains"
            else (
                df_data[crit1_col] == val1 if op1 == "equals"
                else (df_data[crit1_col] != val1 if op1 == "not equals"
                      else pd.to_numeric(df_data[crit1_col], errors="coerce").map(lambda x: eval(f"x {op1} float(val1)") if pd.notna(x) and str(val1) != "" else False))
            )
        )
        mask = mask1
        excel_formula = f"=SUMIFS({structured_ref(table_name, sum_col)}, {structured_ref(table_name, crit1_col)}, {excel_criteria(op1, val1)}"
        if add_second:
            crit2_col = st.selectbox("Criteria 2 column", cols)
            op2 = pick_op(key="op2")
            val2 = st.text_input("Criteria 2 value", key="val2")
            mask2 = (
                df_data[crit2_col].astype(str).str.contains(str(val2), case=False, na=False)
                if op2 == "contains"
                else (
                    df_data[crit2_col] == val2 if op2 == "equals"
                    else (df_data[crit2_col] != val2 if op2 == "not equals"
                          else pd.to_numeric(df_data[crit2_col], errors="coerce").map(lambda x: eval(f"x {op2} float(val2)") if pd.notna(x) and str(val2) != "" else False))
                )
            )
            mask = mask & mask2
            excel_formula += f", {structured_ref(table_name, crit2_col)}, {excel_criteria(op2, val2)}"
        excel_formula += ")"
        result_scalar = float(df_data.loc[mask, sum_col].sum())

    elif func == "AVERAGEIF":
        avg_col = choose_numeric("Average column")
        crit_col = st.selectbox("Criteria column", cols)
        op = pick_op()
        val = st.text_input("Criteria value")
        ser = df_data[crit_col]
        if op == "contains":
            mask = ser.astype(str).str.contains(str(val), case=False, na=False)
        elif op == "equals":
            mask = ser == pd.Series([val]*len(ser), index=ser.index)
        elif op == "not equals":
            mask = ser != pd.Series([val]*len(ser), index=ser.index)
        else:
            mask = pd.to_numeric(ser, errors="coerce").map(lambda x: eval(f"x {op} float(val)") if pd.notna(x) and str(val) != "" else False)
        result_scalar = float(df_data.loc[mask, avg_col].mean())
        excel_formula = f"=AVERAGEIF({structured_ref(table_name, crit_col)}, {excel_criteria(op, val)}, {structured_ref(table_name, avg_col)})"

    elif func == "XLOOKUP":
        key_col = st.selectbox("Lookup column (search in)", cols)
        ret_col = st.selectbox("Return column", cols, index=min(1, len(cols)-1))
        lookup_val = st.text_input("Lookup value")
        hit = df_data[df_data[key_col].astype(str) == str(lookup_val)]
        result_scalar = (None if hit.empty else hit.iloc[0][ret_col])
        excel_formula = f"=XLOOKUP({excel_literal(lookup_val)}, {structured_ref(table_name, key_col)}, {structured_ref(table_name, ret_col)}, \"Not found\")"

    elif func == "INDEX+MATCH":
        key_col = st.selectbox("Lookup column (MATCH against)", cols)
        ret_col = st.selectbox("Return column (INDEX from)", cols, index=min(1, len(cols)-1))
        lookup_val = st.text_input("Lookup value")
        hit = df_data[df_data[key_col].astype(str) == str(lookup_val)]
        result_scalar = (None if hit.empty else hit.iloc[0][ret_col])
        excel_formula = (
            f"=INDEX({structured_ref(table_name, ret_col)}, "
            f"MATCH({excel_literal(lookup_val)}, {structured_ref(table_name, key_col)}, 0))"
        )

    elif func == "FILTER":
        ret_cols = st.multiselect("Return columns", cols, default=cols[:1])
        crit_col = st.selectbox("Filter column", cols)
        op = pick_op()
        val = st.text_input("Filter value")
        ser = df_data[crit_col]
        if op == "contains":
            mask = ser.astype(str).str.contains(str(val), case=False, na=False)
        elif op == "equals":
            mask = ser == pd.Series([val]*len(ser), index=ser.index)
        elif op == "not equals":
            mask = ser != pd.Series([val]*len(ser), index=ser.index)
        else:
            mask = pd.to_numeric(ser, errors="coerce").map(lambda x: eval(f"x {op} float(val)") if pd.notna(x) and str(val) != "" else False)
        result_df = df_data.loc[mask, ret_cols]
        if len(ret_cols) == 1:
            ret_ref = structured_ref(table_name, ret_cols[0])
        else:
            ret_ref = structured_ref_multi(table_name, ret_cols)
        excel_formula = f"=FILTER({ret_ref}, {structured_ref(table_name, crit_col)}={excel_criteria(op, val) if op!='contains' else excel_criteria(op, val)})"
        if op == "contains":
            excel_formula += "  // Note: Excel FILTER cannot natively do contains without helper. Consider SEARCH()>0."

    elif func == "UNIQUE":
        target = st.selectbox("Column", cols)
        result_df = pd.DataFrame({target: df_data[target].dropna().astype(str).unique()})
        excel_formula = f"=UNIQUE({structured_ref(table_name, target)})"

    elif func == "TEXTJOIN":
        target = st.selectbox("Column", cols)
        delim = st.text_input("Delimiter", value=", ")
        ignore_empty = st.checkbox("Ignore empty", value=True)
        series = df_data[target].astype(str)
        if ignore_empty:
            series = series.replace({"": np.nan}).dropna()
        result_scalar = delim.join(series.astype(str).tolist())
        excel_formula = f"=TEXTJOIN({excel_literal(delim)}, {excel_literal(ignore_empty)}, {structured_ref(table_name, target)})"

    st.markdown("---")
    if result_df is not None:
        st.write("Result (table):")
        st.dataframe(result_df, use_container_width=True)
    elif result_scalar is not None:
        st.metric("Result", f"{result_scalar}")
    else:
        st.info("Adjust parameters to see results.")

    if excel_formula:
        st.markdown("Copy-ready Excel formula:")
        st.code(excel_formula)

# ---------- Footer (ASCII-only) ----------
st.markdown("---")
st.caption("(c) Your Team - Built for speed and clarity - Powered by Streamlit")
