import io
from datetime import datetime
from typing import Dict, List, Optional

import pandas as pd
import streamlit as st

# ---------------------------
# Page Config
# ---------------------------
st.set_page_config(
    page_title="Excel Formula Companion",
    page_icon="🧮",
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
            # Normalize column names
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
    # Category filter
    if category_col and category_col in out.columns and category_values:
        out = out[out[category_col].isin(category_values)]
    # Text search across multiple columns
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
            # Excel sheet name limit 31 chars
            safe_name = sheet_name[:31] if sheet_name else "Sheet"
            df.to_excel(writer, index=False, sheet_name=safe_name)
            # Autofit
            ws = writer.sheets[safe_name]
            for idx, col in enumerate(df.columns):
                width = max(10, min(60, int(df[col].astype(str).str.len().quantile(0.9)) + 4))
                ws.set_column(idx, idx, width)
    return buf.getvalue()

def to_pdf_bytes(title: str, sections: Dict[str, pd.DataFrame]) -> bytes:
    # Lightweight PDF export using reportlab (summary style)
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
        c.drawString(2*cm, H-2.6*cm, f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")

    y = H - 3.0*cm
    draw_header(title)

    for section, df in sections.items():
        c.setFont("Helvetica-Bold", 11)
        c.drawString(2*cm, y, section)
        y -= 0.5*cm

        # Render first ~20 rows with 3-5 columns best effort
        c.setFont("Helvetica", 8)
        if not df.empty:
            cols = list(df.columns)[:5]
            # Column headers
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
        f"""
        <span style="
            display:inline-block;
            padding:4px 10px;
            border-radius:999px;
            border:1px solid #e5e7eb;
            background:#f8fafc;
            font-size:12px;
            margin-right:6px;">
            {text}
        </span>
        """,
        unsafe_allow_html=True
    )

# ---------------------------
# UI: Sidebar
# ---------------------------
st.title("🧮 Excel Formula Companion")
st.caption("Search • Filter • Compare • Learn • Export")

with st.sidebar:
    st.header("Step 1 — Upload Master Workbook")
    uploaded = st.file_uploader(
        "Upload your Excel master file",
        type=["xlsx", "xlsm", "xls"],
        accept_multiple_files=False
    )

    st.markdown("---")
    st.subheader("Global Search & Filters")

    search_term = st.text_input("Search text (function, description, syntax, hints)", "")
    selected_categories = []
    selected_functions = []

# ---------------------------
# Load Data
# ---------------------------
if not uploaded:
    st.info("Upload your **Excel_Formulas_Master_Reference_With_Filters.xlsx** to begin.")
    st.stop()

sheets = load_workbook(uploaded.read())

# Expected sheet names (from your file)
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

# Derive category/function options from Full Reference (fallback to Formulas List)
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
    "Exports"
])

# ---------------------------
# Full Reference Tab
# ---------------------------
with tabs[0]:
    st.subheader("Full Reference")
    if full_ref.empty:
        st.warning("No data found in 'Full Reference'.")
    else:
        base_df = full_ref.copy()

        # Apply filters
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

        # Quick details panel
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
                st.write("**Category:**", row.get("Category", "—"))
                st.write("**Short Description:**", row.get("Short Description", "—"))
                st.write("**Syntax Template:**")
                st.code(str(row.get("Syntax Template", "")))
                st.write("**Example Formula (copy/paste):**")
                st.code(str(row.get("Example Formula (copy/paste)", "")))
                st.write("**Hints:**", row.get("Hints", "—"))
                # Chips
                chip(row.get("Category", ""))
                chip(pick)

# ---------------------------
# Filtered View Tab
# ---------------------------
with tabs[1]:
    st.subheader("Filtered View")
    if filtered_view.empty and full_ref.empty:
        st.warning("No data found in 'Filtered View' (using Full Reference instead).")
        df_base = full_ref.copy()
    else:
        # If your workbook drives this view, still allow global filters
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
        st.info("No data in 'Usage Cheat-Sheet'.")
    else:
        st.dataframe(cheats, use_container_width=True)

        # Friendly expanders by topic if columns exist
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
        st.info("No data in 'Category Index'.")
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
        st.info("No data in 'Quick Examples'.")
    else:
        st.dataframe(quick, use_container_width=True)
        # Optional: per-function quick example lookup
        fn_col = "Function" if "Function" in quick.columns else None
        if fn_col:
            pick2 = st.selectbox("Show example for function", sorted(quick[fn_col].dropna().unique()))
            ex_df = quick[quick[fn_col] == pick2]
            st.write(ex_df)

# ---------------------------
# Flashcards (Learning Mode)
# ---------------------------
with tabs[5]:
    st.subheader("Flashcards (Learning Mode)")
    if full_ref.empty:
        st.info("Need 'Full Reference' to use flashcards.")
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
            idx = st.slider("Flashcard index", 1, len(pool), 1, key="flash_idx")
            row = pool.iloc[idx-1].to_dict()
            st.markdown(f"### {row.get('Function', 'Function')}")
            st.write("**Category:**", row.get("Category", "—"))
            st.write("**Short Description:**", row.get("Short Description", "—"))
            with st.expander("Show Syntax"):
                st.code(str(row.get("Syntax Template", "")))
            with st.expander("Show Example"):
                st.code(str(row.get("Example Formula (copy/paste)", "")))
            with st.expander("Hints"):
                st.write(row.get("Hints", "—"))

# ---------------------------
# Exports Tab
# ---------------------------
with tabs[6]:
    st.subheader("Exports")
    st.caption("Export your current, filtered views.")

    # Build exportable slices
    export_slices = {}

    # Full Reference (filtered the same way as tab 1)
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

    # Excel export
    excel_bytes = to_excel_bytes(export_slices)
    st.download_button(
        label="⬇️ Download Excel (.xlsx)",
        data=excel_bytes,
        file_name=f"Excel_Formula_Companion_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # PDF export (summary)
    try:
        pdf_bytes = to_pdf_bytes(
            "Excel Formula Companion — Export",
            export_slices
        )
        st.download_button(
            label="⬇️ Download PDF (summary)",
            data=pdf_bytes,
            file_name=f"Excel_Formula_Companion_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
            mime="application/pdf"
        )
    except Exception as e:
        st.warning(f"PDF export requires 'reportlab'. If missing, add to requirements. ({e})")

st.markdown("---")
st.caption("© Your Team • Built for speed and clarity • Streamlit")
