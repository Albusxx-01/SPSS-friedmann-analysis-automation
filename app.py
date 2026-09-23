"""Streamlit UI for the Friedman analysis pipeline.

Reuses the engine in `friedmann_pipeline.py` via `configure()` + `run_analysis()`
so the web UI and the CLI produce identical outputs.

Run:  streamlit run app.py

Flow:
  1. Upload input files (drag & drop .xlsx/.csv).
  2. Configure analysis options in the sidebar (alpha, correction, sheet,
     --include-metric to re-enable blacklisted columns).
  3. Detect metrics -> combined multiselect with a per-file availability matrix.
  4. Run -> summary table preview -> deliver outputs:
       Case 1: "All Metrics files" only (final workbooks) - per-file buttons + zip.
       Case 2: everything (all outputs/ + final_results/ in one zip).
"""
import glob
import io
import os
import shutil
import tempfile
import zipfile
from contextlib import redirect_stdout

import pandas as pd
import streamlit as st

import friedmann_pipeline as fp

st.set_page_config(page_title="Friedman Analysis", page_icon="📊", layout="wide")

XLSX_MIME = ("application/vnd.openxmlformats-officedocument"
             ".spreadsheetml.sheet")


# ---------------------------------------------------------------------------
# look & feel
# ---------------------------------------------------------------------------
def inject_css() -> None:
    st.markdown(
        """
        <style>
        .stApp { background: linear-gradient(180deg, #fbfbfe 0%, #eef1fa 100%); }

        /* side bar - dark contrasted panel */
        [data-testid="stSidebar"] {
            background: linear-gradient(180deg, #232c4e 0%, #2e3a63 100%);
            border-right: 1px solid #12182b;
        }
        [data-testid="stSidebar"] h1,
        [data-testid="stSidebar"] h2,
        [data-testid="stSidebar"] h3,
        [data-testid="stSidebar"] h4 { color: #ffffff; }
        [data-testid="stSidebar"] p,
        [data-testid="stSidebar"] li,
        [data-testid="stSidebar"] span,
        [data-testid="stSidebar"] label { color: #d9e1f5; }
        [data-testid="stSidebar"] code { background: #3a4670; color: #ffe9a8; }
        [data-testid="stSidebar"] [data-testid="stExpander"] {
            border: 1px solid rgba(255,255,255,.18); border-radius: 12px;
        }
        [data-testid="stSidebar"] .howto ul { padding-left: 16px; }
        [data-testid="stSidebar"] .stButton > button {
            background: rgba(255,255,255,.14); color: #ffffff;
            border: 1px solid rgba(255,255,255,.35); border-radius: 10px;
        }
        [data-testid="stSidebar"] .stButton > button:hover {
            background: rgba(255,255,255,.26); color: #ffffff;
        }

        /* hero banner */
        .hero {
            background: linear-gradient(120deg, #4f46e5 0%, #7c3aed 55%, #a855f7 100%);
            border-radius: 18px; padding: 22px 28px; margin: 4px 0 10px; color: #fff;
            box-shadow: 0 8px 26px rgba(79,70,229,.35);
        }
        .hero h1 { margin: 0 0 4px; color: #fff; font-size: 1.9rem; font-weight: 800; }
        .hero p { margin: 0; color: #ede9fe; font-size: 1rem; }

        /* numbered section badges */
        .sec { display: flex; align-items: center; gap: 12px; margin: 22px 0 10px; }
        .sec .num {
            flex: 0 0 auto; width: 30px; height: 30px; border-radius: 50%;
            background: linear-gradient(135deg, #5b5bd6, #a855f7);
            color: #fff; font-weight: 800; display: flex; align-items: center;
            justify-content: center; box-shadow: 0 3px 8px rgba(124,58,237,.4);
        }
        .sec h2 { margin: 0; color: #1f2a44; }

        /* quick-start chips */
        .chip {
            background: linear-gradient(135deg, #ffffff, #f3f0ff);
            border: 1px solid #e4e8f3; border-top: 3px solid #5b5bd6;
            border-radius: 12px; padding: 12px 14px; text-align: center;
            box-shadow: 0 2px 8px rgba(31,42,68,.06);
        }
        .chip b { color: #5b5bd6; }

        /* cards */
        .card {
            background: #ffffff; border: 1px solid #e4e8f3; border-radius: 14px;
            padding: 14px 18px; margin: 10px 0;
            box-shadow: 0 2px 8px rgba(31,42,68,.06);
        }

        /* buttons */
        .stButton > button[kind="primary"],
        .stDownloadButton > button {
            background: linear-gradient(135deg, #5b5bd6 0%, #8a5bf6 100%);
            color: #ffffff; border: none; border-radius: 10px; font-weight: 600;
            padding: .5rem 1.1rem;
            box-shadow: 0 3px 10px rgba(91,91,214,.35);
            transition: transform .05s ease, box-shadow .15s ease;
        }
        .stButton > button[kind="primary"]:hover,
        .stDownloadButton > button:hover {
            box-shadow: 0 5px 16px rgba(91,91,214,.5); transform: translateY(-1px);
        }
        .stButton > button[kind="secondary"] { border-radius: 10px; }

        /* misc */
        [data-testid="stFileUploaderDropzone"] {
            border-radius: 14px; border: 2px dashed #5b5bd6;
            background: rgba(91,91,214,.05);
        }
        [data-testid="stAlert"] {
            border-radius: 12px; box-shadow: 0 2px 10px rgba(31,42,68,.08);
        }
        [data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def hero() -> None:
    st.markdown(
        '<div class="hero"><h1>📊 Friedman Analysis Pipeline</h1>'
        "<p>SPSS K-Related Samples + Two Related Samples replica — "
        "Friedman + Bonferroni/BH Wilcoxon post-hoc.</p></div>",
        unsafe_allow_html=True,
    )


def section(num: str, title: str) -> None:
    st.markdown(
        f'<div class="sec"><div class="num">{num}</div><h2>{title}</h2></div>',
        unsafe_allow_html=True,
    )


# How-to content (rendered in the main panel when the sidebar button is pressed)
HOWTO_HTML = (
    '<div class="howto">\n'
    "#### 1. Input file format\n\n"
    "Upload **`.xlsx`** or **`.csv`**. Two layouts are auto-detected:\n\n"
    "• **MS long format** — `Number` (model id) + `Run Number` "
    "(iteration) + one column **per metric**; one row per (model, run).\n\n"
    "• **Wide format** — one column **per model/group**, one row per "
    "subject/block (used as-is).\n\n"
    "Rules:\n"
    "- Metric columns must be **numeric** (text like `85.2%` is coerced\n"
    "  automatically).\n"
    '- Identifier/constant columns (`Number`, `Run Number`, `No. of\n'
    "  Parameters`, `MACs`) are excluded by default — re-enable any of\n"
    "  them with **--include-metric** below.\n"
    "- Rows with missing values are **dropped** automatically (see log).\n"
    "- Lower-is-better metrics (`FPR`, `Inference Time`, `MAE`, `HD/SSD`,\n"
    "  ...) are detected by name.\n\n"
    "#### 2. Steps\n\n"
    "1. **Upload** one or more `.xlsx`/`.csv` files.\n"
    "2. *(Optional)* tweak **options below** (alpha / correction / sheet).\n"
    "3. Click **“Detect metrics”**.\n"
    "4. **Select** the metric(s) to analyze (metrics marked “—” for a\n"
    "   file are skipped for that file).\n"
    "5. Click **“Run analysis”**.\n"
    "6. Review the **summary table** + **log**, then download.\n\n"
    "#### 3. Download\n\n"
    "• **Case 1 — All Metrics files only**: final workbooks "
    "(`Ranks_*_All_Metrics.xlsx`, `TestStatistics_*_All_Metrics.xlsx`) "
    "as per-file buttons or one zip.\n"
    "• **Case 2 — All output (everything)**: one zip with `outputs/` "
    "(CSVs, SPSS-style `.txt`, summary, markdown) **and** "
    "`final_results/` (the workbooks).\n\n"
    "</div>"
)


inject_css()
hero()

# ---------------------------------------------------------------------------
# sidebar: how to use + analysis options
# ---------------------------------------------------------------------------
with st.sidebar:
    if st.button("📖 How to use", key="howto_btn", use_container_width=True):
        st.session_state["howto_show"] = True

st.sidebar.header("⚙️ Analysis options")
alpha = st.sidebar.number_input("Alpha (significance level)",
                                min_value=0.001, max_value=1.0,
                                value=fp.DEFAULT_ALPHA, step=0.005, key="alpha")
correction = st.sidebar.selectbox("Post-hoc correction",
                                  ["bonferroni", "bh", "none"], key="correction")
sheet = st.sidebar.text_input("Excel sheet (index or name)", "0", key="sheet")

# ---------------------------------------------------------------------------
# one stable scratch dir per browser session (outputs never touch CLI dirs)
# ---------------------------------------------------------------------------
if "run_dir" not in st.session_state:
    st.session_state.run_dir = tempfile.mkdtemp(prefix="friedman_ui_")
RUN_DIR = st.session_state.run_dir

# ---------------------------------------------------------------------------
# quick-start chips
# ---------------------------------------------------------------------------
c1, c2, c3 = st.columns(3)
c1.markdown('<div class="chip">1️⃣ <b>Upload</b> your `.xlsx`/`.csv` files</div>',
            unsafe_allow_html=True)
c2.markdown('<div class="chip">2️⃣ <b>Detect</b> & select the metrics</div>',
            unsafe_allow_html=True)
c3.markdown('<div class="chip">3️⃣ <b>Run</b> & download the outputs</div>',
            unsafe_allow_html=True)

# how-to content opens here in the main panel (sidebar button toggles it)
if st.session_state.get("howto_show", False):
    st.markdown(
        '<div class="sec"><div class="num">📖</div><h2>How to use</h2></div>',
        unsafe_allow_html=True)
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown(HOWTO_HTML, unsafe_allow_html=True)
    if st.button("✖ Close", key="howto_close", type="secondary"):
        st.session_state["howto_show"] = False
    st.markdown("</div>", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# 1. inputs
# ---------------------------------------------------------------------------
section("1", "Input files")
with st.container():
    up = st.file_uploader("Drop .xlsx/.csv files",
                          type=["xlsx", "csv"], accept_multiple_files=True,
                          key="uploader")
    up_dir = st.session_state.setdefault(
        "upload_dir", tempfile.mkdtemp(prefix="friedman_upload_"))
    input_files = []
    seen = set()
    for u in up:
        ident = (u.name, u.size)
        if ident not in seen:
            seen.add(ident)
            p = os.path.join(up_dir, u.name)
            with open(p, "wb") as f:
                f.write(u.getbuffer())
        else:
            p = os.path.join(up_dir, u.name)
        input_files.append(os.path.abspath(p))

# blacklisted columns actually present in the chosen files (for --include-metric)
if input_files:
    fp.configure(alpha=alpha, sheet=sheet, correction=correction,
                 include_metric=[], clean=False)
    blocked_present = sorted(
        {c for f in input_files for c in fp.NON_METRIC
         if c in fp._metric_ready_df(f).columns})
else:
    blocked_present = []
include_metric = st.sidebar.multiselect(
    "--include-metric (re-enable blacklisted columns)", blocked_present,
    key="include_metric")

# ---------------------------------------------------------------------------
# 2. detect + select metrics
# ---------------------------------------------------------------------------
section("2", "Select metrics")
if st.button("🔍 Detect metrics", type="primary", key="detect_btn"):
    if not input_files:
        st.error("Choose at least one input file first.")
    else:
        per_file = {}
        union = set()
        with st.spinner("Detecting metrics and models..."):
            for p in input_files:
                try:
                    ms = fp.detect_metrics(p)
                except Exception as e:  # noqa: BLE001 - surface to user
                    st.error(f"{os.path.basename(p)}: {e}")
                    ms = []
                per_file[p] = ms
                union |= set(ms)
        st.session_state["per_file"] = per_file
        st.session_state["union"] = sorted(union)
        st.success(f"Detected {len(union)} metric(s) across "
                   f"{len(input_files)} file(s).")

per_file = st.session_state.get("per_file", {})
union = st.session_state.get("union", [])
if not union:
    st.info("👆 Upload files above, then click **Detect metrics** to populate "
            "the list of available metrics.")
else:
    available_metrics = st.multiselect("Metrics to analyze", union,
                                       key="metrics_sel")
    if available_metrics and per_file:
        matrix = pd.DataFrame(
            {os.path.basename(p): [(("✅" if m in ms else "—"))
                                    for m in available_metrics]
             for p, ms in per_file.items()},
            index=available_metrics)
        st.caption("Per-file availability — metrics marked '—' are skipped for "
                   "that file (warning shown in the run log).")
        st.dataframe(matrix, width="stretch")

# ---------------------------------------------------------------------------
# 3. run
# ---------------------------------------------------------------------------
section("3", "Run analysis")
if st.button("▶️ Run analysis", type="primary", key="run_btn"):
    selected = st.session_state.get("metrics_sel", [])
    if not selected:
        st.error("Select at least one metric.")
    elif not input_files:
        st.error("Choose at least one input file first.")
    else:
        # fresh output dirs inside the session scratch dir
        for sub in ("outputs", "final_results"):
            shutil.rmtree(os.path.join(RUN_DIR, sub), ignore_errors=True)
        out_dir = os.path.join(RUN_DIR, "outputs")
        final_dir = os.path.join(RUN_DIR, "final_results")
        fp.configure(alpha=alpha, out_dir=out_dir, final_dir=final_dir,
                     sheet=sheet, correction=correction,
                     include_metric=list(include_metric), clean=False)
        log_buf = io.StringIO()
        try:
            with st.spinner("Running Friedman + Wilcoxon post-hoc..."):
                with redirect_stdout(log_buf):
                    results, summary, skipped = fp.run_analysis(input_files,
                                                                set(selected))
        except SystemExit as e:
            st.session_state["run_log"] = log_buf.getvalue()
            st.error(str(e))
        else:
            st.session_state["meta"] = {
                "out_dir": out_dir, "final_dir": final_dir,
                "log": log_buf.getvalue(), "summary": summary,
                "results": results, "n_files": len(input_files),
            }
            for path, metric, reason in skipped:
                if metric is None:
                    st.warning(f"{os.path.basename(path)}: {reason}.")
                else:
                    st.warning(f"{os.path.basename(path)}: metric "
                               f"'{metric}' {reason} — skipped.")
            st.success(f"✅ Analysis complete: {len(summary)} run(s) across "
                       f"{len({r['view'] for r in results})} view(s).")

# ---------------------------------------------------------------------------
# 4. results preview + 5. downloads
# ---------------------------------------------------------------------------
meta = st.session_state.get("meta")
if meta:
    section("4", "Results")
    st.subheader("Summary table")
    st.dataframe(meta["summary"], width="stretch", hide_index=True)
    with st.expander("Run log"):
        st.code(meta["log"])

    section("5", "Download")
    st.caption("Case 1 = final 'All Metrics' workbooks only. "
               "Case 2 = every output file (CSVs, SPSS-style txt, "
               "summary, workbooks) in one zip.")
    out_mode = st.radio("Output scope",
                        ["Case 1: All Metrics files only",
                         "Case 2: All output (everything)"],
                        key="out_mode", horizontal=True)

    workbooks = sorted(glob.glob(os.path.join(meta["final_dir"], "*.xlsx")))

    def _zip_bytes(items):
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            for arcname, path in items:
                z.write(path, arcname)
        return buf.getvalue()

    if out_mode.startswith("Case 1"):
        if not workbooks:
            st.info("No final workbooks produced for the selected metrics.")
        else:
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.write("**Final workbooks**")
            for wb in workbooks:
                with open(wb, "rb") as f:
                    st.download_button(f"⬇️ {os.path.basename(wb)}", f.read(),
                                       file_name=os.path.basename(wb),
                                       mime=XLSX_MIME,
                                       key=f"dl_{os.path.basename(wb)}")
            if len(workbooks) > 1:
                items = [(os.path.join("final_results",
                                       os.path.basename(p)), p)
                         for p in workbooks]
                st.download_button("⬇️ All Metrics (zip)", _zip_bytes(items),
                                   file_name="All_Metrics_files.zip",
                                   mime="application/zip",
                                   key="dl_allmetrics_zip")
            st.markdown("</div>", unsafe_allow_html=True)
    else:
        items = ([(os.path.join("outputs", os.path.basename(p)), p)
                  for p in glob.glob(os.path.join(meta["out_dir"], "*"))] +
                 [(os.path.join("final_results", os.path.basename(p)), p)
                  for p in glob.glob(os.path.join(meta["final_dir"], "*"))])
        if items:
            st.download_button("⬇️ Download everything (.zip)",
                               _zip_bytes(items),
                               file_name="Friedman_all_outputs.zip",
                               mime="application/zip",
                               key="dl_everything_zip")
        with st.expander("Zip contents"):
            for arcname, _ in items:
                st.code(arcname)