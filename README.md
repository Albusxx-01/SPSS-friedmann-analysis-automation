# Friedman Analysis Pipeline

A full-automation replicate of the **SPSS K-Related + Two Related Samples**
workflow for medical-imaging model comparisons:

- **Friedman test** across models (Q, df, Asymp.Sig. p, Kendall's W, decision vs alpha)
- **Pairwise Wilcoxon signed-rank** post-hoc with **Bonferroni** or **Benjamini–Hochberg** correction (or none)
- Formatted **Excel workbooks** (`Ranks` + `Test Statistics`), per-metric **CSVs**, SPSS-style **.txt** tables, **plots**, and a **Markdown** report

Two interfaces over the same engine (`friedmann_pipeline.py`):

| Interface | Entry point | When to use |
|---|---|---|
| Web UI | `streamlit run app.py` | Interactive upload/detect/select/download |
| CLI | `python friedmann_pipeline.py ...` | Scripted / headless batch runs |

Both produce **byte-identical statistical outputs** — the UI calls the very same
`configure()` + `run_analysis()` functions as the CLI.

---

## Install

Python 3.9+ (tested on 3.13).

```bash
pip install -r requirements.txt
```

Requires: `pandas`, `numpy`, `scipy`, `matplotlib`, `seaborn`, `openpyxl`, `streamlit`.

---

## Quick start

### Web UI

```bash
streamlit run app.py
```

Then open <http://localhost:8501>:

1. **Upload** one or more `.xlsx`/`.csv` files.
2. *(Optional)* tune **alpha / correction / sheet** or re-enable blacklisted
   columns (`--include-metric`) in the sidebar.
3. Click **Detect metrics** → check the per-file availability matrix.
4. **Select** the metrics to analyze.
5. Click **Run analysis** → review the summary table + run log.
6. **Download**: Case 1 final workbooks only, or Case 2 (everything, zipped).

Press **📖 How to use** in the sidebar for the full guide.

### CLI

```bash
# Run everything in the default Datasets/ folder (one level above this script)
python friedmann_pipeline.py

# Specific files
python friedmann_pipeline.py --input file.xlsx file2.csv

# Tune the analysis
python friedmann_pipeline.py --alpha 0.01 --correction bh --clean

# Force a blacklisted column to be treated as a metric
python friedmann_pipeline.py --include-metric "No. of Parameters (M)"
```

---

## Input format

Two layouts are **auto-detected** from the column names:

- **MS long format** — `Number` (model id) + `Run Number` (iteration) plus one
  column per metric; one row per (model, run). The table is pivoted to one row
  per model per metric.
- **Wide format** — one column per model/group, one row per subject/block;
  used as-is.

Rules:

- Metric columns must be **numeric**. Text values like `85.2%` are coerced
  automatically.
- Identifier/constant columns (`Number`, `Run Number`, `No. of Parameters`,
  `MACs`, `Model`, …) are **excluded** unless re-enabled via `--include-metric`.
- Rows with missing values are **dropped** automatically (count reported in the
  log and per-metric `*_descriptives.csv`).
- **Lower-is-better** metrics (`FPR`, `Inference Time`, `MAE`, `MSE`, `RMSE`,
  `HD/SSD/ASSD`, `Loss`, `Latency`, …) are detected by name and ranked accordingly.
- Views are detected from the file stem (`Axial`, `Sagital`/`Sagittal`, else the
  stem is used).

---

## CLI options

```
usage: friedmann_pipeline.py [-h] [--input [INPUT ...]] [--alpha ALPHA]
                             [--out-dir OUT_DIR] [--final-dir FINAL_DIR]
                             [--dataset-dir DATASET_DIR] [--sheet SHEET]
                             [--correction {bonferroni,bh,none}]
                             [--include-metric [INCLUDE_METRIC ...]] [--clean]
```

| Flag | Description | Default |
|---|---|---|
| `--input [PATH...]` | Input file(s); omit for the interactive prompt / default `Datasets/` | prompt |
| `--alpha ALPHA` | Significance level | `0.05` |
| `--out-dir DIR` | Per-metric CSVs, txt, plots, summary | `outputs/` (script-local) |
| `--final-dir DIR` | Final Excel workbooks | `final_results/` (script-local) |
| `--dataset-dir DIR` | Default `Datasets/` when no `--input` is given | `../Datasets` |
| `--sheet SHEET` | Excel sheet index or name | `0` (first) |
| `--correction {bonferroni,bh,none}` | Post-hoc correction | `bonferroni` |
| `--include-metric [NAME...]` | Force blacklisted columns to be metrics | — |
| `--clean` | Wipe all files in output dirs first | off |

> Paths are **script-local** (`SCRIPT_DIR`) by default, so the pipeline works
> regardless of the current working directory.

---

## Outputs

### Per metric (in `outputs/`)

```
outputs/
├── <Metric>_ranks.csv              Friedman mean ranks per view
├── <Metric>_test_statistics.csv    Q, df, p, Kendall W, decision
├── <Metric>/<view>_statistics.txt  SPSS-style .txt tables (ranks + test stat)
├── <Metric>_plot.png               Mean ranks bar chart with p/Q/W annotations
└── summary.csv                     All metrics on one table (per view)
```

A `descriptives.csv` is written for every metric; a `.md` report summarises the
whole run.

### Final workbooks (in `final_results/`)

```
final_results/
├── Ranks_<View>_All_Metrics.xlsx          SPSS-style Ranks sheet
└── TestStatistics_<View>_All_Metrics.xlsx Friedman + post-hoc stats, green-filled
                                          for significant pairwise p-values
```

A file carries the `All_Metrics` suffix only in **Case 1** (all metrics selected).
When a single metric is analyzed it is named after that metric instead.

---

## Decisions

Per metric, the pipeline emits one of:

- **SIGNIFICANT** — Friedman p ≤ alpha; post-hoc computed and significant pairs
  reported, or warning when Wilcoxon fails (ties/dedup).
- **NOT SIGNIFICANT** — p > alpha; note that pairwise comparisons are
  distribution-equality checks and are still reported.
- **INCONCLUSIVE** — non-finite Friedman p (e.g. all rows equal).

---

## Project layout

```
├── friedmann_pipeline.py   # engine: configure() + run_analysis(), CLI
├── app.py                  # Streamlit UI (same engine)
├── .streamlit/config.toml  # UI theme
├── requirements.txt
└── README.md
```

Note: `outputs/` and `final_results/` are created next to the script at runtime.

---

## Verified against the reference workflow

Results match the baseline in the original
`../Friedman_Analysis/WORKFLOW.md` (e.g. Axial Accuracy:
Q = 291.314, p = 2.93e-51, Kendall's W = 0.539, 73 significant pairs of 171).# SPSS-friedmann-analysis-automation
