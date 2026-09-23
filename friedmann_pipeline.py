"""Full-automation Friedman analysis pipeline.

One command flow: provide .xlsx/.csv input file path(s) -> auto-detect
views/metrics -> Friedman + Wilcoxon post-hoc for every metric ->
per-metric CSVs, SPSS-style txt, plots, summary CSV, overall markdown,
and final formatted Excel workbooks (Ranks + Test Statistics).

Replicates SPSS K-Related Samples + Two Related Samples flow (video replica):
  - Overall Friedman: Q, df, Asymp.Sig. p, Kendall W, decision vs alpha
  - Post-hoc correction: Bonferroni (alpha/m, m = k*(k-1)/2),
    Benjamini-Hochberg FDR, or none (--correction)
  - Pairwise Wilcoxon signed-rank: Neg/Pos/Ties + adjusted p vs alpha

Supports both input layouts:
  - MS long format (Number + Run Number columns, one row per run/model) -> pivot
  - Any wide format (one column per group, one row per subject/block) -> as-is

Run:  py friedmann_pipeline.py                      # prompts / Datasets/
      py friedmann_pipeline.py --input file.xlsx file2.csv
      py friedmann_pipeline.py --alpha 0.01 --correction bh --clean
      py friedmann_pipeline.py --include-metric "No. of Parameters (M)"

Defaults are script-local: Datasets/ one level up, outputs/ and
final_results/ next to this script (override with --dataset-dir/--out-dir/--final-dir).
"""
import argparse
import glob
import os
import re
import sys
import warnings
from collections import OrderedDict
from itertools import combinations
from types import SimpleNamespace

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
from openpyxl.styles import Alignment, Font, PatternFill
from scipy.stats import friedmanchisquare, wilcoxon

sns.set(style="whitegrid")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_DATASET_DIR = os.path.abspath(os.path.join(SCRIPT_DIR, "..", "Datasets"))
DEFAULT_OUT_DIR = os.path.join(SCRIPT_DIR, "outputs")
DEFAULT_FINAL_DIR = os.path.join(SCRIPT_DIR, "final_results")
DEFAULT_ALPHA = 0.05

# Runtime configuration; populated by main(). Kept module-level so helpers can
# read alpha / correction / sheet / dirs without threading every signature.
CFG = SimpleNamespace(
    alpha=DEFAULT_ALPHA,
    out_dir=DEFAULT_OUT_DIR,
    final_dir=DEFAULT_FINAL_DIR,
    dataset_dir=DEFAULT_DATASET_DIR,
    sheet=0,
    correction="bonferroni",
    include_metric=set(),
    clean=False,
)

VIEW_PATTERNS = [
    (re.compile(r"axial", re.I), "Axial"),
    (re.compile(r"sagitt?al", re.I), "Sagittal"),  # both 'Sagital' and 'Sagittal'
]

# Columns that are identifiers / constants, never metrics unless the user
# explicitly re-enables them via --include-metric.
NON_METRIC = {"Number", "Run Number", "No. of Parameters", "MACs",
              "No. of Parameters (M)", "MACs (G)", "Model", "Model Number"}

# Name patterns defining lower-is-better metrics (lowest mean rank = best).
LOWER_BETTER_PATTERNS = [
    re.compile(r"fpr", re.I),
    re.compile(r"inference", re.I),
    re.compile(r"error", re.I),
    re.compile(r"loss", re.I),
    re.compile(r"mae", re.I),
    re.compile(r"mse", re.I),
    re.compile(r"rmse", re.I),
    re.compile(r"hd", re.I),
    re.compile(r"hausdorff", re.I),
    re.compile(r"ssd", re.I),
    re.compile(r"assd", re.I),
    re.compile(r"latency", re.I),
    re.compile(r"time", re.I),
]

HEADER_COLOR = "305496"
GREEN = "C6E0B4"

# Per-path parsed-frame cache + dropped-row log (path, metric) -> (before, after).
_DF_CACHE = {}
_DROPPED = {}


def make_console_unicode_safe() -> None:
    """Windows cp1252 consoles crash on non-Latin metric names; use UTF-8-out."""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, OSError):
            pass


def slugify(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9]+", "", str(name))


def detect_view(path: str) -> str:
    stem = os.path.splitext(os.path.basename(path))[0]
    for pat, label in VIEW_PATTERNS:
        if pat.search(stem):
            return label
    return stem


def is_lower_better(metric: str) -> bool:
    return any(p.search(metric) for p in LOWER_BETTER_PATTERNS)


def blocked_names() -> set:
    return set(NON_METRIC) - set(CFG.include_metric)


def safe_remove(path: str) -> None:
    """Remove a file without crashing on Windows file locks / missing files."""
    try:
        os.remove(path)
    except FileNotFoundError:
        pass
    except OSError as e:
        print(f"  WARN could not remove {os.path.basename(path)} "
              f"(is it open in Excel?): {e}")


def _read_input(path: str) -> pd.DataFrame:
    """Read + cache the configured sheet of a csv/xlsx, stripping column names."""
    key = os.path.abspath(path)
    if key not in _DF_CACHE:
        if str(path).lower().endswith(".csv"):
            df = pd.read_csv(path)
        else:
            df = pd.read_excel(path, sheet_name=CFG.sheet)
        df.columns = [str(c).strip() for c in df.columns]
        _DF_CACHE[key] = df
    return _DF_CACHE[key]


def _metric_ready_df(path: str) -> pd.DataFrame:
    """Return the frame with numeric-looking text columns coerced to numeric.

    Handles cells like "85.2%" or "1,234" that pandas would otherwise read as
    object dtype (and thus silently skip as metrics). The coerced frame is
    cached so the analysis phase sees the same values as detection.
    """
    df = _read_input(path).copy()
    for c in df.columns:
        if c in blocked_names() or pd.api.types.is_numeric_dtype(df[c]):
            continue
        s = pd.to_numeric(df[c], errors="coerce")
        if s.notna().any():
            gained = int(s.notna().sum()) - int(df[c].notna().sum())
            if max(gained, 0) > 0:
                frac = gained / max(len(s), 1)
                if frac > 0.25:
                    print(f"  WARN {c}: {frac:.0%} non-numeric cells "
                          f"coerced; 25%+ converted")
            df[c] = s
    _DF_CACHE[os.path.abspath(path)] = df
    return df


def default_inputs() -> list:
    files = sorted(glob.glob(os.path.join(CFG.dataset_dir, "*.xlsx")) +
                   glob.glob(os.path.join(CFG.dataset_dir, "*.csv")))
    return files


def prompt_inputs() -> list:
    try:
        raw = input("Enter file path(s) (space/comma-separated, drag-drop ok), "
                    "[Enter] for default Datasets/: ").strip()
    except EOFError:
        raw = ""
    if not raw:
        return default_inputs()
    tokens = re.split(r"[,\s]+", raw.strip())
    paths = []
    for t in tokens:
        t = t.strip('"\' ')
        if not t:
            continue
        if not os.path.isabs(t):
            t = os.path.join(SCRIPT_DIR, t)
        paths.append(os.path.abspath(t))
    missing = [p for p in paths if not os.path.exists(p)]
    if missing:
        raise FileNotFoundError("Not found: " + ", ".join(missing))
    return paths


def load_wide(path: str, metric: str, df: pd.DataFrame | None = None) -> pd.DataFrame:
    """Long (Number+Run Number) -> pivot; anything else treated as wide.

    Always drops rows with missing values so every analysis (and the branches)
    uses the same clean dataset.
    """
    if df is None:
        df = _read_input(path)
    if "Number" in df.columns and "Run Number" in df.columns:
        assert metric in df.columns, f"{metric} not in {list(df.columns)}"
        wide = df.pivot(index="Run Number", columns="Number", values=metric)
        wide.columns = [f"Model_{c}" for c in wide.columns]
        wide.index.name = "Run"
    elif "Run Number" in df.columns:
        wide = df.set_index("Run Number")
    else:
        wide = df.copy()
        wide.index.name = "Subject"
    before, after = len(wide), len(wide.dropna())
    if after < before:
        _DROPPED[(os.path.abspath(path), metric)] = (before, after)
    return wide.dropna()


def detect_metrics(path: str) -> list:
    df = _metric_ready_df(path)
    candidate_cols = [
        c for c in df.columns
        if c not in blocked_names() and pd.api.types.is_numeric_dtype(df[c])
    ]
    metrics = []
    for c in candidate_cols:
        try:
            wide = load_wide(path, c, df=df)
        except Exception:
            continue
        if wide.shape[1] < 2:
            continue
        # Skip only metrics with no variance at all anywhere (Friedman trivial).
        if (wide.nunique(axis=0) == 1).all():
            continue
        metrics.append(c)
    return metrics


def _apply_correction(pvals: np.ndarray, m: int):
    """Return (alpha_adj, adjusted_p). alpha_adj is the threshold used for sig."""
    pvals = np.asarray(pvals, dtype=float)
    if CFG.correction == "bonferroni":
        p_adj = np.minimum(pvals * m, 1.0)
        return CFG.alpha / m, p_adj
    if CFG.correction == "bh":
        n = len(pvals)
        order = np.argsort(pvals)
        ranks = np.empty_like(order)
        ranks[order] = np.arange(1, n + 1)
        p_sorted = pvals[order]
        adjusted_sorted = p_sorted * n / ranks[order]
        adjusted_sorted = np.minimum.accumulate(adjusted_sorted[::-1])[::-1]
        p_adj = np.empty_like(p_sorted)
        p_adj[order] = adjusted_sorted
        return CFG.alpha, np.minimum(p_adj, 1.0)
    return CFG.alpha, pvals  # none


def run_one(view: str, path: str, metric: str) -> dict:
    wide = load_wide(path, metric)
    n, k = wide.shape
    s = slugify(metric)

    desc = pd.DataFrame({
        "N": wide.count(),
        "Mean": wide.mean(),
        "SD": wide.std(),
        "Min": wide.min(),
        "Max": wide.max(),
    })
    mean_ranks = wide.rank(axis=1).mean().rename("MeanRank")
    desc = desc.join(mean_ranks).sort_values("MeanRank")

    Q, p_overall = friedmanchisquare(*[wide[c].values.astype(float) for c in wide.columns])
    Q = float(Q) if np.isfinite(Q) else float("nan")
    p_overall = float(p_overall) if np.isfinite(p_overall) else float("nan")
    W = float(Q) / (n * (k - 1)) if n > 0 and k > 1 else float("nan")
    if np.isnan(Q) or np.isnan(p_overall):
        decision = "INCONCLUSIVE"
    else:
        decision = "REJECT H0" if p_overall < CFG.alpha else "RETAIN H0"

    # Ranks (SPSS format)
    ranks = pd.DataFrame({"N": n, "Mean Rank": wide.rank(axis=1).mean().round(4)})
    ranks.index.name = "Condition"

    # Test statistics
    ts = pd.DataFrame([{"N": n, "Chi-Square": round(float(Q), 3)
                        if np.isfinite(Q) else float("nan"),
                        "df": k - 1, "Asymp. Sig.": p_overall}])

    # Post-hoc
    pairs = list(combinations(wide.columns, 2))
    m = len(pairs)
    rows = []
    for a, b in pairs:
        x, y = wide[a].astype(float), wide[b].astype(float)
        d = x - y  # match scipy wilcoxon(x, y): x is 'a' -> pos = a>b
        neg, pos, ties = int((d < 0).sum()), int((d > 0).sum()), int((d == 0).sum())
        try:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                _, p = wilcoxon(x, y, zero_method="wilcox")
        except ValueError:
            p = 1.0
        if pd.isna(p):
            p = 1.0
        rows.append({"Pair": f"{a} vs {b}", "Neg": neg, "Pos": pos,
                     "Ties": ties, "p": float(p)})

    alpha_adj, p_adj = _apply_correction(
        np.array([r["p"] for r in rows], dtype=float), m)
    for r, pa in zip(rows, p_adj):
        r["p_adj"] = float(pa)
        r["sig"] = bool(pa < CFG.alpha)
    post = pd.DataFrame(rows)

    dropped = _DROPPED.get((os.path.abspath(path), metric))
    if dropped and dropped[0] != dropped[1]:
        print(f"  NOTE {view}/{metric}: dropped {dropped[0] - dropped[1]} of "
              f"{dropped[0]} rows for missing values")

    return {"view": view, "path": path, "metric": metric, "slug": s,
            "wide": wide, "desc": desc, "ranks": ranks, "ts": ts,
            "post": post, "Q": Q, "df": k - 1, "p": p_overall,
            "W": W, "n": n, "k": k, "m": m, "alpha_adj": alpha_adj,
            "correction": CFG.correction, "decision": decision,
            "n_sig": int(post["sig"].sum())}


def clear_view_outputs(view_slug: str, results: list) -> None:
    slugs = {r["slug"] for r in results if r["view"] == view_slug}
    patterns = []
    for s in slugs:
        patterns += [f"Ranks_{view_slug}_{s}.csv", f"TestStatistics_{view_slug}_{s}.csv",
                     f"SPSS_Friedman_{view_slug}_{s}.txt", f"posthoc_{view_slug}_{s}.csv",
                     f"friedman_plots_{view_slug}_{s}.png",
                     f"Descriptives_{view_slug}_{s}.csv"]
    for pat in patterns:
        for f in glob.glob(os.path.join(CFG.out_dir, pat)):
            safe_remove(f)
    for f in glob.glob(os.path.join(CFG.final_dir, f"Ranks_{view_slug}_All_Metrics.xlsx")):
        safe_remove(f)
    for f in glob.glob(os.path.join(CFG.final_dir, f"TestStatistics_{view_slug}_All_Metrics.xlsx")):
        safe_remove(f)


def wipe_output_dirs() -> None:
    for d in (CFG.out_dir, CFG.final_dir):
        if not os.path.isdir(d):
            continue
        for f in glob.glob(os.path.join(d, "*")):
            safe_remove(f)


def save_per_metric(result: dict) -> None:
    view, s = result["view"], result["slug"]
    path = os.path.join(CFG.out_dir, f"Ranks_{view}_{s}.csv")
    result["ranks"].to_csv(path)
    path = os.path.join(CFG.out_dir, f"TestStatistics_{view}_{s}.csv")
    result["ts"].to_csv(path, index=False)
    path = os.path.join(CFG.out_dir, f"posthoc_{view}_{s}.csv")
    result["post"].to_csv(path, index=False)
    path = os.path.join(CFG.out_dir, f"Descriptives_{view}_{s}.csv")
    result["desc"].to_csv(path, index_label="Condition")

    txt_path = os.path.join(CFG.out_dir, f"SPSS_Friedman_{view}_{s}.txt")
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("FREQUENCIES / FRIEDMAN - SPSS replica (Python)\n")
        f.write(f"File: {os.path.basename(result['path'])} | Metric: {result['metric']} "
                f"| N={result['n']} k={result['k']}\n\n")
        f.write("Descriptive Statistics\n")
        f.write(result["desc"].to_string() + "\n\n")
        f.write("Ranks\n")
        f.write(result["ranks"].to_string() + "\n\n")
        f.write("Test Statistics\nFriedman Test\n")
        f.write(result["ts"].to_string(index=False) + "\n\n")
        f.write(f"a. Friedman Test | alpha={CFG.alpha:g} | p vs alpha: {result['decision']}\n")
        f.write(f"Kendall W = {result['W']:.4f}\n")
        f.write(f"Post-hoc {result['correction']} correction: m={result['m']}, "
                f"alpha_adj={result['alpha_adj']:.6f}, "
                f"significant={result['n_sig']}/{result['m']}\n")
        f.write("Pairwise Wilcoxon (two-sided, zero_method='wilcox'): "
                "Pos = # x>y (x is left condition).\n")

    # Plots
    fig, ax = plt.subplots(1, 2, figsize=(14, 4))
    sns.boxplot(data=result["wide"][result["desc"].index], ax=ax[0])
    ax[0].tick_params(axis="x", rotation=45)
    ax[0].set_title(f"{view} {result['metric']}: scores by model")
    result["desc"]["MeanRank"].sort_values().plot(kind="barh", ax=ax[1])
    ax[1].set_title("Mean ranks")
    ax[1].set_xlabel("Mean rank")
    plt.tight_layout()
    plt.savefig(os.path.join(CFG.out_dir, f"friedman_plots_{view}_{s}.png"), dpi=150)
    plt.close(fig)


def unique_sheet_name(slug: str, taken: set) -> str:
    base = (slug or "metric")[:31]
    name, i = base, 2
    while name in taken or name == "All_Metrics":
        suffix = f"_{i}"
        name = base[: 31 - len(suffix)] + suffix
        i += 1
    taken.add(name)
    return name


def dedupe_slugs_by_view(results: list) -> None:
    """Make slugs unique per view (files/excel sheets). Reuses metric names so a
    single file's metrics keep their natural filenames; only real collisions are
    disambiguated (numeric suffix, or source-stem tag when the same `view
    metric` pair comes from two different files)."""
    seen_slugs = {}
    seen_keys = {}
    for r in results:
        base = r["slug"] or "metric"
        key = (r["view"], r["metric"])
        tag = None
        if key in seen_keys and seen_keys[key] != r["path"]:
            tag = slugify(os.path.splitext(os.path.basename(r["path"]))[0])[:10] or "alt"
        seen_keys[key] = r["path"]
        slug = f"{base}_{tag}" if tag else base
        bucket = seen_slugs.setdefault(r["view"], set())
        candidate, i = slug, 2
        while candidate in bucket:
            candidate = f"{slug}_{i}"
            i += 1
        bucket.add(candidate)
        r["slug"] = candidate


def build_workbooks(view_slug: str, results: list) -> None:
    ordered = [r for r in results if r["view"] == view_slug]
    if not ordered:
        return
    font = lambda color, bold=False, size=11: Font(color="FFFFFF", bold=bold, size=size)
    align = lambda: Alignment(horizontal="center", vertical="center")
    header_fill = lambda: PatternFill("solid", fgColor=HEADER_COLOR)

    # Ranks workbook
    out_path = os.path.join(CFG.final_dir, f"Ranks_{view_slug}_All_Metrics.xlsx")
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        taken = set()
        for r in ordered:
            df = r["ranks"].reset_index()  # Condition + N + Mean Rank
            df.columns = ["Condition", "N", "Mean Rank"]
            sheet = unique_sheet_name(r["slug"], taken)
            df.to_excel(writer, sheet_name=sheet, index=False)
            ws = writer.sheets[sheet]
            ws.insert_rows(1)
            for c in range(df.shape[1]):
                cell = ws.cell(row=2, column=c + 1)
                cell.font = Font(color="FFFFFF", bold=False)
                cell.fill = header_fill()
                cell.alignment = align()
            top = ws.cell(row=1, column=1,
                          value=f"{r['metric']} — Friedman Mean Ranks ({view_slug})")
            top.font = Font(color="FFFFFF", bold=True, size=12)
            top.fill = header_fill()
            ws.column_dimensions["A"].width = 12
            ws.column_dimensions["B"].width = 8
            ws.column_dimensions["C"].width = 12
            best_rank = (
                df["Mean Rank"].min() if is_lower_better(r["metric"])
                else df["Mean Rank"].max()
            )
            for i in range(len(df)):
                if df.iloc[i]["Mean Rank"] == best_rank:
                    ws.cell(row=i + 3, column=3).fill = PatternFill("solid", fgColor=GREEN)

        rank_map = OrderedDict()
        for r in ordered:
            rank_map[r["metric"]] = r["ranks"]["Mean Rank"]
        combined = pd.DataFrame(rank_map)
        combined = combined.reindex(
            sorted(combined.index,
                   key=lambda x: int(re.search(r"\d+", str(x)).group(0))
                   if re.search(r"\d+", str(x)) else 0)
        )
        combined.index.name = "Condition"
        n_models = len(combined)
        for col in combined.columns:
            if is_lower_better(col):
                combined[col] = (n_models + 1) - combined[col]
        combined.insert(0, "Model Rank (avg)",
                        combined.mean(axis=1).rank(method="min", ascending=False).astype(int))
        combined = combined.sort_values("Model Rank (avg)")
        combined_reset = combined.reset_index()
        combined_reset.to_excel(writer, sheet_name="All_Metrics", index=False)
        ws = writer.sheets["All_Metrics"]
        ws.insert_rows(1)
        top = ws.cell(row=1, column=1,
                      value=f"Average Model Rank across {len(ordered)} {view_slug} metrics (lower = better)")
        top.font = Font(color="FFFFFF", bold=True, size=12)
        top.fill = header_fill()
        for c in range(combined_reset.shape[1]):
            cell = ws.cell(row=2, column=c + 1)
            cell.font = Font(color="FFFFFF", bold=True)
            cell.fill = header_fill()
            cell.alignment = align()
        ws.column_dimensions["A"].width = 12
    print(f"Saved -> {out_path} ({len(ordered)} metric sheets + All_Metrics)")

    # Test statistics workbook
    out_path = os.path.join(CFG.final_dir, f"TestStatistics_{view_slug}_All_Metrics.xlsx")
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        taken = set()
        for r in ordered:
            df = r["ts"].copy()
            sheet = unique_sheet_name(r["slug"], taken)
            df.to_excel(writer, sheet_name=sheet, index=False)
            ws = writer.sheets[sheet]
            for c in range(df.shape[1]):
                cell = ws.cell(row=1, column=c + 1)
                cell.font = Font(color="FFFFFF", bold=True)
                cell.fill = header_fill()
                cell.alignment = align()
            ws.insert_rows(1)
            top = ws.cell(row=1, column=1,
                          value=f"{r['metric']} — Friedman Test (SPSS), {view_slug}")
            top.font = Font(color="FFFFFF", bold=True, size=12)
            top.fill = header_fill()
            ws.column_dimensions["A"].width = 22
            for letter in ("B", "C", "D", "E"):
                ws.column_dimensions[letter].width = 14

        rows = [r["ts"].iloc[0].copy() for r in ordered]
        combined = pd.DataFrame(rows)
        combined.insert(0, "Metric", [r["metric"] for r in ordered])
        combined.to_excel(writer, sheet_name="All_Metrics", index=False)
        ws = writer.sheets["All_Metrics"]
        for c in range(combined.shape[1]):
            cell = ws.cell(row=1, column=c + 1)
            cell.font = Font(color="FFFFFF", bold=True)
            cell.fill = header_fill()
            cell.alignment = align()
        ws.insert_rows(1)
        top = ws.cell(row=1, column=1,
                      value=f"Friedman Test Statistics — all {len(ordered)} {view_slug} metrics")
        top.font = Font(color="FFFFFF", bold=True, size=12)
        top.fill = header_fill()
        ws.column_dimensions["A"].width = 22
        for letter in ("B", "C", "D", "E", "F"):
            ws.column_dimensions[letter].width = 14
    print(f"Saved -> {out_path} ({len(ordered)} metric sheets + All_Metrics)")


def write_overall_md(results: list) -> None:
    lines = [
        "# Friedman Analysis — Overall Results",
        "",
        "Fully automated by `friedmann_pipeline.py` (SPSS K-Related + Two Related Samples replica).",
        "",
        f"Post-hoc: Wilcoxon signed-rank per pair (`zero_method='wilcox'`), "
        f"correction = `{CFG.correction}`, alpha = {CFG.alpha:g}. "
        f"(`bonferroni` -> alpha/m, `bh` -> Benjamini-Hochberg FDR). "
        f"Kendall's `W = Q / (n * (k - 1))`.",
        "",
        "| View | Metric | N | k | Q (Chi-Square) | df | Asymp. Sig. (p) | Kendall W | Decision (vs alpha) | Sig pairs |",
        "|---|---|---|---|---|---|---|---|---|---|",
    ]
    for r in results:
        lines.append(
            f"| {r['view']} | {r['metric']} | {r['n']} | {r['k']} | {r['Q']:.3f} | "
            f"{r['df']} | {r['p']:.3e} | {r['W']:.3f} | {r['decision']} | "
            f"{r['n_sig']}/{r['m']} |"
        )
    md_path = os.path.join(CFG.out_dir, "Friedman_Results_Overall.md")
    with open(md_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    print(f"Saved -> {md_path} ({len(results)} runs)")


def configure(alpha: float = DEFAULT_ALPHA, out_dir: str = DEFAULT_OUT_DIR,
              final_dir: str = DEFAULT_FINAL_DIR, dataset_dir: str = DEFAULT_DATASET_DIR,
              sheet=0, correction: str = "bonferroni",
              include_metric=(), clean: bool = False) -> None:
    """Set runtime configuration from explicit params.

    Shared by the CLI entry point (`main()`) and the Streamlit UI so both drive
    exactly the same engine. Clears per-run caches.
    """
    global CFG
    CFG.alpha = float(alpha)
    CFG.out_dir = os.path.abspath(out_dir)
    CFG.final_dir = os.path.abspath(final_dir)
    CFG.dataset_dir = os.path.abspath(dataset_dir)
    try:
        CFG.sheet = int(sheet)
    except (TypeError, ValueError):
        CFG.sheet = sheet
    CFG.correction = correction
    CFG.include_metric = set(include_metric or [])
    CFG.clean = bool(clean)
    _DF_CACHE.clear()
    _DROPPED.clear()


def run_analysis(paths: list, selected: set | None = None):
    """Run the full pipeline for every (file, metric) in `paths`.

    `selected`: optional set of metric names; when given only those metrics are
    analyzed per file (metrics absent from a file are skipped and returned in
    `skipped`). Returns (results, summary_df, skipped).
    """
    os.makedirs(CFG.out_dir, exist_ok=True)
    os.makedirs(CFG.final_dir, exist_ok=True)
    if CFG.clean:
        wipe_output_dirs()

    results = []
    by_view = {}
    selected = set(selected) if selected else None
    skipped = []
    for path in paths:
        view = detect_view(path)
        metrics = detect_metrics(path)
        if not metrics:
            skipped.append((path, None, "no detectable metric columns"))
            print(f"SKIP {path}: no detectable metric columns")
            continue
        if selected is not None:
            for m in sorted(set(selected) - set(metrics)):
                skipped.append((path, m, "metric not present in file"))
                print(f"SKIP {os.path.basename(path)}: metric {m!r} not present")
            local = [m for m in metrics if m in selected]
        else:
            local = metrics
        for metric in local:
            r = run_one(view, path, metric)
            results.append(r)
            by_view.setdefault(view, []).append(r)
            print(f"[{os.path.basename(path)}] view={view} metric={metric}: "
                  f"n={r['n']} k={r['k']} Q={r['Q']:.3f} p={r['p']:.3e} "
                  f"{r['decision']} (sig pairs {r['n_sig']}/{r['m']})")
    if not results:
        raise SystemExit("No metrics processed from any input file.")

    dedupe_slugs_by_view(results)

    summary = pd.DataFrame([{
        "View": r["view"], "Metric": r["metric"], "N": r["n"], "k": r["k"],
        "Q": r["Q"], "df": r["df"], "p": r["p"], "W": r["W"],
        "pairs": r["m"], "correction": r["correction"], "alpha_adj": r["alpha_adj"],
        "sig": r["n_sig"],
    } for r in results])
    summ_path = os.path.join(CFG.out_dir, "Friedman_Summary_All.csv")
    summary.to_csv(summ_path, index=False)
    print(f"\nSaved summary -> {summ_path} ({len(results)} runs)")

    for view_slug, res in by_view.items():
        clear_view_outputs(view_slug, results)
        for r in res:
            save_per_metric(r)
        build_workbooks(view_slug, res)

    write_overall_md(results)

    print("\n=== Pipeline complete ===")
    print(summary[["View", "Metric", "Q", "p", "W", "sig"]].to_string(index=False))
    return results, summary, skipped


def main() -> None:
    make_console_unicode_safe()
    ap = argparse.ArgumentParser(
        description="Full Friedman analysis pipeline (SPSS K-Related + Two Related Samples replica).")
    ap.add_argument("--input", nargs="*", default=None,
                    help="Input file(s). Omit for interactive prompt / default Datasets/.")
    ap.add_argument("--alpha", type=float, default=DEFAULT_ALPHA,
                    help=f"Significance level (default {DEFAULT_ALPHA}).")
    ap.add_argument("--out-dir", default=DEFAULT_OUT_DIR,
                    help="Directory for per-metric CSVs, txt, plots, summary.")
    ap.add_argument("--final-dir", default=DEFAULT_FINAL_DIR,
                    help="Directory for final Excel workbooks.")
    ap.add_argument("--dataset-dir", default=DEFAULT_DATASET_DIR,
                    help="Default Datasets/ folder used when no --input given.")
    ap.add_argument("--sheet", default=0,
                    help="Excel sheet to read (index or name). Default 0 (first).")
    ap.add_argument("--correction", choices=["bonferroni", "bh", "none"],
                    default="bonferroni",
                    help="Pairwise post-hoc correction (default bonferroni).")
    ap.add_argument("--include-metric", nargs="*", default=None,
                    help="Force these columns to be treated as metrics even if "
                         "blacklisted (e.g. 'No. of Parameters').")
    ap.add_argument("--clean", action="store_true",
                    help="Wipe all files in output dirs before running.")
    args = ap.parse_args()

    configure(alpha=args.alpha, out_dir=args.out_dir, final_dir=args.final_dir,
              dataset_dir=args.dataset_dir, sheet=args.sheet,
              correction=args.correction,
              include_metric=args.include_metric, clean=args.clean)

    paths = args.input if args.input not in (None, []) else None
    if paths is None:
        paths = prompt_inputs()
        if not paths:
            raise SystemExit(f"No input files found under {CFG.dataset_dir}.")
    run_analysis(paths)


if __name__ == "__main__":
    main()