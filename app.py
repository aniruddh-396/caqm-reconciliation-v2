"""
CAQM Reconciliation - unified app for both modules.

    pip install streamlit openpyxl
    streamlit run app.py

Module 1 - Social Media Agency-wise Data of CAQM Tagged Complaints
           (04th Oct 2024 - 10th September 2026)
Module 2 - Social Media Category-wise Data of CAQM Tagged Complaints
           (04th Oct 2024 - 10th September 2026)
"""
import io
import os
import tempfile
import traceback
import zipfile
from datetime import date

import streamlit as st

import mapping as CFG
import reconcile as R1
import reconcile_category as R2
import reconcile_state as R3
import reconcile_ncr as R4
from errors import CaqmError
from datefmt import today_ist

st.set_page_config(page_title="CAQM Reconciliation", layout="wide")

M1 = "Social Media Agency-wise Data of CAQM Tagged Complaints (04th Oct 2024 - 10th September 2026)"
M2 = "Social Media Category-wise Data of CAQM Tagged Complaints (04th Oct 2024 - 10th September 2026)"
M3 = "Social Media State-wise Data (split parent into per-state files)"
M4 = "Social Media Agency-wise Data - Delhi-NCR (MCD consolidated, Excel + PDF)"

st.title("CAQM Complaint Data Reconciliation")

module = st.radio("Choose a module", [M1, M2, M3, M4], index=0)
st.divider()

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def save_uploads(**files):
    tmp = tempfile.mkdtemp()
    paths = {}
    for k, up in files.items():
        p = os.path.join(tmp, up.name)
        with open(p, "wb") as fh:
            fh.write(up.getbuffer())
        paths[k] = p
    return tmp, paths


def download_row(output, log, out_name):
    d1, d2 = st.columns(2)
    with open(output, "rb") as fh:
        d1.download_button("Download corrected workbook", fh.read(), out_name,
                           XLSX_MIME, type="primary", use_container_width=True)
    with open(log, "rb") as fh:
        d2.download_button("Download change log", fh.read(),
                           os.path.basename(log), XLSX_MIME, use_container_width=True)


def show_error(e):
    """Render a CaqmError as a readable headline + a table of issues so the
    user can see exactly which file, row and column is wrong."""
    st.error(e.headline)
    rows = []
    for i in e.issues:
        rows.append({
            "File": i.file,
            "Sheet": i.sheet or "-",
            "Row": i.row if i.row != "" else "-",
            "Column": i.column or "-",
            "Problem": i.problem,
            "Detail": i.detail,
            "Expected": i.expected or "-",
            "Found": i.found or "-",
            "How to fix": i.fix or "-",
        })
    st.dataframe(rows, use_container_width=True, hide_index=True)
    st.caption("Nothing was written. Fix the rows listed above in the source "
               "file and run again.")


def run_guarded(fn, *a, **kw):
    """Call an engine, turning every failure into a clear on-screen report."""
    try:
        return fn(*a, **kw), None
    except CaqmError as e:
        show_error(e)
        return None, "handled"
    except Exception:
        st.error("An unexpected problem stopped the run. Nothing was written.")
        st.code(traceback.format_exc())
        return None, "unexpected"


def show_unmatched_alert(unmatched, label="agency"):
    """Prominent alert naming every entry that could not be matched, split by
    direction so the user knows which file to fix."""
    if not unmatched:
        st.success(f"Every {label} in the format was found in the uploaded files.")
        return

    key = "agency" if label == "agency" else "category"
    not_found = [u for u in unmatched if u.get("severity") == "NOT FOUND"]
    not_in_fmt = [u for u in unmatched if u.get("severity") == "NOT IN FORMAT"]
    expected = [u for u in unmatched if u.get("severity", "").startswith("OK")]

    if not_found:
        names = "\n".join(
            f"- **{u[key]}**" + (f"  (row {u['row']}"
                                 + (f", {u['state']}" if u.get('state') and u['state'] != '-' else "")
                                 + ")" if u.get("row") != "-" else "")
            for u in not_found)
        st.error(
            f"**{len(not_found)} {label}(s) in the format were NOT found in the "
            f"uploaded file(s).** Their rows were left unchanged:\n\n{names}\n\n"
            f"Check the spelling in the uploaded base file, or add the missing "
            f"row there. Matching ignores case, spacing and punctuation.")

    if not_in_fmt:
        names = "\n".join(f"- **{u[key]}**" for u in not_in_fmt)
        st.warning(
            f"**{len(not_in_fmt)} {label}(s) in the uploaded file have no row in "
            f"the format**, so their numbers were not carried into the output:"
            f"\n\n{names}")

    if expected:
        st.info(f"{len(expected)} {label}(s) are known to be absent from the base "
                f"files and were skipped on purpose: "
                + ", ".join(u[key] for u in expected) + ".")


def changes_table(changes, has_source):
    rows = []
    for c in changes:
        label = getattr(c, "agency", None) or getattr(c, "category", None)
        row = {"Row": c.row, "State": c.state,
               ("Agency" if has_source else "Category"): label,
               "Col": c.column, "Field": c.field_name,
               "Old": "" if c.old is None else str(c.old),
               "New": "" if c.new is None else str(c.new),
               "Rule": c.rule}
        if has_source:
            row["Source"] = c.sources
        rows.append(row)
    return rows


# ===========================================================================
# MODULE 1 - AGENCY-WISE
# ===========================================================================
if module == M1:
    st.caption("Corrects **Agency_wise.xlsx** against **agency_analysis.xlsx** "
               "(counts) and **resolution_analysis.xlsx** (periods). MCD Delhi "
               "and NHAI are combined first - counts summed, periods averaged.")

    with st.sidebar:
        st.header("Module 1 settings")
        received_mode = st.radio(
            "Complaints Received is derived from",
            ["base_grand_total", "formula"],
            format_func=lambda m: ("Grand Total column in agency_analysis"
                                   if m == "base_grand_total"
                                   else "Pending + Resolved (sheet's own convention)"),
        )
        tolerance = st.number_input("Period tolerance (days)", 0.0, 10.0,
                                    float(CFG.PERIOD_TOLERANCE_DAYS), 0.25)
        update_totals = st.checkbox("Refresh period values on *_Total rows",
                                    value=CFG.UPDATE_TOTAL_ROWS)
        with st.expander("Handle mapping"):
            for h, a in CFG.HANDLE_TO_AGENCY.items():
                st.text(f"{h}  ->  {a}")

    st.markdown("**Editable file:** the fixed Agency-wise format is built in "
                "(`templates/agency_wise.xlsx`). You only supply the two base "
                "files and pick the end date.")

    c1, c2, c3 = st.columns([2, 2, 1.4])
    f_count = c1.file_uploader("agency_analysis.xlsx (counts)", type=["xlsx"], key="m1c")
    f_period = c2.file_uploader("resolution_analysis.xlsx (periods)", type=["xlsx"], key="m1p")
    m1_date = c3.date_input("End date", value=today_ist(), key="m1date",
                            help="Rewrites the end date in the sheet heading "
                                 "and the output file name.")

    with st.expander("Use a different editable file instead of the fixed format"):
        f_edit = st.file_uploader("Custom Agency-wise file (optional)",
                                  type=["xlsx"], key="m1e")
        st.caption("Leave empty to use the fixed format. Any file used here must "
                   "still have the standard columns, or you will get a format "
                   "report naming the exact row and column.")

    st.caption(f"Heading will read: *... (04th Oct 2024 - "
               f"{R1.heading_date(m1_date)})*  |  file: "
               f"`Agency_wise_CORRECTED ({R1.file_date(m1_date)}).xlsx`")

    if st.button("Compare and correct", type="primary",
                 disabled=not (f_count and f_period), key="m1run"):
        CFG.RECEIVED_MODE = received_mode
        CFG.PERIOD_TOLERANCE_DAYS = float(tolerance)
        CFG.UPDATE_TOTAL_ROWS = bool(update_totals)
        ups = {"count": f_count, "period": f_period}
        if f_edit:
            ups["edit"] = f_edit
        _, paths = save_uploads(**ups)
        with st.spinner("Reconciling..."):
            res, err = run_guarded(R1.reconcile, paths.get("edit"),
                                   paths["count"], paths["period"],
                                   tempfile.mkdtemp(), end_date=m1_date)
        st.session_state["m1"] = res
        st.session_state["m1_date"] = m1_date

    res = st.session_state.get("m1")
    if res:
        m = st.columns(5)
        m[0].metric("Matched", len(res["matched"]))
        m[1].metric("Unmatched", len(res["unmatched"]))
        m[2].metric("Cells changed", len(res["changes"]))
        m[3].metric("Total-row edits", len(res["total_updates"]))
        m[4].metric("Base warnings", len(res["integrity"]))

        show_unmatched_alert(res["unmatched"], "agency")

        if res["integrity"]:
            st.warning(f"{len(res['integrity'])} rows in agency_analysis have "
                       "Pending + Resolved != Grand Total - see Base Integrity in the log.")
        d = st.session_state.get("m1_date", today_ist())
        download_row(res["output"], res["log"],
                     f"Agency_wise_CORRECTED ({R1.file_date(d)}).xlsx")
        st.divider()
        t = st.tabs(["Changes", "Total rows", "Aggregated", "Base integrity", "Unmatched"])
        t[0].dataframe(changes_table(res["changes"], True), use_container_width=True, hide_index=True)
        t[1].dataframe(changes_table(res["total_updates"], True), use_container_width=True, hide_index=True)
        t[2].dataframe(
            [{"Agency": r.agency, "Handles": " + ".join(r.handles),
              "Pending": r.pending, "Resolved": r.resolved, "Grand Total": r.grand_total,
              "Avg Resolution": r.avg_resolution, "Avg Pendency": r.avg_pendency}
             for r in res["records"].values() if r.is_aggregate],
            use_container_width=True, hide_index=True)
        t[3].dataframe(res["integrity"], use_container_width=True, hide_index=True)
        t[4].dataframe(res["unmatched"], use_container_width=True, hide_index=True)

# ===========================================================================
# MODULE 2 - CATEGORY-WISE
# ===========================================================================
elif module == M2:
    st.caption("Corrects **caqm_category_wise.xlsx** against "
               "**sectoral_analysis.xlsx**. Rows are matched on the pair "
               "(State, Category). Complaints Received <- Grand Total, "
               "Pending <- Pending, Resolved <- Resolved.")

    st.markdown("**Editable file:** the fixed Category-wise format is built in "
                "(`templates/category_wise.xlsx`). You only supply the base file "
                "and pick the end date.")

    c1, c2 = st.columns([3, 1.4])
    f_base = c1.file_uploader("sectoral_analysis.xlsx (base)", type=["xlsx"], key="m2b")
    m2_date = c2.date_input("End date", value=today_ist(), key="m2date",
                            help="Rewrites the end date in the sheet heading "
                                 "and the output file name.")

    with st.expander("Use a different editable file instead of the fixed format"):
        f_edit = st.file_uploader("Custom Category-wise file (optional)",
                                  type=["xlsx"], key="m2e")
        st.caption("Leave empty to use the fixed format. Any file used here must "
                   "still have the standard columns, or you will get a format "
                   "report naming the exact row and column.")

    st.caption(f"Heading will read: *... (04th Oct 2024 - "
               f"{R2.heading_date(m2_date)})*  |  file: "
               f"`caqm_category_wise_CORRECTED ({R2.file_date(m2_date)}).xlsx`")

    if st.button("Compare and correct", type="primary",
                 disabled=not f_base, key="m2run"):
        ups = {"base": f_base}
        if f_edit:
            ups["edit"] = f_edit
        _, paths = save_uploads(**ups)
        with st.spinner("Reconciling..."):
            res, err = run_guarded(R2.reconcile_category, paths.get("edit"),
                                   paths["base"], tempfile.mkdtemp(),
                                   end_date=m2_date)
        st.session_state["m2"] = res
        st.session_state["m2_date"] = m2_date

    res = st.session_state.get("m2")
    if res:
        m = st.columns(5)
        m[0].metric("Matched", len(res["matched"]))
        m[1].metric("Unmatched", len(res["unmatched"]))
        m[2].metric("Cells changed", len(res["changes"]))
        m[3].metric("Total-row edits", len(res["total_updates"]))
        m[4].metric("Base warnings", len(res["integrity"]))
        show_unmatched_alert(res["unmatched"], "category")

        if res["integrity"]:
            st.warning(f"{len(res['integrity'])} rows in sectoral_analysis have "
                       "Pending + Resolved != Grand Total - see Base Integrity in the log.")
        d2 = st.session_state.get("m2_date", today_ist())
        download_row(res["output"], res["log"],
                     f"caqm_category_wise_CORRECTED ({R2.file_date(d2)}).xlsx")
        st.divider()
        t = st.tabs(["Changes", "Total rows", "Base integrity", "Unmatched"])
        t[0].dataframe(changes_table(res["changes"], False), use_container_width=True, hide_index=True)
        t[1].dataframe(changes_table(res["total_updates"], False), use_container_width=True, hide_index=True)
        t[2].dataframe(res["integrity"], use_container_width=True, hide_index=True)
        t[3].dataframe(res["unmatched"], use_container_width=True, hide_index=True)

# ===========================================================================
# MODULE 3 - STATE-WISE (split parent -> per-state files)
# ===========================================================================
elif module == M3:
    st.caption("Splits the parent workbook (**2025.xlsx**) into one file per "
               "state - Delhi, UP, Haryana, Rajasthan - using the fixed "
               "per-state layouts. Delhi's three MCD zones are summed into one "
               "'Municipal Corporation of Delhi' row. NHAI and Northern Railway "
               "are not used in any state.")

    c1, c2 = st.columns([2, 1])
    f_parent = c1.file_uploader("Parent workbook (2025.xlsx)", type=["xlsx"], key="m3p")
    end_dt = c2.date_input("End date (the '- to' date)", value=today_ist(),
                           key="m3date",
                           help="Sets the end date shown in every heading and "
                                "file name. Start date is fixed at 01 Jan 2025.")

    st.caption(f"Files will be named e.g. "
               f"`Delhi_Social Media Agency-wise Complaint Redressal "
               f"(01-01-2025 - {end_dt.strftime('%d-%m-%Y')}).xlsx`  and headed "
               f"`... (01st Jan 2025 - {R3.heading_date(end_dt)}) in Delhi`.")

    if st.button("Generate state files", type="primary",
                 disabled=not f_parent, key="m3run"):
        tmp = tempfile.mkdtemp()
        pp = os.path.join(tmp, f_parent.name)
        with open(pp, "wb") as fh:
            fh.write(f_parent.getbuffer())
        outdir = os.path.join(tmp, "out")
        with st.spinner("Generating..."):
            res, err = run_guarded(R3.generate_all, pp, end_dt, outdir)
        st.session_state["m3"] = res

    res = st.session_state.get("m3")
    if res:
        st.success(f"Generated {len(res['results'])} state files.")

        # one zip with everything
        zbuf = io.BytesIO()
        with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as z:
            for r in res["results"]:
                z.write(r["output"], arcname=r["filename"])
        zbuf.seek(0)
        st.download_button("Download all 4 as ZIP", zbuf.read(),
                           f"CAQM_State-wise_{end_dt.strftime('%d-%m-%Y')}.zip",
                           "application/zip", type="primary", use_container_width=True)

        st.divider()
        cols = st.columns(len(res["results"]))
        for col, r in zip(cols, res["results"]):
            with open(r["output"], "rb") as fh:
                col.download_button(f"{r['state']}", fh.read(), r["filename"],
                                    XLSX_MIME, use_container_width=True)

        st.divider()
        for r in res["results"]:
            with st.expander(f"{r['state']} - {r['filename']}"):
                st.dataframe(
                    [{"Row": c["row"], "Agency": c["agency"], "Kind": c["kind"],
                      "Received": c["received"], "Resolved": c["resolved"],
                      "Not Resolved": c["notres"]} for c in r["changes"]],
                    use_container_width=True, hide_index=True)

# ===========================================================================
# MODULE 4 - DELHI-NCR AGENCY-WISE, MCD CONSOLIDATED (Excel + PDF)
# ===========================================================================
else:
    st.caption("Upload a flat Delhi-NCR agency-wise workbook (e.g. the "
               "'2025-2026' or '2021-2025' report). East + North + South "
               "Delhi Municipal Corporation are summed into one "
               "'Municipal Corporation of Delhi' row placed at Sr. No. 12. "
               "Produces one corrected Excel and one matching PDF. The date "
               "range is taken from the file's own heading unless you "
               "override the end date below.")

    f_src = st.file_uploader("Source workbook", type=["xlsx"], key="m4src")
    override = st.checkbox("Override the end date shown in the heading",
                           value=False, key="m4override")
    m4_date = None
    if override:
        m4_date = st.date_input("End date", value=today_ist(), key="m4date")

    if st.button("Generate Excel + PDF", type="primary",
                 disabled=not f_src, key="m4run"):
        tmp = tempfile.mkdtemp()
        sp = os.path.join(tmp, f_src.name)
        with open(sp, "wb") as fh:
            fh.write(f_src.getbuffer())
        with st.spinner("Generating..."):
            res, err = run_guarded(R4.generate, sp, os.path.join(tmp, "out"),
                                   end_date=m4_date)
        st.session_state["m4"] = res

    res = st.session_state.get("m4")
    if res:
        m = st.columns(3)
        m[0].metric("Agencies in output", res["row_count"])
        m[1].metric("MCD combined position", res["mcd_position"])
        m[2].metric("MCD total received", res["mcd_combined"].received)

        st.success(f"Heading: *{res['heading']}*")

        d1, d2 = st.columns(2)
        with open(res["xlsx"], "rb") as fh:
            d1.download_button("Download Excel", fh.read(), res["xlsx_name"],
                               XLSX_MIME, type="primary", use_container_width=True)
        with open(res["pdf"], "rb") as fh:
            d2.download_button("Download PDF", fh.read(), res["pdf_name"],
                               "application/pdf", type="primary",
                               use_container_width=True)

        st.divider()
        mcd = res["mcd_combined"]
        st.markdown(f"**Municipal Corporation of Delhi** (East + North + South "
                   f"combined) - Received: {mcd.received}, Resolved: "
                   f"{mcd.resolved}, Not Resolved: {mcd.notres}")
        st.caption(f"Grand Total after merge: Received {res['total']['received']}, "
                  f"Resolved {res['total']['resolved']}, Not Resolved "
                  f"{res['total']['notres']} - identical to the source total, "
                  f"as merging rows can never change a sum.")
