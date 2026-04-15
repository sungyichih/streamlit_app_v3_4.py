import io
from typing import List, Tuple
import pandas as pd
import streamlit as st

st.set_page_config(page_title="BOM Mapping Tool", layout="wide")


# ----------------------------
# Helpers
# ----------------------------
def normalize_text(value):
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.lower() == "nan":
        return ""
    return text


def normalize_key(value):
    return normalize_text(value).upper()


def read_excel_safely(uploaded_file, sheet_name):
    filename = uploaded_file.name.lower()

    if filename.endswith(".xlsx") or filename.endswith(".xlsm"):
        engine = "openpyxl"
    elif filename.endswith(".xls"):
        engine = "xlrd"
    else:
        raise ValueError(
            f"Unsupported file type: {uploaded_file.name}. Please upload .xlsx, .xlsm, or .xls"
        )

    uploaded_file.seek(0)

    try:
        return pd.read_excel(uploaded_file, sheet_name=sheet_name, engine=engine)
    except Exception as e:
        raise ValueError(
            f'Cannot find sheet "{sheet_name}" in file "{uploaded_file.name}". '
            "Please rename or provide the correct mapping sheet."
        ) from e


# ----------------------------
# Original BOM parsing
# A=CPN, B=Description, C=Qty, D=Location, E=MFG, F=MPN
# G onward = alternative fields in MFG/MPN pairs
# ----------------------------
def extract_bom_mpn_pairs(row_values: List[str]) -> List[Tuple[str, str, str]]:
    pairs = []

    primary_mfg = normalize_text(row_values[4]) if len(row_values) > 4 else ""
    primary_mpn = normalize_text(row_values[5]) if len(row_values) > 5 else ""
    if primary_mfg or primary_mpn:
        pairs.append((primary_mfg, primary_mpn, "Primary"))

    alt_values = row_values[6:] if len(row_values) > 6 else []
    alt_index = 1
    for i in range(0, len(alt_values), 2):
        alt_mfg = normalize_text(alt_values[i]) if i < len(alt_values) else ""
        alt_mpn = normalize_text(alt_values[i + 1]) if i + 1 < len(alt_values) else ""
        if alt_mfg or alt_mpn:
            pairs.append((alt_mfg, alt_mpn, f"Alt {alt_index}"))
            alt_index += 1

    return pairs


def read_original_bom(uploaded_file, sheet_name="BOM", data_start_row=2):
    uploaded_file.seek(0)
    filename = uploaded_file.name.lower()

    if filename.endswith(".xlsx") or filename.endswith(".xlsm"):
        engine = "openpyxl"
    elif filename.endswith(".xls"):
        engine = "xlrd"
    else:
        raise ValueError(
            f"Unsupported file type: {uploaded_file.name}. Please upload .xlsx, .xlsm, or .xls"
        )

    skiprows = max(data_start_row - 1, 0)
    try:
        df = pd.read_excel(
            uploaded_file,
            sheet_name=sheet_name,
            header=None,
            skiprows=skiprows,
            engine=engine,
        )
    except Exception as e:
        raise ValueError(
            f'Cannot find sheet "{sheet_name}" in file "{uploaded_file.name}".'
        ) from e

    if df.shape[1] < 6:
        raise ValueError("Original BOM must include at least columns A through F.")

    base_rows = []
    mpn_rows = []

    for _, row in df.iterrows():
        row_values = row.tolist()

        cpn = normalize_text(row_values[0]) if len(row_values) > 0 else ""
        description = normalize_text(row_values[1]) if len(row_values) > 1 else ""
        qty = normalize_text(row_values[2]) if len(row_values) > 2 else ""
        location = normalize_text(row_values[3]) if len(row_values) > 3 else ""

        if not any([cpn, description, qty, location]):
            continue

        base_rows.append({
            "Customer_CPN": cpn,
            "Description": description,
            "Qty_Per_Board": qty,
            "Location": location,
        })

        for mfg, mpn, source in extract_bom_mpn_pairs(row_values):
            if mfg or mpn:
                mpn_rows.append({
                    "Customer_CPN": cpn,
                    "Description": description,
                    "Qty_Per_Board": qty,
                    "Location": location,
                    "BOM_MFG": mfg,
                    "BOM_MPN": mpn,
                    "Source": source,
                })

    base_df = pd.DataFrame(base_rows).drop_duplicates().reset_index(drop=True)
    mpn_df = pd.DataFrame(mpn_rows).drop_duplicates().reset_index(drop=True)

    if base_df.empty:
        raise ValueError("No usable Original BOM data found.")

    return base_df, mpn_df


# ----------------------------
# Organized mapping readers
# Fixed positions:
# SPN-CPN Mapping: A=SPN, B=CPN
# SPN-MPN Mapping: A=SPN, B=MFG, C=MPN
# ----------------------------
def read_organized_cpn_mapping(uploaded_file, sheet_name="SPN-CPN Mapping"):
    df = read_excel_safely(uploaded_file, sheet_name=sheet_name)

    # If first row is a header, skip it
    first_row = [normalize_key(v) for v in df.iloc[0].tolist()] if len(df) > 0 else []
    if len(first_row) >= 2 and first_row[0] == "SPN" and first_row[1] == "CPN":
        df = df.iloc[1:].reset_index(drop=True)

    if df.shape[1] < 2:
        raise ValueError(
            'Sheet "SPN-CPN Mapping" must have at least 2 columns (A=SPN, B=CPN).'
        )

    out = df.iloc[:, [0, 1]].copy()
    out.columns = ["SPN", "CPN"]

    out["SPN"] = out["SPN"].apply(normalize_text)
    out["CPN"] = out["CPN"].apply(normalize_text)

    out = out[(out["SPN"] != "") & (out["CPN"] != "")].drop_duplicates().reset_index(drop=True)
    return out


def read_organized_mpn_mapping(uploaded_file, sheet_name="SPN-MPN Mapping"):
    df = read_excel_safely(uploaded_file, sheet_name=sheet_name)

    # If first row is a header, skip it
    first_row = [normalize_key(v) for v in df.iloc[0].tolist()] if len(df) > 0 else []
    if len(first_row) >= 3 and first_row[0] == "SPN" and first_row[2] == "MPN":
        df = df.iloc[1:].reset_index(drop=True)

    if df.shape[1] < 3:
        raise ValueError(
            'Sheet "SPN-MPN Mapping" must have at least 3 columns (A=SPN, B=MFG, C=MPN).'
        )

    out = df.iloc[:, [0, 1, 2]].copy()
    out.columns = ["SPN", "System_MFG", "System_MPN"]

    out["SPN"] = out["SPN"].apply(normalize_text)
    out["System_MFG"] = out["System_MFG"].apply(normalize_text)
    out["System_MPN"] = out["System_MPN"].apply(normalize_text)

    out = out[(out["SPN"] != "") & (out["System_MPN"] != "")].drop_duplicates().reset_index(drop=True)
    return out


# ----------------------------
# Mapping / comparison
# ----------------------------
def map_cpn_to_spn(original_base_df, original_mpn_df, cpn_mapping_df, system_mpn_df):
    # Build Customer CPN -> BOM MPN set
    bom_groups = {}
    for _, row in original_mpn_df.iterrows():
        cpn = normalize_text(row["Customer_CPN"])
        mpn = normalize_key(row["BOM_MPN"])
        if cpn not in bom_groups:
            bom_groups[cpn] = set()
        if mpn:
            bom_groups[cpn].add(mpn)

    # Build SPN -> System MPN set
    sys_groups = {}
    for _, row in system_mpn_df.iterrows():
        spn = normalize_text(row["SPN"])
        mpn = normalize_key(row["System_MPN"])
        if spn not in sys_groups:
            sys_groups[spn] = set()
        if mpn:
            sys_groups[spn].add(mpn)

    # Build CPN -> candidate SPNs
    mapping = cpn_mapping_df.copy()
    mapping["CPN_KEY"] = mapping["CPN"].apply(normalize_key)

    cpn_to_spns = (
        mapping.groupby("CPN_KEY")["SPN"]
        .agg(lambda x: sorted(set(normalize_text(v) for v in x if normalize_text(v))))
        .to_dict()
    )

    result_rows = []

    for _, row in original_base_df.iterrows():
        customer_cpn = normalize_text(row["Customer_CPN"])
        cpn_key = normalize_key(customer_cpn)
        description = row.get("Description", "")
        qty = row.get("Qty_Per_Board", "")
        location = row.get("Location", "")

        candidate_spns = cpn_to_spns.get(cpn_key, [])
        bom_set = bom_groups.get(customer_cpn, set())

        if not candidate_spns:
            result_rows.append({
                "Customer_CPN": customer_cpn,
                "Description": description,
                "Qty_Per_Board": qty,
                "Location": location,
                "SPN": "",
                "Candidate_SPNs": "",
                "Selection_Status": "Missing SPN",
                "Best_Overlap_Count": 0,
            })
            continue

        if len(candidate_spns) == 1:
            selected_spn = candidate_spns[0]
            overlap_count = len(bom_set & sys_groups.get(selected_spn, set()))
            result_rows.append({
                "Customer_CPN": customer_cpn,
                "Description": description,
                "Qty_Per_Board": qty,
                "Location": location,
                "SPN": selected_spn,
                "Candidate_SPNs": selected_spn,
                "Selection_Status": "Unique match",
                "Best_Overlap_Count": overlap_count,
            })
            continue

        scored = []
        for spn in candidate_spns:
            system_set = sys_groups.get(spn, set())
            overlap = len(bom_set & system_set)
            scored.append((spn, overlap))

        scored = sorted(scored, key=lambda x: x[1], reverse=True)
        best_spn, best_score = scored[0]

        top_ties = [spn for spn, score in scored if score == best_score]

        if len(top_ties) == 1:
            selection_status = "Auto-selected by MPN overlap"
            selected_spn = best_spn
        else:
            selection_status = "Ambiguous - same overlap"
            selected_spn = ""

        result_rows.append({
            "Customer_CPN": customer_cpn,
            "Description": description,
            "Qty_Per_Board": qty,
            "Location": location,
            "SPN": selected_spn,
            "Candidate_SPNs": " / ".join(candidate_spns),
            "Selection_Status": selection_status,
            "Best_Overlap_Count": best_score,
        })

    return pd.DataFrame(result_rows)


def build_mpn_compare(mapped_df, original_mpn_df, system_mpn_df):
    bom_groups = {}
    for _, row in original_mpn_df.iterrows():
        cpn = normalize_text(row["Customer_CPN"])
        mpn = normalize_key(row["BOM_MPN"])
        if cpn not in bom_groups:
            bom_groups[cpn] = set()
        if mpn:
            bom_groups[cpn].add(mpn)

    sys_groups = {}
    for _, row in system_mpn_df.iterrows():
        spn = normalize_text(row["SPN"])
        mpn = normalize_key(row["System_MPN"])
        if spn not in sys_groups:
            sys_groups[spn] = set()
        if mpn:
            sys_groups[spn].add(mpn)

    compare_rows = []
    for _, row in mapped_df.iterrows():
        cpn = normalize_text(row["Customer_CPN"])
        spn = normalize_text(row.get("SPN", ""))
        desc = normalize_text(row.get("Description", ""))
        loc = normalize_text(row.get("Location", ""))
        selection_status = normalize_text(row.get("Selection_Status", ""))

        bom_set = bom_groups.get(cpn, set())

        if not spn:
            compare_rows.append({
                "Customer_CPN": cpn,
                "SPN": "",
                "Description": desc,
                "Location": loc,
                "Candidate_SPNs": normalize_text(row.get("Candidate_SPNs", "")),
                "Selection_Status": selection_status,
                "BOM_MPN_List": " / ".join(sorted(bom_set)),
                "System_MPN_List": "",
                "Missing_In_System": "",
                "Extra_In_System": "",
                "MPN_Compare_Status": selection_status if selection_status else "Missing SPN",
            })
            continue

        system_set = sys_groups.get(spn, set())
        missing_in_system = sorted(bom_set - system_set)
        extra_in_system = sorted(system_set - bom_set)

        if not bom_set and not system_set:
            status = "No MPN Data"
        elif not missing_in_system and not extra_in_system:
            status = "Full Match"
        elif missing_in_system and extra_in_system:
            status = "Partial Match"
        elif missing_in_system:
            status = "Missing in System"
        else:
            status = "Extra in System"

        compare_rows.append({
            "Customer_CPN": cpn,
            "SPN": spn,
            "Description": desc,
            "Location": loc,
            "Candidate_SPNs": normalize_text(row.get("Candidate_SPNs", "")),
            "Selection_Status": selection_status,
            "BOM_MPN_List": " / ".join(sorted(bom_set)),
            "System_MPN_List": " / ".join(sorted(system_set)),
            "Missing_In_System": " / ".join(missing_in_system),
            "Extra_In_System": " / ".join(extra_in_system),
            "MPN_Compare_Status": status,
        })

    return pd.DataFrame(compare_rows)


def build_missing_spn_list(mapped_df, original_mpn_df):
    mpn_grouped = (
        original_mpn_df.groupby("Customer_CPN")["BOM_MPN"]
        .agg(lambda x: " / ".join(sorted({normalize_key(v) for v in x if normalize_key(v)})))
        .reset_index()
    )

    missing = mapped_df[
        mapped_df["Selection_Status"].isin(["Missing SPN", "Ambiguous - same overlap"])
    ].copy()

    missing = missing.merge(mpn_grouped, on="Customer_CPN", how="left")
    missing = missing.rename(columns={"BOM_MPN": "BOM_MPN_List"})
    return missing.reset_index(drop=True)


def build_summary(original_base_df, mapped_df, compare_df):
    return pd.DataFrame([
        {"Metric": "Original BOM rows", "Value": len(original_base_df)},
        {"Metric": "Unique match count", "Value": int((mapped_df["Selection_Status"] == "Unique match").sum())},
        {"Metric": "Auto-selected by MPN overlap count", "Value": int((mapped_df["Selection_Status"] == "Auto-selected by MPN overlap").sum())},
        {"Metric": "Ambiguous count", "Value": int((mapped_df["Selection_Status"] == "Ambiguous - same overlap").sum())},
        {"Metric": "Missing SPN count", "Value": int((mapped_df["Selection_Status"] == "Missing SPN").sum())},
        {"Metric": "Full Match MPN count", "Value": int((compare_df["MPN_Compare_Status"] == "Full Match").sum())},
        {"Metric": "Partial Match count", "Value": int((compare_df["MPN_Compare_Status"] == "Partial Match").sum())},
        {"Metric": "Missing in System count", "Value": int((compare_df["MPN_Compare_Status"] == "Missing in System").sum())},
        {"Metric": "Extra in System count", "Value": int((compare_df["MPN_Compare_Status"] == "Extra in System").sum())},
    ])


def make_result_excel(original_base_df, original_mpn_df, mapped_df, compare_df, missing_spn_df, summary_df):
    output = io.BytesIO()

    # 先排序 MPN_Compare，讓問題集中在前面
    compare_sorted = compare_df.copy()

    status_priority = {
        "Missing SPN": 1,
        "Ambiguous - same overlap": 1,
        "Missing in System": 2,
        "Extra in System": 2,
        "Partial Match": 3,
        "No MPN Data": 4,
        "Full Match": 5,
    }

    compare_sorted["__sort_priority"] = compare_sorted["MPN_Compare_Status"].map(status_priority).fillna(99)
    compare_sorted = compare_sorted.sort_values(
        by=["__sort_priority", "MPN_Compare_Status", "Customer_CPN"],
        ascending=[True, True, True],
        kind="stable"
    ).drop(columns=["__sort_priority"])

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        original_base_df.to_excel(writer, sheet_name="Original_BOM_Normalized", index=False)
        original_mpn_df.to_excel(writer, sheet_name="Original_BOM_MPN_List", index=False)
        mapped_df.to_excel(writer, sheet_name="CPN_to_SPN_Map", index=False)
        compare_sorted.to_excel(writer, sheet_name="MPN_Compare", index=False)
        missing_spn_df.to_excel(writer, sheet_name="Missing_SPN", index=False)
        summary_df.to_excel(writer, sheet_name="Summary", index=False)

        from openpyxl.styles import PatternFill

        ws = writer.sheets["MPN_Compare"]

        yellow_fill = PatternFill(fill_type="solid", start_color="FFF2CC", end_color="FFF2CC")
        orange_fill = PatternFill(fill_type="solid", start_color="FCE5CD", end_color="FCE5CD")
        dark_red_fill = PatternFill(fill_type="solid", start_color="EA9999", end_color="EA9999")

        headers = [cell.value for cell in ws[1]]
        status_col_idx = headers.index("MPN_Compare_Status") + 1

        for row_idx in range(2, ws.max_row + 1):
            status = ws.cell(row=row_idx, column=status_col_idx).value

            if status == "Partial Match":
                fill = yellow_fill
            elif status in ["Missing in System", "Extra in System"]:
                fill = orange_fill
            elif status in ["Missing SPN", "Ambiguous - same overlap"]:
                fill = dark_red_fill
            else:
                fill = None

            if fill:
                for col_idx in range(1, ws.max_column + 1):
                    ws.cell(row=row_idx, column=col_idx).fill = fill

    output.seek(0)
    return output

# ----------------------------
# UI
# ----------------------------
st.title("Original BOM → CPN/SPN/MPN Mapping Tool")
st.caption(
    "Uses organized mapping files:\n"
    '- Original BOM sheet name = "BOM"\n'
    '- System CPN mapping sheet name = "SPN-CPN Mapping"\n'
    '- System MPN mapping sheet name = "SPN-MPN Mapping"'
)

st.info(
    'Please prepare files with these sheet names:\n'
    '• Original BOM: "BOM"\n'
    '• Organized CPN file: "SPN-CPN Mapping"\n'
    '• Organized MPN file: "SPN-MPN Mapping"'
)

col1, col2, col3 = st.columns(3)

with col1:
    original_bom_file = st.file_uploader(
        "Upload Original BOM",
        type=["xlsx", "xls", "xlsm"],
        key="bom",
    )
    original_bom_start_row = st.number_input(
        "Original BOM data starts at row",
        min_value=1,
        value=2,
        step=1,
    )

with col2:
    organized_cpn_file = st.file_uploader(
        "Upload Organized CPN-SPN file",
        type=["xlsx", "xls", "xlsm"],
        key="cpn",
    )

with col3:
    organized_mpn_file = st.file_uploader(
        "Upload Organized SPN-MPN file",
        type=["xlsx", "xls", "xlsm"],
        key="mpn",
    )

process = st.button("Process files", type="primary")

if process:
    if original_bom_file is None or organized_cpn_file is None or organized_mpn_file is None:
        st.error("Please upload all 3 files first.")
    else:
        try:
            original_base_df, original_mpn_df = read_original_bom(
                original_bom_file,
                sheet_name="BOM",
                data_start_row=int(original_bom_start_row),
            )

            cpn_mapping_df = read_organized_cpn_mapping(
                organized_cpn_file,
                sheet_name="SPN-CPN Mapping",
            )

            mpn_mapping_df = read_organized_mpn_mapping(
                organized_mpn_file,
                sheet_name="SPN-MPN Mapping",
            )

            mapped_df = map_cpn_to_spn(
                original_base_df,
                original_mpn_df,
                cpn_mapping_df,
                mpn_mapping_df,
            )

            compare_df = build_mpn_compare(
                mapped_df,
                original_mpn_df,
                mpn_mapping_df,
            )

            missing_spn_df = build_missing_spn_list(
                mapped_df,
                original_mpn_df,
            )

            summary_df = build_summary(
                original_base_df,
                mapped_df,
                compare_df,
            )

            st.success("Files processed successfully.")

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Original BOM rows", len(original_base_df))
            m2.metric("Missing SPN", int((mapped_df["Selection_Status"] == "Missing SPN").sum()))
            m3.metric("Ambiguous", int((mapped_df["Selection_Status"] == "Ambiguous - same overlap").sum()))
            diff_count = int(
                compare_df["MPN_Compare_Status"].isin(
                    ["Partial Match", "Missing in System", "Extra in System", "Missing SPN", "Ambiguous - same overlap"]
                ).sum()
            )
            m4.metric("MPN differences", diff_count)

            st.subheader("Summary")
            st.dataframe(summary_df, use_container_width=True, height=280)

            st.subheader("CPN to SPN Mapping")
            st.dataframe(mapped_df, use_container_width=True, height=320)

            st.subheader("MPN Compare")
            diff_only = st.checkbox("Show differences only", value=True)
            if diff_only:
                display_compare = compare_df[
                    compare_df["MPN_Compare_Status"] != "Full Match"
                ].copy()
            else:
                display_compare = compare_df.copy()

            st.dataframe(display_compare, use_container_width=True, height=360)

            tab1, tab2, tab3 = st.tabs([
                "Original BOM MPN List",
                "Missing SPN",
                "Organized SPN-MPN Mapping",
            ])

            with tab1:
                st.dataframe(original_mpn_df, use_container_width=True, height=320)

            with tab2:
                st.dataframe(missing_spn_df, use_container_width=True, height=320)

            with tab3:
                st.dataframe(mpn_mapping_df, use_container_width=True, height=320)

            excel_data = make_result_excel(
                original_base_df,
                original_mpn_df,
                mapped_df,
                compare_df,
                missing_spn_df,
                summary_df,
            )

            st.download_button(
                label="Download result Excel",
                data=excel_data,
                file_name="original_bom_mapping_result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error(f"Processing failed: {e}")
