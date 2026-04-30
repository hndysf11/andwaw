#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import pandas as pd
import re
import os

def run_shift_analysis(english_file, arabic_file, output_file):
    # === Load Excel files ===
    df_en = pd.read_excel(english_file)
    df_ar = pd.read_excel(arabic_file)

    # Number Arabic sentences if not already numbered
    if 'Sentence #' not in df_ar.columns:
        df_ar.insert(0, 'Sentence #', range(1, len(df_ar) + 1))

    df_en.columns = df_en.columns.str.strip()
    df_ar.columns = df_ar.columns.str.strip()

    # Rename columns for merging
    df_en.rename(columns={"Sentence #": "Sentence_Num", "Sheet": "Sheet_Name"}, inplace=True)
    df_ar.rename(columns={"Sentence #": "Sentence_Num", "Sheet": "Sheet_Name"}, inplace=True)
    df_en['Sheet_Name'] = 'Sheet1'
    df_ar['Sheet_Name'] = 'Sheet1'

    # Merge English and Arabic
    combined_df = pd.merge(
        df_en, df_ar,
        on=["Sheet_Name", "Sentence_Num"],
        how="inner",
        suffixes=('_EN', '_AR')
    )

    combined_df.columns = combined_df.columns.str.replace("-", "_").str.replace(" ", "_").str.strip()
    combined_df = combined_df.drop_duplicates(subset=["Sheet_Name", "Sentence_Num"], keep='first')

    # Fix Taxis_AR values
    combined_df["Taxis_AR"] = combined_df["Taxis_AR"].replace({
        "Paratactic (implicit)": "Paratactic",
        "Hypotactic (implicit)": "Hypotactic"
    })
    combined_df["Taxis_AR"] = combined_df["Taxis_AR"].replace(["", None], "—").fillna("—")

    # --- Normalization helpers ---
    def normalize_taxis(value):
        if pd.isna(value):
            return ''
        value = re.sub(r"\s*\(.*?\)", "", str(value)).strip().lower()
        if value in ["paratactic", "parataxis"]:
            return "paratactic"
        elif value in ["hypotactic", "hypotaxis"]:
            return "hypotactic"
        elif value == "simple":
            return "simple"
        return value

    def normalize_lsr(lsr_str):
        if pd.isna(lsr_str):
            return set()
        lsr_str = str(lsr_str).lower().strip()
        raw_parts = re.split(r"[+;]", lsr_str)
        normalized = set()
        for part in raw_parts:
            part = part.strip()
            if "projection" in part:
                if "locution" in part:
                    normalized.add("projection (locution)")
                elif "idea" in part:
                    normalized.add("projection (idea)")
                else:
                    normalized.add("projection")
            elif "extension" in part:
                normalized.add("extension")
            elif "elaboration" in part:
                normalized.add("elaboration")
            elif "enhancement" in part:
                normalized.add("enhancement")
        return normalized

    def is_unclear(lsr):
        if pd.isna(lsr):
            return True
        lsr_str = str(lsr).strip().lower().replace("\u200f", "").replace("\u200e", "")
        lsr_str = re.sub(r"\s+", "", lsr_str)
        return lsr_str in {"", "-", "—", "unclear", "n/a", "none"}

    # --- Taxis shift detection ---
    def check_taxis_shift(row):
        taxis_en = normalize_taxis(row['Taxis_EN'])
        taxis_ar = normalize_taxis(row['Taxis_AR'])
        if taxis_en in ["", "nan"] or taxis_ar in ["", "nan"]:
            return "Unclear"
        if taxis_en == "simple" and taxis_ar == "simple":
            return "Unclear"
        if taxis_en == "—" and taxis_ar in ["paratactic", "hypotactic"]:
            return "Shift inc"
        if taxis_ar == "—" and taxis_en in ["paratactic", "hypotactic"]:
            return "Shift dec"
        if taxis_en == taxis_ar:
            return "No Shift"
        return "Shift"

    combined_df["Taxis_Shift"] = combined_df.apply(check_taxis_shift, axis=1)

    # --- Logico shift detection ---
    def check_logico_shift(row):
        lsr_en = normalize_lsr(row["Logico_Semantic_Relation_EN"])
        lsr_ar = normalize_lsr(row["Logico_Semantic_Relation_AR"])
        if is_unclear(row["Logico_Semantic_Relation_EN"]) and is_unclear(row["Logico_Semantic_Relation_AR"]):
            return "Unclear"
        if is_unclear(row["Logico_Semantic_Relation_EN"]) and not is_unclear(row["Logico_Semantic_Relation_AR"]):
            return "Shift inc"
        if is_unclear(row["Logico_Semantic_Relation_AR"]) and not is_unclear(row["Logico_Semantic_Relation_EN"]):
            return "Shift dec"
        if lsr_en == lsr_ar:
            return "No Shift"
        if lsr_ar.issubset(lsr_en) and len(lsr_en) > len(lsr_ar):
            return "No Shift dec"
        if lsr_en.issubset(lsr_ar) and len(lsr_ar) > len(lsr_en):
            return "No Shift inc"
        if len(lsr_en) == len(lsr_ar):
            return "Shift"
        return "Shift"

    combined_df["Logico_Shift"] = combined_df.apply(check_logico_shift, axis=1)

    # --- Logico Shift Label ---
    def check_logico_shift_label(row):
        if row.get("Logico_Shift") != "Shift":
            return ""
        lsr_en = normalize_lsr(row["Logico_Semantic_Relation_EN"])
        lsr_ar = normalize_lsr(row["Logico_Semantic_Relation_AR"])
        overlap = lsr_en & lsr_ar
        shift_parts = (lsr_en | lsr_ar) - overlap
        if overlap and shift_parts:
            return "No Shift + Shift"
        return ""

    combined_df["Logico_Shift_Label"] = combined_df.apply(check_logico_shift_label, axis=1)

    # --- Shift Pattern ---
    def show_shift_pattern(row):
        taxis_en = str(row['Taxis_EN']).strip().lower()
        taxis_ar = str(row['Taxis_AR']).strip().lower()
        lsr_en = str(row['Logico_Semantic_Relation_EN']).strip().lower()
        lsr_ar = str(row['Logico_Semantic_Relation_AR']).strip().lower()
        # Corrected taxis_pattern logic
        if taxis_en in ["", "nan"] or taxis_ar in ["", "nan"]:
            taxis_pattern = f"Unclear → {taxis_ar if taxis_ar not in ['', 'nan'] else 'Unclear'}"
        elif taxis_en == "simple" and taxis_ar == "simple":
            taxis_pattern = "Unclear"
        elif taxis_en == taxis_ar or \
             (taxis_en == "parataxis" and taxis_ar == "paratactic") or \
             (taxis_en == "hypotaxis" and taxis_ar == "hypotactic"):
            taxis_pattern = "No Shift"
        else:
            taxis_pattern = f"{taxis_en} → {taxis_ar}"
            
        # Existing logic for logico shift
        if lsr_en in ["", "nan", "unclear"] or lsr_ar in ["", "nan", "unclear"]:
            lsr_pattern = f"Unclear → {lsr_ar if lsr_ar not in ['','nan','unclear'] else 'Unclear'}"
        elif lsr_en == lsr_ar:
            lsr_pattern = "No Shift"
        else:
            lsr_pattern = f"{lsr_en} → {lsr_ar}"
        return f"Taxis: {taxis_pattern} | Logico: {lsr_pattern}"

    combined_df['Shift_Pattern'] = combined_df.apply(show_shift_pattern, axis=1)

    # --- Drop unneeded clause-level columns ---
    combined_df.drop(columns=["Clause #", "Clause Text", "Clause Type", "Sentence Type"], inplace=True, errors="ignore")

    # === Step 8: Shift summary (with and without 'Unclear') ===
    summary_data_with_unclear = {
        "Shift Type": [], "Shift": [], "No Shift": [], "Unclear": [], "Deleted": [], "Total": [],
        "Shift (%)": [], "No Shift (%)": [], "Unclear (%)": [], "Deleted (%)": []
    }

    summary_data_no_unclear = {
        "Shift Type": [], "Shift": [], "No Shift": [], "Total": [], "Shift (%)": [], "No Shift (%)": []
    }

    detailed_shift_data = {
        "Shift Type": [], "Shift": [], "Shift Inc": [], "Shift dec": [], "No Shift": [],
        "No Shift Inc": [], "No Shift dec": [], "Total": [], "Shift Inc (%)": [], "ShiftDec (%)": [],
        "Shift (%)": [], "No Shift (%)": [], "No Shift Inc(%)": [], "No Shift Dec (%)": []
    }

    for shift_col in ["Taxis_Shift", "Logico_Shift"]:
        df_main_only = combined_df.get("Clause_Type", pd.Series(["Main"]*len(combined_df))) == "Main"
        counts = combined_df[shift_col].value_counts(dropna=False)
        total_all = counts.sum()

        shift = counts.get("Shift", 0)
        shift_inc = counts.get("Shift inc", 0)
        shift_dec = counts.get("Shift dec", 0)
        no_shift = counts.get("No Shift", 0)
        no_shift_inc = counts.get("No Shift inc", 0)
        no_shift_dec = counts.get("No Shift dec", 0)
        both_shift = counts.get("No Shift + Shift", 0)
        unclear = counts.get("Unclear", 0)
        deleted = counts.get("DELETED", 0)

        shift_total = shift + shift_inc + shift_dec + both_shift
        no_shift_total = no_shift + no_shift_inc + no_shift_dec + both_shift

        # --- summary_data_with_unclear ---
        summary_data_with_unclear["Shift Type"].append(shift_col.replace("_", " "))
        summary_data_with_unclear["Shift"].append(shift_total)
        summary_data_with_unclear["No Shift"].append(no_shift_total)
        summary_data_with_unclear["Unclear"].append(unclear)
        summary_data_with_unclear["Deleted"].append(deleted)
        summary_data_with_unclear["Total"].append(total_all)

        summary_data_with_unclear["Shift (%)"].append(round(100 * shift_total / total_all, 2))
        summary_data_with_unclear["No Shift (%)"].append(round(100 * no_shift_total / total_all, 2))
        summary_data_with_unclear["Unclear (%)"].append(round(100 * unclear / total_all, 2))
        summary_data_with_unclear["Deleted (%)"].append(round(100 * deleted / total_all, 2))

        # --- summary_data_no_unclear ---
        filtered = combined_df[~combined_df[shift_col].isin(["Unclear", "DELETED"])]
        counts_filtered = filtered[shift_col].value_counts()
        total_filtered = counts_filtered.sum()

        shift = counts_filtered.get("Shift", 0)
        shift_inc = counts_filtered.get("Shift inc", 0)
        shift_dec = counts_filtered.get("Shift dec", 0)
        no_shift = counts_filtered.get("No Shift", 0)
        no_shift_inc = counts_filtered.get("No Shift inc", 0)
        no_shift_dec = counts_filtered.get("No Shift dec", 0)
        both_shift = counts_filtered.get("No Shift + Shift", 0)

        shift_total = shift + shift_inc + shift_dec + both_shift
        no_shift_total = no_shift + no_shift_inc + no_shift_dec + both_shift

        summary_data_no_unclear["Shift Type"].append(shift_col.replace("_", " "))
        summary_data_no_unclear["Shift"].append(shift_total)
        summary_data_no_unclear["No Shift"].append(no_shift_total)
        summary_data_no_unclear["Total"].append(total_filtered)
        summary_data_no_unclear["Shift (%)"].append(round(100 * shift_total / total_filtered, 2))
        summary_data_no_unclear["No Shift (%)"].append(round(100 * no_shift_total / total_filtered, 2))

        # --- detailed_shift_data ---
        detailed_shift_data["Shift Type"].append(shift_col.replace("_", " "))
        detailed_shift_data["Shift"].append(shift)
        detailed_shift_data["Shift Inc"].append(shift_inc)
        detailed_shift_data["Shift dec"].append(shift_dec)
        detailed_shift_data["No Shift"].append(no_shift)
        detailed_shift_data["No Shift Inc"].append(no_shift_inc)
        detailed_shift_data["No Shift dec"].append(no_shift_dec)
        detailed_shift_data["Total"].append(total_filtered)

        detailed_shift_data["Shift Inc (%)"].append(round(100 * shift_inc / total_filtered, 2))
        detailed_shift_data["ShiftDec (%)"].append(round(100 * shift_dec / total_filtered, 2))
        detailed_shift_data["Shift (%)"].append(round(100 * shift / total_filtered, 2))
        detailed_shift_data["No Shift (%)"].append(round(100 * no_shift  / total_filtered, 2))
        detailed_shift_data["No Shift Inc(%)"].append(round(100 * no_shift_inc / total_filtered, 2))
        detailed_shift_data["No Shift Dec (%)"].append(round(100 * no_shift_dec / total_filtered, 2))

    summary_df_with_unclear = pd.DataFrame(summary_data_with_unclear)
    summary_df_no_unclear = pd.DataFrame(summary_data_no_unclear)
    shift_incdec_df = pd.DataFrame(detailed_shift_data)

    # --- Step 9 & 10: Split patterns & summaries remain the same ---
    def split_shift_pattern(shift_pattern):
        try:
            taxis_part, logico_part = shift_pattern.split(" | ")
            taxis_value = taxis_part.replace("Taxis: ", "").strip()
            logico_value = logico_part.replace("Logico: ", "").strip()
            return pd.Series([taxis_value, logico_value])
        except Exception:
            return pd.Series([None, None])

    combined_df[['Taxis_Shift_Pattern', 'Logico_Shift_Pattern']] = combined_df['Shift_Pattern'].apply(split_shift_pattern)

    def fix_logico_if_simple(row):
        lsr_en = row["Logico_Semantic_Relation_EN"]
        lsr_ar = row["Logico_Semantic_Relation_AR"]
        taxis_en = row["Taxis_EN"]
        taxis_ar = row["Taxis_AR"]

        if str(taxis_en).strip().lower() == "simple" and str(lsr_en).strip().lower() == "unclear":
            row["Logico_Semantic_Relation_EN"] = "Simple"
        if str(taxis_ar).strip().lower() == "simple" and str(lsr_ar).strip().lower() == "unclear":
            row["Logico_Semantic_Relation_AR"] = "Simple"
        return row

    combined_df = combined_df.apply(fix_logico_if_simple, axis=1)

    def calc_counts_and_percentages(series):
        counts = series.value_counts(dropna=False)
        total = counts.sum()
        data = {"Category": [], "Count": [], "Percentage (%)": []}
        for cat, cnt in counts.items():
            data["Category"].append(cat)
            data["Count"].append(cnt)
            data["Percentage (%)"].append(round(100 * cnt / total, 2))
        return pd.DataFrame(data)

    taxis_shift_pattern_summary = calc_counts_and_percentages(combined_df['Taxis_Shift_Pattern'])
    logico_shift_pattern_summary = calc_counts_and_percentages(combined_df['Logico_Shift_Pattern'])

    # Filter by Logico_Shift values
    shift_inc_df = combined_df[combined_df['Logico_Shift'] == 'Shift inc']
    shift_dec_df = combined_df[combined_df['Logico_Shift'] == 'Shift dec']
    shift_shift_df = combined_df[combined_df['Logico_Shift'] == 'Shift']
    shift_noShift_df = combined_df[combined_df['Logico_Shift'] == 'No Shift']
    shift_unclear_df = combined_df[combined_df['Logico_Shift'] == 'Unclear']

    def summarize(df):
        summary = df['Logico_Shift_Pattern'].value_counts().reset_index()
        summary.columns = ['Logico_Shift_Pattern', 'Count']
        total = summary['Count'].sum()
        summary['Percentage'] = (summary['Count'] / total * 100).round(2)
        return summary.sort_values(by='Count', ascending=False)

    pattern_counts_inc = summarize(shift_inc_df)
    pattern_counts_dec = summarize(shift_dec_df)
    pattern_counts_shift = summarize(shift_shift_df)
    pattern_counts_noShift = summarize(shift_noShift_df)
    pattern_counts_unclear = summarize(shift_unclear_df)

    # --- Save all to Excel ---
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        combined_df.to_excel(writer, sheet_name="Combined Data", index=False)
        summary_df_with_unclear.to_excel(writer, sheet_name="Shift Summary (with Unclear)", index=False)
        summary_df_no_unclear.to_excel(writer, sheet_name="Shift Summary (No Unclear)", index=False)
        shift_incdec_df.to_excel(writer, sheet_name="Shift (inc-dec)", index=False)
        taxis_shift_pattern_summary.to_excel(writer, sheet_name="Taxis Shift Pattern Summary", index=False)
        logico_shift_pattern_summary.to_excel(writer, sheet_name="Logico Shift Pattern Summary", index=False)
        pattern_counts_inc.to_excel(writer, sheet_name='Logico_Shift_Inc_Summary', index=False)
        pattern_counts_dec.to_excel(writer, sheet_name='Logico_Shift_Dec_Summary', index=False)
        pattern_counts_shift.to_excel(writer, sheet_name='Logico_Shift_Shift_Summary', index=False)
        pattern_counts_noShift.to_excel(writer, sheet_name='Logico_Shift_NoShift_Summary', index=False)
        pattern_counts_unclear.to_excel(writer, sheet_name='Logico_Shift_Unclear_Summary', index=False)

    print(f"✅ Shift Analysis complete. File saved as: {output_file}")
    return combined_df

# -----------------------------
# MAIN: CLI support
# -----------------------------
# MAIN: CLI support
if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python shift_analysis.py English.xlsx Arabic.xlsx [output_file_or_folder]")
        sys.exit(1)

    english_file = sys.argv[1]
    arabic_file = sys.argv[2]

    # Validate input files
    if not os.path.exists(english_file):
        print(f"❌ File not found: {english_file}")
        sys.exit(1)

    if not os.path.exists(arabic_file):
        print(f"❌ File not found: {arabic_file}")
        sys.exit(1)

    # Optional output path
    if len(sys.argv) >= 4:
        output_arg = sys.argv[3]
        if output_arg.lower().endswith(".xlsx"):
            # User provided a full file path
            output_excel = output_arg
            os.makedirs(os.path.dirname(output_excel), exist_ok=True)
        else:
            # User provided a folder
            output_dir = output_arg
            os.makedirs(output_dir, exist_ok=True)
            en_base = os.path.splitext(os.path.basename(english_file))[0]
            ar_base = os.path.splitext(os.path.basename(arabic_file))[0]
            output_excel = os.path.join(output_dir, f"processed_shift_{en_base}_{ar_base}.xlsx")
    else:
        # Default folder
        output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output_files")
        os.makedirs(output_dir, exist_ok=True)
        en_base = os.path.splitext(os.path.basename(english_file))[0]
        ar_base = os.path.splitext(os.path.basename(arabic_file))[0]
        output_excel = os.path.join(output_dir, f"processed_shift_{en_base}_{ar_base}.xlsx")

    print(f"✅ Output will be saved to: {output_excel}")
    run_shift_analysis(english_file, arabic_file, output_excel)


