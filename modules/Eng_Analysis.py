#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Lightweight EN Analysis for web version — no spaCy, uses regex only.
"""

import re
import pandas as pd

# ── Marker dictionaries ───────────────────────────────────────────────────────
LSR_MAP = {
    "Extension": [
        "and", "also", "moreover", "furthermore", "besides", "likewise",
        "but", "yet", "however", "nevertheless", "nonetheless", "whereas",
        "or", "either", "neither", "alternatively",
    ],
    "Elaboration": [
        "for example", "for instance", "in other words", "namely", "that is",
        "such as", "including", "which", "who", "whose", "that",
    ],
    "Enhancement": [
        "then", "after", "before", "when", "while", "since", "until", "once",
        "because", "so", "therefore", "thus", "hence", "consequently",
        "if", "unless", "although", "even though", "though", "so that",
        "in order that", "as a result", "due to",
    ],
    "Projection (Locution)": [
        "said", "says", "say", "told", "tell", "asked", "claimed", "claim",
        "explained", "explain", "stated", "state", "argued", "argue",
        "announced", "declared", "replied", "responded",
    ],
    "Projection (Idea)": [
        "thought", "think", "believed", "believe", "knew", "know", "felt",
        "feel", "supposed", "suppose", "assumed", "assume", "wondered",
        "realized", "understood",
    ],
}

HYPOTACTIC = [
    "although", "because", "when", "while", "if", "unless", "since",
    "until", "after", "before", "which", "who", "whose", "that",
    "even though", "though", "so that", "in order that",
]

PARATACTIC = ["and", "but", "or", "nor", "yet", "so"]


def detect_taxis(sentence):
    s = sentence.lower()
    for m in HYPOTACTIC:
        if re.search(r'\b' + re.escape(m) + r'\b', s):
            return "Hypotaxis"
    for m in PARATACTIC:
        if re.search(r'\b' + re.escape(m) + r'\b', s):
            return "Parataxis"
    return "—"


def detect_lsr(sentence):
    s = sentence.lower()
    matched_labels = []
    matched_markers = []

    for label, kws in LSR_MAP.items():
        for kw in kws:
            if re.search(r'\b' + re.escape(kw) + r'\b', s):
                matched_labels.append(label)
                matched_markers.append(kw)
                break

    if not matched_labels:
        return "—", "—"

    seen = dict.fromkeys(matched_labels)
    return " + ".join(seen.keys()), ", ".join(matched_markers)


def process_excel_file(input_path, output_path, include_embedded=False):
    xls = pd.read_excel(input_path, sheet_name=None)
    all_rows = []

    for sheet, df in xls.items():
        # Auto-detect sentence column
        col = next((c for c in df.columns
                    if "english" in c.lower() or "sentence" in c.lower()), None)
        if col is None:
            print(f"⚠️  Skipping sheet '{sheet}' — no sentence column found.")
            continue

        sents = df[col].dropna().astype(str).tolist()
        for i, sent in enumerate(sents, 1):
            if not sent.strip() or sent == "nan":
                continue
            taxis      = detect_taxis(sent)
            lsr, markers = detect_lsr(sent)
            stype      = "Complex" if taxis != "—" else "Simple"

            all_rows.append({
                "Sheet":                    sheet,
                "Sentence #":               i,
                "Sentence":                 sent,
                "Sentence Type":            stype,
                "Taxis":                    taxis,
                "Logico-Semantic Relation": lsr,
                "Taxis Marker":             markers,
                "Logico Marker":            markers,
            })

    if not all_rows:
        pd.DataFrame().to_excel(output_path, index=False)
        return

    final_df = pd.DataFrame(all_rows)

    # Summary
    taxis_df = (final_df["Taxis"].value_counts()
                .reset_index()
                .rename(columns={"index": "Taxis", "Taxis": "Count"}))
    lsr_df = (final_df["Logico-Semantic Relation"].value_counts()
              .reset_index()
              .rename(columns={"index": "LSR",
                                "Logico-Semantic Relation": "Count"}))

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        final_df.to_excel(writer, sheet_name="Main_Clauses", index=False)
        taxis_df.to_excel(writer, sheet_name="Taxis_Summary", index=False)
        lsr_df.to_excel(writer, sheet_name="LSR_Summary", index=False)

    print(f"✅ EN Analysis done → {output_path}")
