#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
import re
import os
import sys

# -----------------------------
# Normalisation helper
# -----------------------------
def normalize_word(word):
    """Remove diacritics, punctuation; normalise alef variants."""
    word = re.sub(r'[\u064B-\u0652\u0670]', '', word)   # diacritics
    word = re.sub(r'[إأٱآا]', 'ا', word)
    word = re.sub(r'[يى]', 'ي', word)
    word = re.sub(r'ة', 'ه', word)
    word = re.sub(r'[^\u0600-\u06FF]', '', word)         # keep Arabic only
    return word.strip()

def normalize_text(text):
    """Normalise a full sentence for marker matching."""
    text = re.sub(r'[\u064B-\u0652\u0670]', '', text)
    text = re.sub(r'[إأٱآا]', 'ا', text)
    text = re.sub(r'[يى]', 'ي', text)
    text = re.sub(r'ة', 'ه', text)
    return text

# -----------------------------
# Marker Dictionaries (normalised — no diacritics)
# -----------------------------
hypotactic_markers = {
    "Relative": (
        ["الذي", "التي", "الذين", "اللتان", "اللذان", "اللواتي", "اللائي",
         "والذي", "والتي", "والذين", "اللاتي", "اللذين"],
        "Elaboration (Identification)"
    ),
    "Causal": (
        ["لان", "حيث ان", "اذ", "اذ ان", "بما ان", "لكون", "بسبب",
         "نتيجه ل", "جراء", "من اجل ان"],
        "Enhancement (Cause, Purpose)"
    ),
    "Result": (
        ["حتي", "لكي", "كي", "لاجل ان", "بهدف ان", "علي ان", "الي ان"],
        "Enhancement (Purpose, Result)"
    ),
    "Temporal": (
        ["عندما", "حينما", "كلما", "منذ ان", "بينما", "في الوقت الذي",
         "ما ان", "بعد ان", "قبل ان", "حالما", "حين"],
        "Enhancement (Time)"
    ),
    "Conditional": (
        ["ان", "اذا", "لو", "لولا", "ما لم", "مهما"],
        "Enhancement (Condition)"
    ),
    "Concessive": (
        ["رغم ان", "بالرغم من ان", "مع ان", "حتي لو"],
        "Enhancement (Concession)"
    ),
    "Comparative": (
        ["كما", "كما ان", "كما لو ان", "مثلما", "علي نحو ما"],
        "Elaboration (Comparison)"
    ),
    "Projection: Locution": (
        ["قال", "يقول", "تقول", "اشار", "اشارت", "يضيف", "اضاف",
         "افاد", "ذكر", "صرح", "اوضح", "اعلن", "اكد"],
        "Projection (Locution)"
    ),
    "Projection: Idea": (
        ["يعتقد", "اعتقد", "يظن", "ظن", "اظن", "يري", "راي",
         "اري", "تصور", "تخيل"],
        "Projection (Idea)"
    )
}

paratactic_markers = {
    "Addition": (
        ["و", "كذلك", "كما", "ايضا", "علاوه علي ذلك", "فضلا عن ذلك",
         "بالاضافه الي", "وكذلك", "وايضا"],
        "Extension (Addition)"
    ),
    "Contrast": (
        ["لكن", "غير ان", "بيد ان", "الا ان", "مع ذلك", "ولكن"],
        "Extension (Contrast)"
    ),
    "Alternative": (
        ["او", "ام", "اما"],
        "Extension (Alternative)"
    ),
    "Result": (
        ["ف", "لذلك", "ولهذا", "ومن ثم", "اذن", "فلذلك"],
        "Enhancement (Result)"
    ),
    "Time": (
        ["ثم", "بعد ذلك", "قبل ذلك", "لاحقا"],
        "Enhancement (Time)"
    ),
    "Restatement": (
        ["اي", "بمعني", "اعني", "وهذا يعني ان"],
        "Elaboration (Restatement)"
    ),
    "Apposition": (
        ["بل", "انما", "لا بل", "خصوصا"],
        "Elaboration (Apposition)"
    )
}

# -----------------------------
# Build flat lookup dictionaries
# -----------------------------
marker_dict = {}        # normalised_marker → (taxis, category, lsr)
SUBORDINATE_MARKERS = set()

for cat, (markers, lsr) in hypotactic_markers.items():
    for m in markers:
        nm = normalize_word(m) if ' ' not in m else normalize_text(m)
        marker_dict[nm] = ("Hypotactic", cat, lsr)
        SUBORDINATE_MARKERS.add(nm)

for cat, (markers, lsr) in paratactic_markers.items():
    for m in markers:
        nm = normalize_word(m) if ' ' not in m else normalize_text(m)
        marker_dict[nm] = ("Paratactic", cat, lsr)

# Common words starting with و that are NOT coordinators
COMMON_WA_WORDS = {
    "وصف", "وجه", "وضعت", "وثيقه", "واجب", "واقعه",
    "واد", "وضع", "ورد", "ورده", "وقت", "وطن"
}

# -----------------------------
# Marker detection
# -----------------------------
def detect_markers(sentence):
    norm_sent = normalize_text(sentence)
    words     = norm_sent.split()
    found     = []
    seen      = set()

    # 1. Multi-word phrase matching (longest first)
    phrases = sorted([m for m in marker_dict if ' ' in m],
                     key=len, reverse=True)
    for phrase in phrases:
        if phrase in norm_sent and phrase not in seen:
            found.append((phrase, *marker_dict[phrase]))
            seen.add(phrase)

    # 2. Single-word matching
    for word in words:
        nw = normalize_word(word)
        if not nw:
            continue

        # exact match
        if nw in marker_dict and nw not in seen:
            found.append((word, *marker_dict[nw]))
            seen.add(nw)
            continue

        # و / ف prefix stripping
        for prefix in ["و", "ف"]:
            if nw.startswith(prefix) and len(nw) > 1:
                stem = nw[len(prefix):]
                if stem in marker_dict and nw not in COMMON_WA_WORDS and stem not in seen:
                    found.append((word, *marker_dict[stem]))
                    seen.add(stem)
                    break

        # standalone و
        if nw == "و" and "و" not in seen:
            found.append((word, *marker_dict.get("و", (word, "Paratactic", "Addition", "Extension (Addition)"))))
            seen.add("و")

    return found if found else [("", "Other", "Other", "Unclear")]


# -----------------------------
# Sentence classification
# -----------------------------
def classify_sentence(sentence):
    norm = normalize_text(sentence)
    words = norm.split()
    for word in words:
        nw = normalize_word(word)
        if nw in SUBORDINATE_MARKERS:
            return "Complex"
        for prefix in ["و", "ف"]:
            if nw.startswith(prefix) and len(nw) > 1:
                stem = nw[len(prefix):]
                if stem in SUBORDINATE_MARKERS:
                    return "Complex"
    # multi-word subordinators
    for phrase in [m for m in SUBORDINATE_MARKERS if ' ' in m]:
        if phrase in norm:
            return "Complex"
    return "Simple"


# -----------------------------
# Main Excel Processing
# -----------------------------
def process_excel(input_path, output_path):
    df = pd.read_excel(input_path)

    # Auto-detect Arabic sentence column
    ar_col = None
    for col in df.columns:
        if "arabic" in col.lower() or "sentence" in col.lower() or "ar" in col.lower():
            ar_col = col
            break
    if ar_col is None:
        ar_col = df.columns[0]
    print(f"ℹ️  Using column: '{ar_col}'")

    results = []
    for _, row in df.iterrows():
        text = str(row[ar_col]).strip()
        if not text or text == "nan":
            continue

        matches = detect_markers(text)
        first_marker, first_taxis, first_cat, first_lsr = matches[0]

        all_lsrs = list(dict.fromkeys(
            [m[3] for m in matches if m[3] != "Unclear"]
        ))
        all_cats = list(dict.fromkeys(
            [m[2] for m in matches if m[2] != "Other"]
        ))
        all_marker_words = [m[0] for m in matches if m[3] != "Unclear"]

        sentence_type = classify_sentence(text)

        results.append({
            "Sentence":                  text,
            "Sentence Type":             sentence_type,
            "Clause Marker":             first_marker,
            "Taxis":                     first_taxis if first_marker else "",
            "Logico-Semantic Relation":  ", ".join(all_lsrs) if all_lsrs else "Unclear",
            "Marker Category":           ", ".join(all_cats) if all_cats else "Other",
            "LSR Markers":               ", ".join(all_marker_words)
        })

    output_df = pd.DataFrame(results)
    output_df.to_excel(output_path, index=False)
    print(f"✅ Analysis complete. Output saved to: {output_path}")
