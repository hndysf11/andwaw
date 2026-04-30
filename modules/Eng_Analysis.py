#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import spacy
try:
    import spacy_transformers
except ImportError:
    pass
import pandas as pd
from collections import Counter


import os
import sys
import spacy
try:
    import spacy_transformers
except ImportError:
    pass
def get_spacy_model_path():
    """
    Returns the correct path to en_core_web_sm
    Works in development and inside a frozen PyInstaller app on macOS
    """
    if getattr(sys, "frozen", False):
        # In frozen app, _MEIPASS points to Resources
        base = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
        model_dir = os.path.join(base, "spacy_models", "en_core_web_sm", "en_core_web_sm-3.8.0")
        if os.path.exists(model_dir):
            return model_dir
        else:
            raise FileNotFoundError(f"Cannot find spaCy model at: {model_dir}")

    # In development mode
    return "en_core_web_sm"

# Load the model
# nlp loaded lazily
_nlp = None
def get_nlp():
    global _nlp
    if _nlp is None:
        _nlp = spacy.load(get_spacy_model_path())
    return _nlp
print(f"✅ spaCy model loaded from: {get_spacy_model_path()}")


# === Sentence Type Classification ===
def classify_sentence(sentence):
    doc = nlp(sentence)

    # Define clause-like dependency labels
    clause_like_deps = {"ccomp", "xcomp", "advcl", "relcl", "conj", "acl"}
    clause_count = 0

    for token in doc:
        # Count common clause-level structures
        if token.dep_ in clause_like_deps and token.head.pos_ in {"VERB", "AUX"}:
            clause_count += 1

        # Count cleft constructions (e.g., "It was the boy who cried.")
        if token.text.lower() == "who" and token.dep_ in {"nsubj", "nsubjpass"}:
            clause_count += 1

        # Detect relative pronouns that might begin a clause
        if token.lower_ in {"who", "whose", "which", "that", "where", "when"}:
            clause_count += 1

        # Reduced relative clause via participles (e.g., "The man standing there...")
        if token.pos_ == "VERB" and token.tag_ in {"VBG", "VBN"} and token.dep_ == "acl":
            clause_count += 1

    # Count root/main verbs
    clause_count += sum(1 for t in doc if t.dep_ == "ROOT" and t.pos_ in {"VERB", "AUX"})

    # Return classification
    return "Complex" if clause_count > 1 else "Simple"


# === Clause Splitting ===
def split_clauses(doc):
    clauses = []
    root_sent_text = next(doc.sents).text.strip()
    for token in doc:
        if token.dep_ in ("ccomp", "xcomp", "advcl", "relcl", "acl") and token.head.pos_ == "VERB":
            clause_text = " ".join([t.text for t in token.subtree])
            clauses.append((clause_text.strip(), "Embedded"))
        elif token.dep_ == "conj" and token.pos_ == "VERB":
            clause_text = " ".join([t.text for t in token.subtree])
            clauses.append((clause_text.strip(), "Coordinate"))
    if not any(cl[0] == root_sent_text for cl in clauses):
        clauses.append((root_sent_text, "Main"))
    unique = []
    seen = set()
    for clause_text, clause_type in clauses:
        if clause_text not in seen:
            unique.append((clause_text, clause_type))
            seen.add(clause_text)
    return unique

# === Marker Dictionaries ===
coordination_only = {
    "and", "or", "but", "nor", "yet", "so", "either", "neither", "not only", "also",
    "for", "plus", "as well as", "both", "not just", "rather", "nor yet", "no more",
    "either or", "neither nor", "else"
}

logic_map = {
    "Extension": {
        #  Addition
        "and", "also", "moreover", "furthermore", "in addition", "besides", "as well", "too", "what is more",
        "additionally", "further", "coupled with", "along with", "together with", "likewise", "similarly",
        # Contrast
        "but", "yet", "however", "nevertheless", "nonetheless", "still", "whereas", "while", "on the contrary",
        "albeit", "notwithstanding", "even so", "though", "in spite of that", "despite that",
        # Alternation
        "or", "either", "neither", "alternatively", "on the other hand", "otherwise", "else", "on the flip side",
        "conversely",
    },
    "Elaboration": {
        "for example", "for instance", "in other words", "namely", "that is", "i.e.", "which", "who", "whose", "that", "such as", "especially",
        "including", "e.g." "viz", "to illustrate", "particularly", "chiefly", "in particular", "among others",
    },
    "Enhancement": {
        # Time
        "then", "after", "before", "when", "while", "since", "until", "once", "meanwhile", "as soon as", "by the time", "as", "as long as",
        "earlier", "lately", "recently", "subsequently", "eventually", "afterward", "henceforth",
        # Cause/Result
        "because", "since", "so", "therefore", "thus", "as a result", "hence", "consequently",
        "due to", "owing to", "on account of", "thereupon", "accordingly",
        # Condition
        "if", "unless", "provided that", "in case", "as long as", "even if", "only if", "otherwise",
        "assuming that", "in the event that", "lest", "supposing",
        # Concession
        "although", "even though", "though", "while", "whereas", "despite", "in spite of",
        "granted that", "albeit", "notwithstanding", "even if",
        # Comparison
        "as", "as if", "as though", "like", "in the same way", "equally", "in like manner", "correspondingly",
        # Contrastive Comparison
        "rather than", "instead of", "in contrast", "whereas", "while", "as opposed to",
        # Purpose/Result
        "so that", "in order that", "so", "to", "so as to", "with the aim of", "with a view to", "so as to", "for the purpose of",
    },
    "Projection (Locution)": {
        "say", "says", "said", "tell", "told", "report", "ask", "claim", "explain", "add", "adds", "added",
        "comment", "commented", "mention", "mentioned", "state", "stated", "argue", "argued", "announce", "declare",
         "reply", "replied", "respond", "responded", "assert", "asserted", "acknowledge", "acknowledged", "confirm", "confirmed"},
    "Projection (Idea)": {
        "think", "thinks", "thought", "believe", "wonder", "know", "knew", "realize", "understand", "feel",
        "suspect", "assume", "guess", "hope", "wish", "doubt", "remember", "forget", "consider", "considered",
        "reckon", "reckoned", "imagine", "imaged", "suppose", "supposed"},
}

def determine_taxis(doc):
    for token in doc:
        if token.dep_ == "cc" and token.text.lower() in {"and", "or", "but"}:
            for child in token.head.children:
                if child.dep_ == "conj":
                    return "Parataxis"
    for token in doc:
        if token.dep_ in {"mark", "advcl", "relcl", "ccomp", "xcomp", "acl"}:
            return "Hypotaxis"
    return "—"

# === LSR + Marker Detection with split
def determine_logico_semantic(doc):
    text = doc.text.lower()
    lsr_matched = []
    taxis_markers = set()
    logic_markers = set()

    for label, keywords in logic_map.items():
        for kw in keywords:
            if " " in kw and kw in text:
                if label.startswith("Extension") and kw in coordination_only:
                    taxis_markers.add(kw)
                else:
                    logic_markers.add(kw)
                lsr_matched.append((label, kw))

    for token in doc:
        w = token.text.lower()
        lemma = token.lemma_.lower()
        for label, keywords in logic_map.items():
            if w in {k.lower() for k in keywords} or lemma in {k.lower() for k in keywords}:
                if label.startswith("Extension") and w in coordination_only:
                    taxis_markers.add(w)
                else:
                    logic_markers.add(w)
                lsr_matched.append((label, token.text))

    for i, token in enumerate(doc[:-1]):
        if token.lemma_.lower() in {"add", "say", "tell", "claim", "explain"}:
            next_token = doc[i + 1]
            if next_token.text in {":", "“", '"', "'"}:
                logic_markers.add(token.text)
                lsr_matched.append(("Projection (Locution)", token.text))

    seen = set()
    result = []
    for label, marker in lsr_matched:
        key = (label.lower(), marker.lower())
        if key not in seen:
            result.append((label, marker))
            seen.add(key)

    return result or [("—", "—")], taxis_markers, logic_markers

# === Clause-Level Analyzer ===
def analyze_sentences_to_rows(sentences, include_embedded=False):
    all_rows = []
    for i, sentence in enumerate(sentences, 1):
        if not isinstance(sentence, str) or not sentence.strip():
            continue
        doc = nlp(sentence)
        sentence_type = classify_sentence(sentence)

        if sentence_type == "Complex":
            clauses = split_clauses(doc)
            taxis = determine_taxis(doc)
            lsr_marker_pairs, taxis_markers, logic_markers = determine_logico_semantic(doc)
            unique_lsr_labels = []
            seen_lsr = set()
            for lsr, _ in lsr_marker_pairs:
                if lsr not in seen_lsr:
                    unique_lsr_labels.append(lsr)
                    seen_lsr.add(lsr)
            logic = " + ".join(unique_lsr_labels)
            taxis_marker = " + ".join(sorted(set(taxis_markers))) if taxis_markers else "—"
            logic_marker = " + ".join(sorted(set(logic_markers))) if logic_markers else "—"
        else:
            clauses = [(sentence, "Main")]
            taxis = "—"
            logic = "—"
            taxis_marker = "—"
            logic_marker = "—"

        for j, (clause, clause_type) in enumerate(clauses, 1):
            if clause_type == "Embedded" and not include_embedded:
                continue
            all_rows.append({
                "Sheet": "",
                "Sentence #": i,
                "Sentence": sentence,
                "Clause #": j,
                "Clause Text": clause,
                "Clause Type": clause_type,
                "Sentence Type": sentence_type,
                "Taxis": taxis,
                "Logico-Semantic Relation": logic,
                "Taxis Marker": taxis_marker,
                "Logico Marker": logic_marker
            })
    return all_rows

# === Normalize LSR Label ===
def normalize_lsr_label(label):
    if label in {"Projection (Locution)", "Projection (Idea)"}:
        return label
    base_label = label.split(" (")[0] if label != "—" else label
    if base_label in {"Extension", "Elaboration", "Enhancement"}:
        return base_label
    return label

# === Summary Stats ===
def compute_summary_stats(df):
    # Only count main clauses
    df_main = df[df["Clause Type"] == "Main"]

    normalized_lsr = []
    for lsr_cell in df_main["Logico-Semantic Relation"]:
        if lsr_cell != "—":
            split_lsrs = [normalize_lsr_label(label.strip()) for label in lsr_cell.split("+")]
            normalized_lsr.extend(split_lsrs)

    lsr_counter = Counter(normalized_lsr)
    taxis_counts = Counter(df_main["Taxis"])

    def count_individual_markers(series):
        counter = Counter()
        for cell in series:
            if cell != "—":
                markers = [m.strip() for m in cell.split("+")]
                counter.update(markers)
        return counter

    taxis_marker_counts = count_individual_markers(df_main["Taxis Marker"])
    logic_marker_counts = count_individual_markers(df_main["Logico Marker"])

    for d in [taxis_counts, taxis_marker_counts, logic_marker_counts]:
        d.pop("—", None)

    extension = lsr_counter.get("Extension", 0)
    elaboration = lsr_counter.get("Elaboration", 0)
    enhancement = lsr_counter.get("Enhancement", 0)
    projection_loc = lsr_counter.get("Projection (Locution)", 0)
    projection_idea = lsr_counter.get("Projection (Idea)", 0)

    expansion_count = extension + elaboration + enhancement
    projection_count = projection_loc + projection_idea
    total_lsr = sum(lsr_counter.values())

    logic_data = []
    for key in ["Extension", "Elaboration", "Enhancement", "Projection (Locution)", "Projection (Idea)"]:
        count = lsr_counter.get(key, 0)
        pct = round(count / total_lsr * 100, 2) if total_lsr > 0 else 0
        logic_data.append({
            "Logico-Semantic Relation": key,
            "Count": count,
            "Percentage": pct
        })

    logic_data.append({
        "Logico-Semantic Relation": "Expansion (Extension + Elaboration + Enhancement)",
        "Count": expansion_count,
        "Percentage": round(expansion_count / total_lsr * 100, 2) if total_lsr > 0 else 0
    })
    logic_data.append({
        "Logico-Semantic Relation": "Projection (Locution + Idea)",
        "Count": projection_count,
        "Percentage": round(projection_count / total_lsr * 100, 2) if total_lsr > 0 else 0
    })

    logic_df = pd.DataFrame(logic_data)

    taxis_df = make_df(taxis_counts, "Taxis Type")
    taxis_marker_df = make_df(taxis_marker_counts, "Taxis Marker")
    logic_marker_df = make_df(logic_marker_counts, "Logico Marker")

    return taxis_df, logic_df, taxis_marker_df, logic_marker_df

def make_df(counter, label):
    total = sum(counter.values())
    return pd.DataFrame([
        {label: k, "Count": v, "Percentage": round(v / total * 100, 2) if total > 0 else 0}
        for k, v in counter.items()
    ])
# === Excel Processor ===
def process_excel_file(input_path, output_path, include_embedded=False):
    xls = pd.read_excel(input_path, sheet_name=None)
    all_results = []

    for sheet_name, df in xls.items():
        if "English Sentence" not in df.columns:
            print(f"⚠️ Skipping sheet '{sheet_name}' — no 'English Sentence' column found.")
            continue

        sentences = df["English Sentence"].dropna().astype(str).tolist()
        result_rows = analyze_sentences_to_rows(sentences, include_embedded=include_embedded)
        for row in result_rows:
            row["Sheet"] = sheet_name
        all_results.extend(result_rows)

    if all_results:
        final_df = pd.DataFrame(all_results)
        taxis_df, logic_df, taxis_marker_df, logic_marker_df = compute_summary_stats(final_df)

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            final_df.to_excel(writer, sheet_name="Main_Clauses", index=False)
            taxis_df.to_excel(writer, sheet_name="Taxis_Summary", index=False)
            logic_df.to_excel(writer, sheet_name="LSR_Summary", index=False)
            taxis_marker_df.to_excel(writer, sheet_name="Taxis_Marker_Summary", index=False)
            logic_marker_df.to_excel(writer, sheet_name="Logico_Marker_Summary", index=False)

        print(f"✅ Output with summaries saved to: {output_path}")
    else:
        print("⚠️ No valid data found to process.")

# -----------------------------
# MAIN: CLI support
# -----------------------------
if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python analysis_eng_excel.py input.xlsx [output_folder]")
        sys.exit(1)

    input_excel = sys.argv[1]

    if not os.path.exists(input_excel):
        print(f"❌ File not found: {input_excel}")
        sys.exit(1)

    # Optional: user can provide output folder, else default to "output_files" in script dir
    if len(sys.argv) >= 3:
        output_dir = sys.argv[2]
    else:
        output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output_files")

    # Make sure the output directory exists
    os.makedirs(output_dir, exist_ok=True)

    # Use the same input filename, but save in the output directory
    base_name = os.path.basename(input_excel)
    output_excel = os.path.join(output_dir, f"processed_{base_name}")

    print(f"✅ Output will be saved to: {output_excel}")

    process_excel_file(input_excel, output_excel)
