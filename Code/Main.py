import os
import re
import json
import pdfplumber
import pandas as pd
import numpy as np
from sklearn.cluster import DBSCAN
from langchain_ollama import OllamaLLM

# =========================
# PATHS
# =========================
INPUT_DIR = r"C:\Drive_d\Python\F-AI\Padala Hemanth Subbi Reddy\Input"
TEMPLATE_PATH = r"C:\Drive_d\Python\F-AI\Padala Hemanth Subbi Reddy\Output Template.xlsx"
OUTPUT_DIR = r"C:\Drive_d\Python\F-AI\Padala Hemanth Subbi Reddy\Output"
DEBUG_DIR = os.path.join(OUTPUT_DIR, "debug")

os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(DEBUG_DIR, exist_ok=True)

# =========================
# LOAD LLM (MISTRAL)
# =========================
def load_llm():
    return OllamaLLM(
        model="mistral",
        temperature=0.7
    )

import pdfplumber
import pandas as pd
import re


def extract_totals(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            if not tables:
                continue

            table = max(tables, key=len)

            # Detect table style
            max_cols = max(len([c for c in row if c]) for row in table)

            # =========================
            # AMAZON TABLE
            # =========================
            if max_cols >= 6:
                df = pd.DataFrame(table).dropna(how="all").reset_index(drop=True)

                header = df.iloc[0].astype(str)
                data = df.iloc[1:].reset_index(drop=True)

                tax_col = None
                total_col = None

                for idx, col in enumerate(header):
                    c = col.lower()
                    if "tax" in c and "amount" in c:
                        tax_col = idx
                    if "total" in c and "amount" in c:
                        total_col = idx

                if tax_col is None or total_col is None:
                    return 0.0, 0.0

                # Find TOTAL row
                total_row = None
                for i, row in data.iterrows():
                    if any("total" in str(cell).lower() for cell in row):
                        total_row = row
                        break

                if total_row is None:
                    return 0.0, 0.0

                tax_val = re.findall(r"[\d.]+", str(total_row[tax_col]))
                amt_val = re.findall(r"[\d.]+", str(total_row[total_col]))

                total_tax = float(tax_val[0]) if tax_val else 0.0
                total_amount = float(amt_val[0]) if amt_val else 0.0

                return round(total_tax, 2), round(total_amount, 2)

            # =========================
            # FLIPKART TABLE
            # =========================
            else:
                lines = []
                for row in table:
                    for cell in row:
                        if cell:
                            lines.extend(cell.split("\n"))

                lines = [l.strip() for l in lines if l.strip()]

                igst_vals = []
                total_vals = []

                for line in lines:
                    nums = re.findall(r"\d+\.\d+", line)
                    if len(nums) >= 2:
                        igst_vals.append(float(nums[-2]))
                        total_vals.append(float(nums[-1]))

                return round(sum(igst_vals), 2), round(sum(total_vals), 2)

    return 0.0, 0.0


# =========================
# CLUSTER → SINGLE LINE TEXT
# (AMAZON vs FLIPKART LOGIC)
# =========================
def extract_cluster_text(pdf_path):
    blocks = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            chars = page.chars
            if not chars:
                continue

            # ---------- detect invoice type (same as totals) ----------
            tables = page.extract_tables()
            max_cols = 0
            if tables:
                table = max(tables, key=len)
                max_cols = max(len([c for c in row if c]) for row in table)

            points, refs = [], []
            for ch in chars:
                x = (ch["x0"] + ch["x1"]) / 2
                y = (ch["top"] + ch["bottom"]) / 2
                points.append([x, y])
                refs.append(ch)

            points = np.array(points)

            # =========================
            # AMAZON → NORMALIZED DBSCAN
            # =========================
            if max_cols >= 6:
                points_norm = (points - points.mean(axis=0)) / points.std(axis=0)
                labels = DBSCAN(eps=0.1, min_samples=1).fit_predict(points_norm)

            # =========================
            # FLIPKART → X-BIASED DBSCAN
            # =========================
            else:
                x = (points[:, 0] - points[:, 0].mean()) / points[:, 0].std()
                y = (points[:, 1] - points[:, 1].mean()) / points[:, 1].std()
                points_scaled = np.column_stack([x * 3.0, y])
                labels = DBSCAN(eps=0.12, min_samples=1).fit_predict(points_scaled)

            clusters = {}
            for lbl, ch in zip(labels, refs):
                if lbl == -1:
                    continue
                clusters.setdefault(lbl, []).append(ch)

            for group in clusters.values():
                group = sorted(group, key=lambda c: (c["top"], c["x0"]))
                parts = []
                prev_x1 = None
                prev_top = None

                for ch in group:
                    # Detect new line (vertical jump)
                    if prev_top is not None and abs(ch["top"] - prev_top) > 3:
                        parts.append(" ")   # or "\n" if you prefer

                    parts.append(ch["text"])

                    prev_x1 = ch["x1"]
                    prev_top = ch["top"]

                text = "".join(parts).replace("  ", " ").strip()
                if (
                    len(text) > 30 or
                    re.search(r"\b(FSSAI|LICENSE|GST|PAN)\b", text, re.IGNORECASE) or
                    re.search(r"\b\d{14}\b", text)   # FSSAI number
                ):
                    blocks.append(text)

    return " | ".join(blocks)

# =========================
# ROBUST JSON PARSER
# =========================
def safe_json_parse(text):
    if not text:
        return {}

    # Keep printable characters only
    text = "".join(c for c in text if c.isprintable()).strip()

    # Extract JSON block
    start = text.find("{")
    end = text.rfind("}")

    if start == -1 or end == -1 or end <= start:
        return {}

    json_str = text[start:end + 1]

    # 🔴 CRITICAL FIX: escape invalid backslashes
    json_str = json_str.replace("\\", "\\\\")

    try:
        return json.loads(json_str)
    except Exception as e:
        print("❌ JSON decode error:", e)
        print("RAW JSON STRING:\n", json_str)
        return {}




# =========================
# LLM EXTRACTION
# =========================
def extract_with_llm(cluster_text, file_name):
    llm = load_llm()

    debug_path = os.path.join(DEBUG_DIR, file_name.replace(".pdf", "_cluster.txt"))
    with open(debug_path, "w", encoding="utf-8") as f:
        f.write(cluster_text)

    print("\n--- TEXT SENT TO MODEL ---\n")
    print(cluster_text[:1500])
    print("\n-------------------------\n")

    prompt = f"""
Strictly Return ONLY valid JSON,
No explanations. No markdown.

If missing, try to fetch answer from remaining data. If not found, return empty string "".
the seller name and address are present in Sold by section.

Schema:
{{
  "billing_address": "",
  "shipping_address": "",
  "invoice_type": "",
  "order_number": "",
  "invoice_number": "",
  "order_date": "",
  "invoice_details": "",
  "invoice_date": "",
  "seller_info": "",
  "seller_pan": "",
  "seller_gst": "",
  "fssai_license": "",
  "billing_state_code": "",
  "shipping_state_code": "",
  "place_of_supply": "",
  "place_of_delivery": "",
  "reverse_charge": "",
  "amount_in_words": "",
  "seller_name": "",
  "seller_address": ""
}}

Invoice text:
\"\"\"{cluster_text}\"\"\"
"""

    response = llm.invoke(prompt)
    data = safe_json_parse(response)

    if not data:
        print("⚠ JSON parse failed\n", response)

    return data

# =========================
# MAIN PIPELINE
# =========================
for file in os.listdir(INPUT_DIR):
    if not file.lower().endswith(".pdf"):
        continue

    print(f"\n▶ Processing {file}")
    pdf_path = os.path.join(INPUT_DIR, file)

    cluster_text = extract_cluster_text(pdf_path)
    llm_data = extract_with_llm(cluster_text, file)
    # llm_data["fssai_license"] = extract_fssai_number(pdf_path)

    tax, total = extract_totals(pdf_path)
    llm_data["total_tax"] = tax
    llm_data["total_amount"] = total

    df = pd.read_excel(TEMPLATE_PATH)
    df["Value"] = df["Field"].map(llm_data).fillna("")

    out_path = os.path.join(OUTPUT_DIR, file.replace(".pdf", "_output.xlsx"))
    df.to_excel(out_path, index=False)

    print(f"✔ Saved: {out_path}")
