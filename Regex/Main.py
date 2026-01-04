import os
import re
import pdfplumber
import pandas as pd
import numpy as np
from sklearn.cluster import DBSCAN

# =========================
# PATHS
# =========================
INPUT_DIR = r"C:\Drive_d\Python\F-AI\Padala Hemanth Subbi Reddy\Input"
TEMPLATE_PATH = r"C:\Drive_d\Python\F-AI\Padala Hemanth Subbi Reddy\Output Template.xlsx"
OUTPUT_DIR = r"C:\Drive_d\Python\F-AI\Padala Hemanth Subbi Reddy\Output"

os.makedirs(OUTPUT_DIR, exist_ok=True)

# =========================
# TOTALS EXTRACTION
# =========================
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
# =========================
def extract_cluster_text(pdf_path):
    blocks = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            chars = page.chars
            if not chars:
                continue

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

            if max_cols >= 6:
                points = (points - points.mean(axis=0)) / points.std(axis=0)
                labels = DBSCAN(eps=0.1, min_samples=1).fit_predict(points)
            else:
                x = (points[:, 0] - points[:, 0].mean()) / points[:, 0].std()
                y = (points[:, 1] - points[:, 1].mean()) / points[:, 1].std()
                labels = DBSCAN(eps=0.12, min_samples=1).fit_predict(
                    np.column_stack([x * 3, y])
                )

            clusters = {}
            for lbl, ch in zip(labels, refs):
                if lbl != -1:
                    clusters.setdefault(lbl, []).append(ch)

            for group in clusters.values():
                group = sorted(group, key=lambda c: (c["top"], c["x0"]))
                text = ""
                prev_top = None
                for ch in group:
                    if prev_top is not None and abs(ch["top"] - prev_top) > 3:
                        text += " "
                    text += ch["text"]
                    prev_top = ch["top"]

                text = text.strip()
                if len(text) > 30:
                    blocks.append(text)

    return " | ".join(blocks)

# =========================
# RULE-BASED FIELD EXTRACTION
# =========================
def extract_with_rules(cluster_text):
    text = re.sub(r"\s+", " ", cluster_text.replace("|", " ")).strip()

    def grab(key, max_len=200):
        idx = text.lower().find(key.lower())
        if idx == -1:
            return ""
        return text[idx + len(key): idx + len(key) + max_len].strip()

    data = {}

    # Invoice Type
    data["invoice_type"] = (
        "Tax Invoice"
        if re.search(r"\b(tax|gst|igst|cgst|sgst)\b", text.lower())
        else ""
    )

    # Order / Invoice Numbers
    m = re.search(r"Order (Number|Id)[:\s]*([\w\-]+)", text, re.IGNORECASE)
    data["order_number"] = m.group(2) if m else ""

    m = re.search(r"Invoice (No|Number)[:\s]*([\w\-]+)", text, re.IGNORECASE)
    data["invoice_number"] = m.group(2) if m else ""

    # Dates
    date_pattern = r"\d{2}[./-]\d{2}[./-]\d{4}"

    m = re.search(r"Order Date[:\s]*(" + date_pattern + ")", text, re.IGNORECASE)
    data["order_date"] = m.group(1) if m else ""

    m = re.search(r"Invoice Date[:\s]*(" + date_pattern + ")", text, re.IGNORECASE)
    data["invoice_date"] = m.group(1) if m else ""

    # Invoice Details
    data["invoice_details"] = grab("Invoice Details", 120)

    # Addresses
    data["shipping_address"] = grab("Shipping Address", 220)
    data["billing_address"] = grab("Billing Address", 220)

    # Seller
    seller_block = grab("Sold By", 300)
    data["seller_info"] = seller_block
    data["seller_name"] = seller_block.split(",")[0] if seller_block else ""
    data["seller_address"] = seller_block

    # GST / PAN / FSSAI
    gst = re.search(r"\b\d{2}[A-Z]{5}\d{4}[A-Z][1-9A-Z]Z[0-9A-Z]\b", text)
    data["seller_gst"] = gst.group(0) if gst else ""

    pan = re.search(r"\b[A-Z]{5}\d{4}[A-Z]\b", text)
    data["seller_pan"] = pan.group(0) if pan else ""

    fssai = re.search(r"\b\d{14}\b", text)
    data["fssai_license"] = fssai.group(0) if fssai else ""

    # Supply / Delivery
    data["place_of_supply"] = grab("Place of supply", 60)
    data["place_of_delivery"] = grab("Place of delivery", 60)

    sc = re.search(r"State/UT Code[:\s]*(\d{1,2})", text)
    data["billing_state_code"] = sc.group(1) if sc else ""
    data["shipping_state_code"] = sc.group(1) if sc else ""

    # Reverse Charge
    data["reverse_charge"] = "No" if "reverse charge" in text.lower() and "no" in text.lower() else ""

    # Amount in Words (STRICT UNTIL 'only')
    amt_words = re.search(
        r"Amount in Words[:\s]*(.*?\bonly\b)",
        text,
        re.IGNORECASE
    )
    data["amount_in_words"] = amt_words.group(1).strip() if amt_words else ""

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
    data = extract_with_rules(cluster_text)

    tax, total = extract_totals(pdf_path)
    data["total_tax"] = tax
    data["total_amount"] = total

    df = pd.read_excel(TEMPLATE_PATH)
    df["Value"] = df["Field"].map(data).fillna("")

    out_path = os.path.join(OUTPUT_DIR, file.replace(".pdf", "_output.xlsx"))
    df.to_excel(out_path, index=False)

    print(f"✔ Saved: {out_path}")
