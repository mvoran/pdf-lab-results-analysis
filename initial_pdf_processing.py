import fitz  # PyMuPDF
import re
import pandas as pd
from collections import defaultdict
from dateutil.parser import parse

def extract_text_from_pdf(filepath):
    """Extract all text from a PDF file using PyMuPDF."""
    doc = fitz.open(filepath)
    text = ""
    for page in doc:
        text += page.get_text()
    doc.close()
    return text

# List your PDF file paths here.
filepaths = [
    "/mnt/data/Scan - COMPLETE BLOOD COUNT WITH AUTO DIFFERENTIAL - Mar 31, 2025.PDF",
    "/mnt/data/Scan - COMPREHENSIVE METABOLIC PANEL - Mar 31, 2025.PDF",
    "/mnt/data/Labcorp_20250306.pdf",
    "/mnt/data/Labcorp_20230906.pdf",
    "/mnt/data/Scan - Combined - Mar 31, 2025 copy.PDF",
    "/mnt/data/Labcorp_20230728.pdf",
]

# Read all PDFs and store text in a dictionary keyed by filename.
pdf_texts = {}
for fp in filepaths:
    pdf_texts[fp] = extract_text_from_pdf(fp)

# We'll assume that each PDF has a table of lab results formatted roughly as:
#   Test Name
#   Normal Range: [range]
#   [Value]   (possibly more than one value if multiple dates are present)
#
# This regex attempts to capture that structure.
result_pattern = re.compile(
    r"(?P<Test>[A-Za-z0-9 \(\)\[\]\-\/]+)\nNormal Range:\s*(?P<Range>[><=0-9.\-–\s]+[a-zA-Z/%²μ]+)\n(?P<Value>[><=0-9. \s]+[a-zA-Z/%²μ]+)",
    re.MULTILINE
)

# For date extraction, we assume that a full date string like "Mar 31, 2025" appears either in the filename or in the text.
# Here, we define a function to extract the first date we find in the filename or text.
full_date_pattern = re.compile(r"((?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2},\s+\d{4})")

def extract_date(filepath, text):
    # Try to extract a full date from the filename
    m = full_date_pattern.search(filepath)
    if m:
        return m.group(1)
    # Otherwise, look in the text
    m = full_date_pattern.search(text)
    if m:
        return m.group(1)
    return "Unknown Date"

# Now, process each PDF, extract lab results and associate them with a date.
lab_data = defaultdict(dict)
reference_ranges = {}

for fp, text in pdf_texts.items():
    date = extract_date(fp, text)
    for match in result_pattern.finditer(text):
        test = match.group("Test").strip()
        rng = match.group("Range").strip()
        value = match.group("Value").strip()
        # If the test already has a value for this date, we leave it (or update as needed)
        lab_data[test][date] = value
        reference_ranges[test] = rng

# Create a DataFrame from lab_data.
df = pd.DataFrame(lab_data).T
df["Reference Range"] = pd.Series(reference_ranges)
df.reset_index(inplace=True)
df.rename(columns={"index": "Test"}, inplace=True)

# Pivot the data so that each row is a test, each unique date is a column,
# and "Reference Range" remains as the last column.
# First, we ensure that each test row has entries for all dates.
all_dates = set()
for test, vals in lab_data.items():
    all_dates.update(vals.keys())
all_dates = sorted(list(all_dates), key=lambda d: parse(d) if d != "Unknown Date" else parse("1/1/1900"))

# Create a new DataFrame with one row per test and one column per date.
rows = []
for test in df["Test"]:
    row = {"Test": test, "Reference Range": reference_ranges.get(test, "")}
    for d in all_dates:
        row[d] = lab_data[test].get(d, "")
    rows.append(row)
combined_df = pd.DataFrame(rows)

# Reorder columns: Test, date columns (in mm/dd/yyyy format), then Reference Range.
def reformat_date(date_str):
    try:
        dt = parse(date_str)
        return dt.strftime("%m/%d/%Y")
    except Exception as e:
        return date_str

date_cols_formatted = sorted(all_dates, key=lambda d: parse(d) if d != "Unknown Date" else parse("1/1/1900"))
date_cols_formatted = [reformat_date(d) for d in date_cols_formatted]
# Rename the date columns in the DataFrame accordingly.
rename_mapping = {old: reformat_date(old) for old in all_dates}
combined_df.rename(columns=rename_mapping, inplace=True)
ordered_columns = ["Test"] + date_cols_formatted + ["Reference Range"]
combined_df = combined_df[ordered_columns]

# Save the raw combined data to an intermediate Excel file.
raw_combined_excel = "Combined_Lab_Results_Raw.xlsx"
combined_df.to_excel(raw_combined_excel, index=False)
print("Raw combined data saved to", raw_combined_excel)

# (Optional) At this point, you could inspect the raw Excel output.

# The next step would be deduplication and conditional formatting, but this code snippet
# covers reading files, extracting lab results, and adding them to a combined worksheet.