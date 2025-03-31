import fitz  # PyMuPDF
import re
import pandas as pd
from collections import defaultdict
from dateutil.parser import parse
import os
import argparse
from pathlib import Path

def extract_text_from_pdf(filepath):
    """Extract all text from a PDF file using PyMuPDF."""
    doc = fitz.open(filepath)
    text = ""
    for page in doc:
        text += page.get_text()
    doc.close()
    return text

def get_pdf_files(directory):
    """Get all PDF files from the specified directory."""
    pdf_files = []
    for file in os.listdir(directory):
        if file.lower().endswith('.pdf'):
            pdf_files.append(os.path.join(directory, file))
    return pdf_files

def extract_table_data_scan(text):
    """
    Extract table data from Scan files.
    This function now looks for dates in the header row.
    The header row is expected to begin with "Component" followed by one or more date strings.
    """
    # Split text into lines and remove empty lines
    lines = [line.strip() for line in text.split('\n') if line.strip()]

    # Split the extracted text into lines.
    lines2 = text.splitlines()

    # Print lines that contain the dates of interest.
    for i, line in enumerate(lines2):
        if "Feb 9, 2024" in line or "Oct 16, 2024" in line or "Component" in line:
            print(f"Line {i}: {line}")
        
    # Find the header row that starts with "Component"
    header_row_index = None
    for i, line in enumerate(lines):
        if line.startswith("Component"):
            header_row_index = i
            break
    
    if header_row_index is None:
        print("Error: 'Component' header row not found in text.")
        return None, None
    
    # Split the header row by whitespace (assuming columns are separated by multiple spaces or tabs)
    headers = re.split(r'\s{2,}', lines[header_row_index])
    print(lines[header_row_index])
    # The first token should be "Component", and the rest are date headers.
    if len(headers) < 2:
        print("Error: No date columns found in header row.")
        return None, None
    
    date_headers = headers[1:]
    print("Extracted date headers from Scan file:", date_headers)
    
    # Now process the subsequent rows.
    # We assume that each row in the table has the same number of columns as the header row.
    data_rows = []
    for line in lines[header_row_index+1:]:
        # Use 2+ spaces as delimiter; adjust as needed for your file.
        tokens = re.split(r'\s{2,}', line)
        if len(tokens) < len(headers):
            continue  # skip rows that don't have enough columns
        data_rows.append(tokens)
    
    # Create a DataFrame using the header row as column names.
    df = pd.DataFrame(data_rows, columns=headers)
    return df, date_headers

def extract_table_data_other(text):
    """
    Extract table data from non-Scan files using a regex-based approach.
    This assumes each lab result is in a block:
      Test Name
      Normal Range: [range]
      [Value]
    """
    result_pattern = re.compile(
        r"(?P<Test>[A-Za-z0-9 \(\)\[\]\-\/]+)\nNormal Range:\s*(?P<Range>[><=0-9.\-–\s]+[a-zA-Z/%²μ]+)\n(?P<Value>[><=0-9. \s]+[a-zA-Z/%²μ]+)",
        re.MULTILINE
    )
    
    # Extract a full date from the text using the header (if present)
    full_date_pattern = re.compile(r"((?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2},\s+\d{4})")
    m = full_date_pattern.search(text)
    date = m.group(1) if m else "Unknown Date"
    
    results = []
    for match in result_pattern.finditer(text):
        test = match.group("Test").strip()
        rng = match.group("Range").strip()
        value = match.group("Value").strip()
        results.append({
            'Component': test,
            date: value,
            'Reference Range': rng
        })
    if not results:
        return None, None
    df = pd.DataFrame(results)
    return df, [date]

def process_pdf_file(filepath):
    """
    Process a PDF file based on its filename structure.
    If the filename starts with 'Scan', use the scan-specific extraction,
    otherwise use the generic extraction.
    """
    filename = os.path.basename(filepath)
    text = extract_text_from_pdf(filepath)
    if filename.lower().startswith('scan'):
        return extract_table_data_scan(text)
    else:
        return extract_table_data_other(text)

def main():
    parser = argparse.ArgumentParser(description='Process Labcorp PDF files and extract lab results.')
    parser.add_argument('directory', help='Directory containing PDF files to process')
    parser.add_argument('--output', '-o', default='Combined_Lab_Results_Raw.xlsx',
                      help='Output Excel file name (default: Combined_Lab_Results_Raw.xlsx)')
    args = parser.parse_args()

    filepaths = get_pdf_files(args.directory)
    if not filepaths:
        print(f"No PDF files found in directory: {args.directory}")
        return

    print(f"Found {len(filepaths)} PDF files to process.")

    all_data = []
    for fp in filepaths:
        try:
            df, dates = process_pdf_file(fp)
            if df is not None and dates:
                print(f"Successfully processed: {os.path.basename(fp)}")
                all_data.append((df, dates))
            else:
                print(f"Could not extract table data from: {os.path.basename(fp)}")
        except Exception as e:
            print(f"Error processing {os.path.basename(fp)}: {str(e)}")
    
    if not all_data:
        print("No data was successfully extracted from any PDFs.")
        return

    # Combine data from all files.
    combined_data = defaultdict(dict)
    reference_ranges = {}

    for df, dates in all_data:
        for _, row in df.iterrows():
            test = row['Component']
            for date in dates:
                if pd.notna(row.get(date, "")):
                    combined_data[test][date] = row.get(date, "")
            if 'Reference Range' in row and row['Reference Range']:
                reference_ranges[test] = row['Reference Range']

    all_dates = set()
    for test, vals in combined_data.items():
        all_dates.update(vals.keys())
    # Sort dates chronologically using dateutil.parser.parse
    all_dates = sorted(list(all_dates), key=lambda d: parse(d) if d != "Unknown Date" else parse("1/1/1900"))
    
    # Build rows for the final DataFrame
    rows = []
    for test in combined_data.keys():
        row = {"Test": test}
        for d in all_dates:
            row[d] = combined_data[test].get(d, "")
        row["Reference Range"] = reference_ranges.get(test, "")
        rows.append(row)
    
    final_df = pd.DataFrame(rows)
    
    # Convert date columns to mm/dd/yyyy format
    def reformat_date(date_str):
        try:
            dt = parse(date_str)
            return dt.strftime("%m/%d/%Y")
        except Exception as e:
            return date_str
    date_cols_formatted = [reformat_date(d) for d in all_dates]
    rename_mapping = {old: reformat_date(old) for old in all_dates}
    final_df.rename(columns=rename_mapping, inplace=True)
    
    # Order columns: "Test", sorted date columns, then "Reference Range"
    ordered_columns = ["Test"] + date_cols_formatted + ["Reference Range"]
    final_df = final_df[ordered_columns]
    
    # Save raw combined data to Excel (intermediate output)
    raw_output_path = args.output
    final_df.to_excel(raw_output_path, index=False)
    print(f"\nCombined lab results saved to: {raw_output_path}")

if __name__ == "__main__":
    main()