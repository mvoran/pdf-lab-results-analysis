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
    """Extract table data from Scan files, where dates are in column headers."""
    # Split text into lines and remove empty lines
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    
    print("\nDebug: Looking for dates in the text:")
    for i, line in enumerate(lines):
        if i >= 3 and i <= 5:  # Show lines around line 4
            print(f"Line {i}: {line}")
            if 'Feb 9, 2024' in line or 'Oct 16, 2024' in line:
                print(f"Found target date in line {i}")
    
    # Find the Component row (first row)
    component_row = None
    for i, line in enumerate(lines):
        if 'Component' in line:
            component_row = i
            break
    
    if component_row is None:
        print("\nDebug: Could not find 'Component' in any line")
        return None, None
    
    print(f"\nDebug: Found Component row at index {component_row}")
    
    # Extract components and dates
    components = [item.strip() for item in lines[component_row].split('\t')]
    print(f"\nDebug: Found components: {components}")
    
    dates = []
    values = []
    
    # Process subsequent rows
    for line in lines[component_row + 1:]:
        items = [item.strip() for item in line.split('\t')]
        if len(items) == len(components):
            values.append(items)
    
    print(f"\nDebug: Found {len(values)} data rows")
    
    # Convert to DataFrame
    df = pd.DataFrame(values, columns=components)
    
    # Extract dates from column headers (skip the Component column)
    dates = components[1:]
    print(f"\nDebug: Found dates: {dates}")
    
    return df, dates

def extract_table_data_other(text):
    """Extract table data from non-Scan files using the original regex pattern."""
    # We'll assume that each PDF has a table of lab results formatted roughly as:
    #   Test Name
    #   Normal Range: [range]
    #   [Value]   (possibly more than one value if multiple dates are present)
    result_pattern = re.compile(
        r"(?P<Test>[A-Za-z0-9 \(\)\[\]\-\/]+)\nNormal Range:\s*(?P<Range>[><=0-9.\-–\s]+[a-zA-Z/%²μ]+)\n(?P<Value>[><=0-9. \s]+[a-zA-Z/%²μ]+)",
        re.MULTILINE
    )

    # For date extraction, we assume that a full date string like "Mar 31, 2025" appears either in the filename or in the text.
    full_date_pattern = re.compile(r"((?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2},\s+\d{4})")
    
    # Extract date from text
    m = full_date_pattern.search(text)
    date = m.group(1) if m else "Unknown Date"
    
    # Extract test results
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
    
    # Convert to DataFrame
    df = pd.DataFrame(results)
    return df, [date]

def process_pdf_file(filepath):
    """Process a PDF file based on its filename structure."""
    filename = os.path.basename(filepath)
    text = extract_text_from_pdf(filepath)
    
    if filename.lower().startswith('scan'):
        return extract_table_data_scan(text)
    else:
        return extract_table_data_other(text)

def main():
    # Set up command line argument parsing
    parser = argparse.ArgumentParser(description='Process Labcorp PDF files and extract lab results.')
    parser.add_argument('directory', help='Directory containing PDF files to process')
    parser.add_argument('--output', '-o', default='Combined_Lab_Results_Raw.xlsx',
                      help='Output Excel file name (default: Combined_Lab_Results_Raw.xlsx)')
    args = parser.parse_args()

    # Get all PDF files from the specified directory
    filepaths = get_pdf_files(args.directory)
    
    if not filepaths:
        print(f"No PDF files found in directory: {args.directory}")
        return

    print(f"Found {len(filepaths)} PDF files to process")

    # Process each PDF file
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
        print("No data was successfully extracted from any PDFs")
        return

    # Combine all data
    combined_data = defaultdict(dict)
    reference_ranges = {}

    for df, dates in all_data:
        for _, row in df.iterrows():
            test = row['Component']
            for date in dates:
                if pd.notna(row[date]):  # Only add non-null values
                    combined_data[test][date] = row[date]
            # Store reference range if available
            if 'Reference Range' in row:
                reference_ranges[test] = row['Reference Range']

    # Create final DataFrame
    all_dates = set()
    for test, vals in combined_data.items():
        all_dates.update(vals.keys())
    all_dates = sorted(list(all_dates), key=lambda d: parse(d) if d != "Unknown Date" else parse("1/1/1900"))

    # Create rows for the final DataFrame
    rows = []
    for test in combined_data.keys():
        row = {"Test": test}
        for d in all_dates:
            row[d] = combined_data[test].get(d, "")
        if test in reference_ranges:
            row["Reference Range"] = reference_ranges[test]
        rows.append(row)

    final_df = pd.DataFrame(rows)

    # Reorder columns: Test, date columns (in mm/dd/yyyy format), then Reference Range
    def reformat_date(date_str):
        try:
            dt = parse(date_str)
            return dt.strftime("%m/%d/%Y")
        except Exception as e:
            return date_str

    date_cols_formatted = [reformat_date(d) for d in all_dates]
    rename_mapping = {old: reformat_date(old) for old in all_dates}
    final_df.rename(columns=rename_mapping, inplace=True)
    
    # Determine column order based on whether we have reference ranges
    ordered_columns = ["Test"] + date_cols_formatted
    if "Reference Range" in final_df.columns:
        ordered_columns.append("Reference Range")
    final_df = final_df[ordered_columns]

    # Save the combined data to Excel file
    final_df.to_excel(args.output, index=False)
    print(f"\nCombined lab results saved to: {args.output}")

if __name__ == "__main__":
    main()