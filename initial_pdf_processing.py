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
    print(f"\nProcessing PDF with {len(doc)} pages")
    for page_num, page in enumerate(doc):
        page_text = page.get_text()
        print(f"\nPage {page_num + 1} content length: {len(page_text)}")
        text += page_text + "\n" + "="*80 + "\n"  # Add page separator
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
    try:
        # Split text into lines and remove empty lines
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        print(lines)

        print("\nDebug: First few lines of text:")
        for i, line in enumerate(lines[:10]):
            print(f"Line {i}: {line}")
            
        print("\nDebug: Total number of lines:", len(lines))
        
        # Find all Component header rows
        component_indices = []
        for i, line in enumerate(lines):
            if line.startswith("Component"):
                component_indices.append(i)
                print(f"Found Component header at line {i}")
        
        if not component_indices:
            print("Error: No 'Component' header rows found in text.")
            return None, None
        
        print(f"\nDebug: Found {len(component_indices)} Component rows at indices: {component_indices}")
        
        # Process each section
        all_data_rows = []
        all_date_headers = set()
        
        for section_start in component_indices:
            print(f"\nProcessing section starting at line {section_start}")
            
            # Look for dates in subsequent lines until we hit a non-date line
            date_headers = []
            date_pattern = re.compile(r"((?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2},\s+\d{4})")
            
            # Start looking at the line after Component
            current_line = section_start + 1
            while current_line < len(lines):
                line = lines[current_line]
                print(f"\nDebug: Checking line {current_line}: {line}")
                # Check if this line contains a date
                if date_pattern.search(line):
                    date_headers.append(line.strip())
                    print(f"Found date: {line.strip()}")
                    current_line += 1
                else:
                    print("No date found in this line, stopping date search")
                    # If we hit a non-date line, we've found all the date headers
                    break
            
            if not date_headers:
                print("Error: No date headers found after Component row.")
                continue
            
            print(f"Found date headers: {date_headers}")
            all_date_headers.update(date_headers)
            
            # Now process the data rows (starting after the date headers)
            data_rows = []
            print(f"\nDebug: Processing data rows starting from line {current_line}")
            
            # Initialize variables for tracking the current component being processed
            current_component = None      # The name of the current lab test (e.g., "Chloride", "CO2")
            current_values = []          # List of test results for the current component
            current_range = None         # Reference range for the current component
            value_count = 0              # Number of values collected for current component
            i = current_line
            
            # Find the next section start or end of file
            next_section_start = None
            for next_idx in component_indices:
                if next_idx > section_start:
                    next_section_start = next_idx
                    break
            
            # Process until we hit the next section or end of file
            while i < (next_section_start if next_section_start else len(lines)):
                line = lines[i]
                
                # CASE 1: Found a new Component header row
                # This marks the start of a new section in the PDF
                if line.startswith("Component"):
                    # Save the previous component's data if we have any
                    if current_component and len(current_values) > 0:
                        # Pad values with empty strings if we don't have enough
                        while len(current_values) < len(date_headers):
                            current_values.append("")
                        row = [current_component] + current_values + [current_range]
                        data_rows.append(row)
                        current_values = []
                        value_count = 0
                        current_range = None
                    current_component = line
                    i += 1
                    continue
                
                # CASE 2: Found a date header
                # Skip these as they're just column headers
                if date_pattern.search(line):
                    i += 1
                    continue
                
                # CASE 3: Found a new component name
                # This is a line that doesn't contain numbers and isn't a normal range
                # Examples: "Chloride", "CO2", "Sodium"
                if not line.startswith("Normal Range:") and not any(c.isdigit() for c in line):
                    # Skip if this is just a unit (like "m2")
                    if line.strip().lower() in ['m2']:
                        i += 1
                        continue
                        
                    # Save the previous component's data if we have any
                    if current_component and len(current_values) > 0:
                        # Pad values with empty strings if we don't have enough
                        while len(current_values) < len(date_headers):
                            current_values.append("")
                        row = [current_component] + current_values + [current_range]
                        data_rows.append(row)
                        current_values = []
                        value_count = 0
                        current_range = None
                    
                    # Start processing the new component
                    current_component = line
                    current_values = []
                    current_range = None
                    i += 1
                    continue

                # CASE 4: Found a reference range
                # This starts with "Normal Range:" and may span multiple lines
                # Example: "Normal Range: 21 - 31 mmol/L"
                elif line.startswith("Normal Range:"):
                    # Start collecting the range text, removing the "Normal Range:" prefix
                    range_text = line.replace("Normal Range:", "").strip()
                    
                    # Check subsequent lines to see if the range continues
                    while i + 1 < len(lines):
                        next_line = lines[i + 1].strip()
                        # Stop if we hit a new component or a value (but not a unit)
                        if next_line.startswith("Component") or (any(c.isdigit() for c in next_line) and not next_line.lower() in ['m2']):
                            break
                        # Add the next line to the range if it's not another range
                        if not next_line.startswith("Normal Range:"):
                            range_text += " " + next_line
                        i += 1
                    current_range = range_text
                    print(f"Debug: Found reference range for {current_component}: {current_range}")
                    i += 1
                    continue

                # CASE 5: Found a test result value
                # This is a line containing numbers that isn't part of a reference range
                # Examples: "103 mmol/L", "30 mmol/L"
                elif any(c.isdigit() for c in line):
                    # Skip if this is just a unit (like "m2")
                    if line.strip().lower() in ['m2'] or not any(c.isdigit() for c in line):
                        i += 1
                        continue
                    # Skip if this is part of the reference range
                    if current_range and line.strip() in current_range:
                        i += 1
                        continue
                    # Skip if this is a new component name
                    if not any(c.isdigit() for c in line) and not line.startswith("Normal Range:"):
                        i += 1
                        continue
                    # Only add if it's a valid value (contains a number and not part of a range)
                    if any(c.isdigit() for c in line) and not line.startswith("Normal Range:"):
                        # Special case: Check if this line is actually a new component (like CO2)
                        if line.strip().upper() in ['CO2']:
                            # Save the previous component's data if we have any
                            if current_component and len(current_values) > 0:
                                # Pad values with empty strings if we don't have enough
                                while len(current_values) < len(date_headers):
                                    current_values.append("")
                                row = [current_component] + current_values + [current_range]
                                data_rows.append(row)
                                current_values = []
                                value_count = 0
                                current_range = None
                            # Start processing the new component
                            current_component = line
                            current_values = []
                            current_range = None
                            i += 1
                            continue
                        # Add the value to the current component's results
                        current_values.append(line.strip())
                        value_count += 1
                        print(f"Debug: Added value {line.strip()} for component {current_component}")
                i += 1
            
            # Add the data rows from this section
            all_data_rows.extend(data_rows)
        
        # Convert all date headers to a sorted list
        all_date_headers = sorted(list(all_date_headers))
        
        print(f"Found {len(all_data_rows)} total valid data rows")
        for row in all_data_rows:
            print(f"Debug: Data row: {row}")
        
        # Create DataFrame with Test, date columns, and Reference Range
        columns = ['Test'] + all_date_headers + ['Reference Range']
        df = pd.DataFrame(all_data_rows, columns=columns)
        
        return df, all_date_headers
    except Exception as e:
        print(f"Error in extract_table_data_scan: {str(e)}")
        import traceback
        print(traceback.format_exc())
        return None, None

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
            test = row['Test'] #changed this
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
        row = {"Test": test} #changed this from Component to Test
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
    ordered_columns = ["Test"] + date_cols_formatted + ["Reference Range"] #changed this from Component to Test
    final_df = final_df[ordered_columns]
    
    # Save raw combined data to Excel (intermediate output)
    raw_output_path = args.output
    final_df.to_excel(raw_output_path, index=False)
    print(f"\nCombined lab results saved to: {raw_output_path}")

if __name__ == "__main__":
    main()