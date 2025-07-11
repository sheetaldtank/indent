import pandas as pd

def expand_ranges(cell_value):
    """
    Expand comma-separated values and ranges into individual numbers.
    e.g., "122, 5582-5589" → "122; 5582; 5583; ...; 5589"
    """
    if pd.isna(cell_value):
        return ''
    
    parts = [part.strip() for part in str(cell_value).split(',')]
    expanded = []
    
    for part in parts:
        if '-' in part:
            start, end = map(int, part.split('-'))
            expanded.extend(map(str, range(start, end + 1)))
        else:
            expanded.append(part)
    
    return '; '.join(expanded)

# === Configuration ===
input_file = 'indent.xlsx'     # Path to your input Excel file
output_file = 'indent_op.xlsx'   # Path to save the updated Excel file
column_to_convert = 'A'       # Column letter or index to convert (e.g., 'A' or 0)
output_column = 'B'           # Column letter or name for result

# === Load Excel File ===
df = pd.read_excel(input_file)

# Convert column letter to index if needed
if isinstance(column_to_convert, str):
    col_index = ord(column_to_convert.upper()) - ord('A')
else:
    col_index = column_to_convert

# Get column name
col_name = df.columns[col_index]

# Apply conversion
df[output_column] = df[col_name].apply(expand_ranges)

# Save to new Excel file
df.to_excel(output_file, index=False)

print(f"Processed Excel saved as '{output_file}'")
