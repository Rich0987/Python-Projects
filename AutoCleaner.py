import csv
import os

# Ask user for CSV file
csv_path = input("Enter the full path of the CSV file to clean: ")

if not os.path.isfile(csv_path):
    print("File not found.")
    exit()

# Load the data
with open(csv_path, newline='', encoding='utf-8') as f:
    reader = csv.reader(f)
    rows = list(reader)

if not rows:
    print("The file is empty.")
    exit()

# Step 1: Clean header
headers = [col.replace("shell", "").strip() for col in rows[0]]

# Remove column D if exists (header named 'D')
try:
    idx_D = headers.index('D')
    headers.pop(idx_D)
except ValueError:
    idx_D = None

# Step 2: Clean and prepare data rows
cleaned_rows = [headers]
for row in rows[1:]:
    # Remove column D by index if it exists
    if idx_D is not None and len(row) > idx_D:
        row.pop(idx_D)

    # Ensure the row has enough columns (pad if needed)
    if len(row) < len(headers):
        row += [''] * (len(headers) - len(row))

    # Clean each cell
    cleaned = [cell.replace("ILG1-", "")
                   .replace(".ilg1", "")
                   .replace(".com", "")
                   .replace(".vrsn", "") for cell in row]
    cleaned_rows.append(cleaned)

# Step 3: Sort by column C then D (i.e., index 2 then 3)
data_rows = cleaned_rows[1:]

try:
    data_rows.sort(key=lambda x: (x[2].lower(), x[3].lower()))
except IndexError:
    print("Not enough columns (need at least 4) to sort by column C and D. Skipping sort.")

# Rebuild final list
final_rows = [headers] + data_rows

# Step 4: Get filename from N2 (row 2, column 14)
try:
    n2_value = final_rows[1][13]  # Column N is index 13 (0-based)
    if not n2_value.strip():
        raise ValueError
except (IndexError, ValueError):
    n2_value = "cleaned_output"

output_filename = f"{n2_value}-cleaned.csv"
output_path = os.path.join(os.path.dirname(csv_path), output_filename)

# Step 5: Write cleaned CSV
with open(output_path, "w", newline='', encoding='utf-8') as f:
    writer = csv.writer(f)
    writer.writerows(final_rows)

print(f"✅ Cleaned and sorted CSV saved to: {output_path}")
