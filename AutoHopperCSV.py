import csv
import re

# File paths (replace with your actual paths)
CLEANED_PATH = 'cleaned.csv'
HOPS_PATH = 'hops.csv'
OUTPUT_PATH = 'verisign.csv'

# Load cleaned.csv
with open(CLEANED_PATH, newline='', encoding='utf-8') as f1:
    cleaned_data = list(csv.reader(f1))

# Load hops.csv
with open(HOPS_PATH, newline='', encoding='utf-8') as f2:
    hops_data = list(csv.reader(f2))

# Make sure hops_data has enough rows
while len(hops_data) < len(cleaned_data):
    hops_data.append([''] * len(hops_data[0]))

# -----------------------------
# COPY C → C, F → N, D → B, E → M, G → E, H → P, M → Q
# (use zero-based indexing)
for i in range(1, len(cleaned_data)):
    row_clean = cleaned_data[i]
    row_hops = hops_data[i]

    # Extend row if needed
    while len(row_hops) < 20:
        row_hops.extend([''] * (20 - len(row_hops)))

    row_hops[2]  = row_clean[2]  # C → C
    row_hops[13] = row_clean[5]  # F → N
    row_hops[1]  = row_clean[3]  # D → B
    row_hops[12] = row_clean[4]  # E → M
    row_hops[4]  = row_clean[6]  # G → E
    row_hops[15] = row_clean[7]  # H → P
    row_hops[16] = row_clean[12] # M → Q

# -----------------------------
# APPEND A#, P1#, P2#, Z# if column B (index 1) is non-empty
for i in range(1, len(hops_data)):
    row = hops_data[i]
    if row[1].strip():
        row[0]  = 'A#'   # A
        row[3]  = 'P1#'  # D
        row[8]  = 'P2#'  # I
        row[11] = 'Z#'   # L

# -----------------------------
# COPY first 10 characters: C → G (6), N → K (10)
for i in range(1, len(hops_data)):
    row = hops_data[i]
    c_val = row[2]
    n_val = row[13]

    row[6] = c_val[:10] if c_val else ''
    row[10] = n_val[:10] if n_val else ''

# -----------------------------
# APPEND POD patch locations to column G based on K
for i in range(1, len(hops_data)):
    row = hops_data[i]
    k_val = row[10]

    if re.match(r'^RM2602', k_val):
        row[6] += 'C10.U45'
    elif re.match(r'^RM2606', k_val):
        row[6] += 'C11.U45'

# -----------------------------
# APPEND network patch cab and U heights to column K based on G
patch_map = {
    'RM2302-R1': 'C17.U42',
    'RM2302-R2': 'C17.U37',
    'RM2302-R3': 'C17.U32',
    'RM2302-R4': 'C17.U27',
    'RM2302-R5': 'C17.U22',
    'RM2302-R6': 'C17.U17',
    'RM2802-R1': 'C19.U42',
    'RM2802-R2': 'C19.U37',
    'RM2802-R3': 'C19.U32',
    'RM2802-R4': 'C19.U27',
    'RM2802-R5': 'C19.U22',
    'RM2802-R6': 'C19.U17',
    'RM2304-R1': 'C21.U42',
    'RM2304-R2': 'C21.U37',
    'RM2304-R3': 'C21.U32',
    'RM2304-R4': 'C21.U27',
    'RM2304-R5': 'C21.U22',
    'RM2304-R6': 'C21.U17',
    'RM2602-R1': 'C26.U32',
    'RM2606-R1': 'C26.U32'
}

for i in range(1, len(hops_data)):
    row = hops_data[i]
    g_val = row[6]
    for prefix, suffix in patch_map.items():
        if g_val.startswith(prefix):
            row[10] += suffix
            break

# -----------------------------
# Save to verisign.csv
with open(OUTPUT_PATH, 'w', newline='', encoding='utf-8') as fout:
    writer = csv.writer(fout)
    writer.writerows(hops_data)

print("Saved to verisign.csv")
