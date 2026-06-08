import pandas as pd
import os
from datetime import datetime

FILE = "OLT_DATA1.xlsx"
BACKUP_FMT = "OLT_DATA1.backup.{ts}.xlsx"

if not os.path.exists(FILE):
    print(f"File not found: {FILE}")
    raise SystemExit(1)

print(f"Reading {FILE}...")
df = pd.read_excel(FILE)
print(f"Columns: {list(df.columns)}")

if 'NE' not in df.columns:
    print("Column 'NE' not found in the Excel file. Aborting.")
    raise SystemExit(1)

# Normalize NE for deduplication
ne_norm = df['NE'].astype(str).str.strip().str.lower()
df['_NE_norm'] = ne_norm

dup_mask = df['_NE_norm'].duplicated(keep='first')
num_dups = dup_mask.sum()

timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
backup_name = BACKUP_FMT.replace('{ts}', timestamp)

print(f"Found {num_dups} duplicate rows based on normalized 'NE'.")
print(f"Backing up original file to {backup_name}...")

# Save backup
pd.ExcelWriter(backup_name, engine='openpyxl')
# using pandas to_excel to create backup
_df_orig = pd.read_excel(FILE)
_df_orig.to_excel(backup_name, index=False)

if num_dups == 0:
    print("No duplicates to remove. No changes made.")
    raise SystemExit(0)

# Remove duplicates, keep first occurrence
cleaned = df[~dup_mask].copy()
# Drop helper column
cleaned.drop(columns=['_NE_norm'], inplace=True)

print(f"Writing cleaned data back to {FILE} (overwriting)...")
try:
    cleaned.to_excel(FILE, index=False)
    print(f"Done. Removed {num_dups} duplicate rows. Original backed up as {backup_name}.")
except PermissionError:
    alt_name = f"OLT_DATA1.cleaned.{timestamp}.xlsx"
    cleaned.to_excel(alt_name, index=False)
    print(f"Permission denied writing {FILE}. Saved cleaned copy as {alt_name}. Original backed up as {backup_name}.")
