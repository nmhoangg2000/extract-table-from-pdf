import pandas as pd
from tabula.io import read_pdf
import time
# Path to your PDF file
pdf_path = 'E:\mcc\extract_table\CIC-Dung-quất-2.2024.pdf'

# Use lattice mode, which is more suitable for tables with clear grid lines
tables = read_pdf(pdf_path, pages='all', multiple_tables=True, lattice=True)

# Assuming you want to preprocess tables as mentioned before
def preprocess_table(table):
    # Drop unnamed columns
    return table.loc[:, ~table.columns.str.contains('^Unnamed')]
start = time.time()
# Preprocess each table to remove unnamed columns
preprocessed_tables = [preprocess_table(table) for table in tables if not table.empty]

# Logic to merge tables with the same headers
final_tables = []
for table in preprocessed_tables:
    if not final_tables:
        final_tables.append(table)
    else:
        # Compare headers and merge tables if they have the same headers
        if list(table.columns) == list(final_tables[-1].columns):
            final_tables[-1] = pd.concat([final_tables[-1], table], ignore_index=True)
        else:
            final_tables.append(table)

# Write merged tables to Excel
with pd.ExcelWriter('merged_tables_lattice1.xlsx') as writer:
    for i, final_table in enumerate(final_tables):
        final_table.to_excel(writer, sheet_name=f'Table {i+1}')
end = time.time()
print("The tables have been processed and saved to 'merged_tables_lattice1.xlsx'.")
print(f'time {end-start}')

