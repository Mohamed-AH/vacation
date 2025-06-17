import csv
from datetime import datetime

# Define input and output files
input_file = 'filtered_output.csv'
output_file = 'approved_2025_output.csv'

# Get current year for filtering
current_year = datetime.now().year  # Should be 2025 as of now

# Read header manually to avoid issues
header = None
rows_to_keep = []

with open(input_file, mode='r', encoding='utf-8') as infile:
    reader = csv.reader(infile)
    for row in reader:
        if not header:
            header = row
            continue  # Skip header

        # Extract relevant fields
        status = row[9].strip()  # "Status" is column index 9
        start_date = row[5].strip()  # "Start Date" is column index 5

        try:
            # Try parsing date in format like '31/05/2025' or '2025-05-31'
            start_year = int(start_date.split("/")[-1].split(" ")[0])
        except:
            print(f"Skipping invalid date format: {start_date}")
            continue

        # Apply filters
        if status == "Approved" and start_year == current_year:
            rows_to_keep.append(row)

# Write filtered data to new CSV
with open(output_file, mode='w', newline='', encoding='utf-8') as outfile:
    writer = csv.writer(outfile)
    writer.writerow(header)  # Write header
    writer.writerows(rows_to_keep)  # Write filtered rows

print(f"\n✅ Filtering complete.")
print(f"- Total approved records for {current_year}: {len(rows_to_keep)}")
print(f"- Filtered output saved to '{output_file}'")