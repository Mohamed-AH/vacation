import csv

# Step 1: Read employee codes from employee.csv
employee_codes = set()

with open('employee.csv', mode='r', encoding='utf-8') as emp_file:
    reader = csv.reader(emp_file)
    for row in reader:
        if len(row) >= 3:
            code = row[2].strip().strip('"')  # Third column is Employee Code
            if code.isdigit() or code.startswith(("60", "12")):
                employee_codes.add(code)

print(f"Loaded {len(employee_codes)} employee codes.")

# Step 2: Read output.csv and filter raw lines
input_file = 'output.csv'
output_file = 'filtered_output.csv'

# Define header manually since the file is messy
header = ['PID', 'Submission Date', 'Workflow Type', 'Employee Code', 'Employee_Name',
          'Start Date', 'End Date', 'Period', 'Sent To Payroll', 'Status']

valid_rows = []

with open(input_file, mode='r', encoding='utf-8', errors='ignore') as infile:
    for line in infile:
        line = line.strip()
        if not line:
            continue

        # Try to find lines starting with numeric PID and containing at least one employee code
        if any(code in line for code in employee_codes):
            parts = [field.strip('"') for field in line.split('","')]
            if len(parts) >= 10 and parts[3].strip('"') in employee_codes:
                valid_rows.append(parts)

# Write filtered data to new CSV
with open(output_file, mode='w', newline='', encoding='utf-8') as outfile:
    writer = csv.writer(outfile)
    writer.writerow(header)
    writer.writerows(valid_rows)

print(f"\n✅ Filtering complete. Matching records: {len(valid_rows)}")
print(f"Filtered output saved to '{output_file}'")