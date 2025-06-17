import csv
import os

# Step 1: Load all employee codes from employee.csv
employee_codes = set()

with open('employee.csv', mode='r', encoding='utf-8') as emp_file:
    reader = csv.reader(emp_file)
    for row in reader:
        if len(row) >= 3:
            code = row[2].strip().strip('"')
            if code.isdigit() or code.startswith(("60", "12")):
                employee_codes.add(code)

print(f"Total employee codes in employee.csv: {len(employee_codes)}")

# Step 2: Load employee codes present in filtered_output.csv
present_in_output = set()

output_file = 'filtered_output.csv'

if not os.path.exists(output_file):
    print(f"\n❌ File '{output_file}' does not exist.")
else:
    with open(output_file, mode='r', encoding='utf-8') as out_file:
        reader = csv.DictReader(out_file)
        for row in reader:
            code = row.get("Employee Code", "").strip()
            if code in employee_codes:
                present_in_output.add(code)

    # Step 3: Find missing employees
    missing_employees = employee_codes - present_in_output

    print(f"\n✅ Records found for {len(present_in_output)} employees in {output_file}.")
    print(f"❌ Missing employees: {len(missing_employees)}")

    if missing_employees:
        print("\n🚫 These employee codes are missing in filtered_output.csv:")
        for code in sorted(missing_employees):
            print(code)
    else:
        print("\n🎉 All employees from employee.csv are present in filtered_output.csv!")