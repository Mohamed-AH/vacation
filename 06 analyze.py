import csv
from datetime import datetime

# Define date range to check
start_check = datetime.strptime("14/06/2025", "%d/%m/%Y")
end_check = datetime.strptime("20/06/2025", "%d/%m/%Y")

# Load employee codes from employee.csv
employee_codes = set()

with open("employee.csv", mode="r", encoding="utf-8") as emp_file:
    reader = csv.reader(emp_file)
    for row in reader:
        if len(row) >= 3:
            code = row[2].strip().strip('"')
            if code.isdigit() or code.startswith(("60", "12")):
                employee_codes.add(code)

print(f"Loaded {len(employee_codes)} employee codes.")

# Load approved leaves from approved_2025_output.csv
leave_records = {}

with open("approved_2025_output.csv", mode="r", encoding="utf-8") as out_file:
    reader = csv.DictReader(out_file)
    for row in reader:
        emp_code = row["Employee Code"]
        pid = row["PID"]
        start_date = datetime.strptime(row["Start Date"], "%d/%m/%Y")
        end_date = datetime.strptime(row["End Date"], "%d/%m/%Y")

        if emp_code not in leave_records:
            leave_records[emp_code] = []

        leave_records[emp_code].append({
            "PID": pid,
            "Start Date": start_date,
            "End Date": end_date
        })

print(f"Found leave data for {len(leave_records)} employees.")

# Prepare final output
results = {}

for emp_code in employee_codes:
    results[emp_code] = {"status": "no leave", "details": []}

    if emp_code not in leave_records:
        continue

    overlapping = []
    full_coverage = False

    for record in leave_records[emp_code]:
        s = record["Start Date"]
        e = record["End Date"]

        # Check for overlap
        if not (e < start_check or s > end_check):
            overlapping.append(record)

    if overlapping:
        # Now check if fully covered
        merged_dates = []
        sorted_leaves = sorted(overlapping, key=lambda x: x["Start Date"])

        for interval in sorted_leaves:
            if not merged_dates:
                merged_dates.append(interval)
            else:
                last = merged_dates[-1]
                if interval["Start Date"] <= last["End Date"]:
                    # Overlap, merge
                    new_start = last["Start Date"]
                    new_end = max(last["End Date"], interval["End Date"])
                    merged_dates[-1] = {
                        "Start Date": new_start,
                        "End Date": new_end,
                        "PID": f"{last['PID']}, {interval['PID']}",
                    }
                else:
                    merged_dates.append(interval)

        # Now check if merged coverage includes full period
        full_overlap = False
        for merged in merged_dates:
            if merged["Start Date"] <= start_check and merged["End Date"] >= end_check:
                full_overlap = True
                break

        results[emp_code]["details"] = [
            {
                "PID": r["PID"],
                "Leave Start": r["Start Date"].strftime("%d/%m/%Y"),
                "Leave End": r["End Date"].strftime("%d/%m/%Y")
            }
            for r in overlapping
        ]

        if full_overlap:
            results[emp_code]["status"] = "fully covered"
        else:
            results[emp_code]["status"] = "partially covered"

# Save results to CSV
output_file = "leave_analysis_report.csv"

with open(output_file, mode="w", newline="", encoding="utf-8") as report_file:
    writer = csv.writer(report_file)
    writer.writerow([
        "Employee Code",
        "Status",
        "PID(s)",
        "Leave Start",
        "Leave End"
    ])

    for emp_code, info in results.items():
        if info["status"] == "no leave":
            writer.writerow([emp_code, "no leave", "", "", ""])
        else:
            for entry in info["details"]:
                writer.writerow([
                    emp_code,
                    info["status"],
                    entry["PID"],
                    entry["Leave Start"],
                    entry["Leave End"]
                ])

print(f"\n✅ Leave analysis completed.")
print(f"Results saved to '{output_file}'")