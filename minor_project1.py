#!/usr/bin/env python
# coding: utf-8

# In[20]:


import openpyxl
file_path=r"C:\Users\Hi\OneDrive\Desktop\CAP776\12628918.xlsx"
wb= openpyxl.load_workbook(file_path)
ws = wb.active

from openpyxl import load_workbook
from datetime import date, datetime
from collections import Counter
from statistics import mean, correlation
from math import isfinite

FILE_PATH = r"C:\Users\Hi\OneDrive\Desktop\CAP776\12628918.xlsx"
SHEET_NAME = "Daily Log"
START_DATE = date(2026, 8, 13)
END_DATE = date(2026, 9, 25)
MINUTES_PER_DAY = 1440


print(f"---STUDENT DETAILS---")
print(f"Student name = Palak")
print(f"Registration number = 12628918")
print(f"Roll no. = RD1P2636A18")
print(f"Program = MCA")
print(f"section = D1P2636")
print(f"Activity recording period ={START_DATE} TO {END_DATE} ")



FEELING_SCORES = {"Excellent": 5, "Good": 4, "Neutral": 3, "Low": 2, "Stressed": 1}
SATISFACTION_SCORES = {"Very Satisfied": 5, "Satisfied": 4, "Neutral": 3, "Unsatisfied": 2, "Very Unsatisfied": 1}
ENERGY_SCORES = {"High": 3, "Medium": 2, "Low": 1}
print(f"\n1. Activity data summary")
EXPECTED_DAYS = (END_DATE - START_DATE).days + 1
print(f"Expected recording days: {EXPECTED_DAYS} ({START_DATE} to {END_DATE})")

def as_date(value):
    """Convert an Excel date/datetime to date; return None for other values."""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return None

def number_or_error(value, label):
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ValueError(f"{label} must be a number")
    if not isfinite(value) or value < 0:
        raise ValueError(f"{label} must be a finite, non-negative number")
    return float(value)

def load_activity_rows(file_path):
    workbook = load_workbook(file_path, data_only=True, read_only=True)
    if SHEET_NAME not in workbook.sheetnames:
        raise KeyError(f"Sheet '{SHEET_NAME}' not found. Available sheets: {workbook.sheetnames}")
    sheet = workbook[SHEET_NAME]
    headers = [cell.value for cell in sheet[5]]
    expected_headers = ["Date", "Sleep (min)", "Fitness (min)", "Study (min)",
                        "Coding (min)", "Class (min)", "Classes Attended",
                        "Other Activities (min)", "Total Tracked (min)",
                        "Free/Unaccounted (min)", "Day's Feeling",
                        "Satisfaction Level", "Energy Level", "Notes"]
    if headers[:len(expected_headers)] != expected_headers:
        raise ValueError("Unexpected column headings in row 5. Check that this is the supplied Daily Log template.")

    candidates = []
    invalid_records = []
    for excel_row, row in enumerate(sheet.iter_rows(min_row=6, values_only=True), start=6):
        if not row or all(value is None for value in row):
            continue
        logged_date = as_date(row[0])

        if logged_date is None or not (START_DATE <= logged_date <= END_DATE):
            continue
        candidates.append((excel_row, logged_date, row))

    date_counts = Counter(item[1] for item in candidates)
    valid_rows = []
    for excel_row, logged_date, row in candidates:
        errors = []
        if date_counts[logged_date] > 1:
            errors.append("duplicate date; all rows for this date are excluded")
        try:
            sleep = number_or_error(row[1], "Sleep")
            fitness = number_or_error(row[2], "Fitness")
            study = number_or_error(row[3], "Study")
            coding = number_or_error(row[4], "Coding")
            class_time = number_or_error(row[5], "Class")
            other = number_or_error(row[7], "Other Activities")
        except ValueError as error:
            errors.append(str(error))
            sleep = fitness = study = coding = class_time = other = None
        feeling = str(row[10]).strip() if row[10] is not None else ""
        satisfaction = str(row[11]).strip() if row[11] is not None else ""
        energy = str(row[12]).strip() if row[12] is not None else ""
        if feeling not in FEELING_SCORES:
            errors.append(f"unrecognized/missing feeling: {feeling!r}")
        if satisfaction not in SATISFACTION_SCORES:
            errors.append(f"unrecognized/missing satisfaction: {satisfaction!r}")
        if energy not in ENERGY_SCORES:
            errors.append(f"unrecognized/missing energy: {energy!r}")

        if not errors:
            total_tracked = sleep + fitness + study + coding + class_time + other
            if total_tracked > MINUTES_PER_DAY:
                errors.append(f"tracked time {total_tracked:g} exceeds 1,440 minutes")
        if errors:
            invalid_records.append((excel_row, logged_date, errors))
            continue

        valid_rows.append({
            "date": logged_date, "sleep": sleep, "fitness": fitness,
            "study": study, "coding": coding, "class": class_time,
            "other": other, "total_tracked": total_tracked,
            "free_time": MINUTES_PER_DAY - total_tracked,
            "feeling": feeling, "satisfaction": satisfaction, "energy": energy,
            "feeling_score": FEELING_SCORES[feeling],
            "satisfaction_score": SATISFACTION_SCORES[satisfaction],
            "energy_score": ENERGY_SCORES[energy],
        })
    workbook.close()
    return valid_rows, invalid_records

try:
    daily_records, invalid_records = load_activity_rows(FILE_PATH)
except FileNotFoundError:
    raise FileNotFoundError(f"Could not find workbook: {FILE_PATH}. Update FILE_PATH and run this cell again.")

if not daily_records:
    raise ValueError("No valid activity rows found in the selected date range. Check the workbook and validation messages.")

recorded_dates = {record["date"] for record in daily_records}
missing_dates = [START_DATE.fromordinal(day) for day in range(START_DATE.toordinal(), END_DATE.toordinal() + 1)
                 if START_DATE.fromordinal(day) not in recorded_dates]
print(f"Valid recorded days: {len(daily_records)}")
print(f"Missing days: {len(missing_dates)}")
print(f"Invalid/excluded in-period rows: {len(invalid_records)}")
if invalid_records:
    print("\nInvalid rows:")
    for row_number, logged_date, errors in invalid_records:
        print(f"  Excel row {row_number} ({logged_date}): {'; '.join(errors)}")
if missing_dates:
    print("\nDates with no valid record:", ", ".join(day.isoformat() for day in missing_dates))

def average_field(records, field):
    return mean(record[field] for record in records)

def calculate_indices(records, expected_days):
    if not records:
        raise ValueError("Cannot calculate averages without valid recorded days.")
    tpi = average_field(records, "coding")
    aai = mean(record["study"] + record["class"] for record in records)
    phai = average_field(records, "fitness")
    sri = average_field(records, "sleep")
    abi = average_field(records, "free_time")
    tui = average_field(records, "total_tracked")
    ei = mean((record["feeling_score"] + record["satisfaction_score"] + record["energy_score"]) / 3
              for record in records)
    dci = len(records) / expected_days * 100
    pai = (0.15 * tpi + 0.20 * aai + 0.15 * phai + 0.20 * sri
           + 0.15 * tui + 0.10 * ei + 0.05 * dci)
    return {"TPI": tpi, "AAI": aai, "PhAI": phai, "SRI": sri,
            "ABI": abi, "TUI": tui, "EI": ei, "DCI": dci, "PAI": pai}

indices = calculate_indices(daily_records, EXPECTED_DAYS)
print("Average daily activity (minutes/day)")
for label, key in [("Sleep", "sleep"), ("Fitness", "fitness"),
                   ("Study", "study"), ("Coding", "coding"),
                   ("Class", "class"), ("Other Activities", "other"),
                   ("Free/Unaccounted", "free_time")]:
    print(f"  {label}: {average_field(daily_records, key):.2f}")



print("\n2. Index values")
for name in ("TPI", "AAI", "PhAI", "SRI", "ABI", "TUI"):
    print(f"  {name}: {indices[name]:.2f} min/day")
print(f"  EI: {indices['EI']:.2f} / 5 (per report formula)")
print(f"  DCI: {indices['DCI']:.2f}%")
print(f"  PAI: {indices['PAI']:.2f} (report's weighted formula)")

def safe_correlation(records, x_key, y_key):
    x_values = [record[x_key] for record in records]
    y_values = [record[y_key] for record in records]
    if len(records) < 2:
        return None
    if len(set(x_values)) < 2 or len(set(y_values)) < 2:
        return None
    return correlation(x_values, y_values)

relationships = [
    ("Sleep ↔ Energy", "sleep", "energy_score"),
    ("Study ↔ Satisfaction", "study", "satisfaction_score"),
    ("Coding ↔ Energy", "coding", "energy_score"),
]
print("\n3. key findings from my data")
for label, x_key, y_key in relationships:
    value = safe_correlation(daily_records, x_key, y_key)
    if value is None:
        print(f"  {label}: not available (need at least two days and variation in both values)")
    else:
        direction = "positive" if value > 0 else "negative" if value < 0 else "no linear association"
        print(f"  {label}: r = {value:.3f} ({direction} association)")


def describe_relationship(value):
    if value is None:
        return "could not be calculated"

    strength = (
        "negligible" if abs(value) < 0.10 else
        "weak" if abs(value) < 0.30 else
        "moderate" if abs(value) < 0.50 else
        "strong"
    )
    direction = "positive" if value > 0 else "negative" if value < 0 else ""
    return f"a {strength} {direction} relationship (r = {value:.2f})"


def write_findings(records, expected_days):
    recorded_days = len(records)

    average_sleep = average_field(records, "sleep")
    average_fitness = average_field(records, "fitness")
    average_study = average_field(records, "study")
    average_coding = average_field(records, "coding")
    average_class = average_field(records, "class")

    positive_feelings = sum(
        record["feeling"] in ("Good", "Excellent") for record in records
    )
    satisfied_days = sum(
        record["satisfaction"] in ("Satisfied", "Very Satisfied")
        for record in records
    )

    sleep_energy = safe_correlation(records, "sleep", "energy_score")
    study_satisfaction = safe_correlation(
        records, "study", "satisfaction_score"
    )
    coding_energy = safe_correlation(records, "coding", "energy_score")

    findings = (
        f"I recorded my activities on {recorded_days} of {expected_days} "
        f"expected days, showing consistent tracking. My routine focused mainly "
        f"on academics and technical work: I averaged {average_study:.0f} "
        f"minutes of study, {average_class:.0f} minutes of class, and "
        f"{average_coding:.0f} minutes of coding daily. Fitness averaged "
        f"{average_fitness:.0f} minutes per day, while sleep averaged "
        f"{average_sleep:.0f} minutes. This suggests physical activity was a "
        f"small part of my routine, and sleep may need more attention. I "
        f"reported Good or Excellent feelings on {positive_feelings} days and "
        f"Satisfied or Very Satisfied feelings on {satisfied_days} days. In my "
        f"records, sleep and energy had {describe_relationship(sleep_energy)}; "
        f"study and satisfaction had "
        f"{describe_relationship(study_satisfaction)}; coding and energy had "
        f"{describe_relationship(coding_energy)}. These patterns describe my "
        f"data, but they do not prove that an activity caused a particular "
        f"feeling."
    )

    return findings

print("\n4. Finding myself")
print(write_findings(daily_records, EXPECTED_DAYS))

average_sleep_hours = average_field(daily_records, "sleep") / 60
average_fitness = average_field(daily_records, "fitness")
days_without_fitness = sum(
    record["fitness"] == 0 for record in daily_records
)
average_study = average_field(daily_records, "study")
average_coding = average_field(daily_records, "coding")

improvement_plan = (
    f"Based on my {len(daily_records)} recorded days, I would like to improve "
    f"my sleep and physical activity. I averaged {average_sleep_hours:.1f} "
    f"hours of sleep per day, so I will try to keep a consistent bedtime and "
    f"gradually work toward at least seven hours. My fitness averaged "
    f"{average_fitness:.0f} minutes daily, and I recorded no fitness on "
    f"{days_without_fitness} days. I plan to start with a 20-minute walk or "
    f"workout five days a week and build the habit gradually. I also averaged "
    f"{average_study:.0f} minutes of study and {average_coding:.0f} minutes "
    f"of coding per day. I will take short breaks between focused sessions "
    f"and avoid reducing sleep to finish work. I will review my activity log "
    f"weekly to see whether these changes are helping."
)

print("\n5. Areas and Ways to Improve My Lifestyle")
print(improvement_plan)


student_name = ws["B2"].value
report_date = date.today().strftime("%d %B %Y")
print("\n6. Student Declaration")
print(
    "I confirm that the values entered in this report have been generated "
    "from my submitted activity data using my Python program and that the "
    "report represents my own analysis."
)
print(f"\nStudent Name: {student_name}")
print("Signature: ____________________")
print(f"Date: {report_date}")

