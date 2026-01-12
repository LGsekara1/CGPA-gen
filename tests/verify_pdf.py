import sys
import os
import json
from pathlib import Path

# Add parent directory to path to import main
current_dir = Path(__file__).resolve().parent
parent_dir = current_dir.parent
sys.path.append(str(parent_dir))

# Import necessary functions
from main import extract_results_from_pdf, load_grades
import main

# Manually load grades
GRADES_FILE = parent_dir / "config" / "grades.json"
if GRADES_FILE.exists():
    main.GRADES = load_grades(GRADES_FILE)
else:
    main.GRADES = {"A+":{}, "A":{}, "A-":{}, "B+":{}, "B":{}, "B-":{}, "C+":{}, "C":{}, "C-":{}, "D":{}, "I-we":{}, "F":{}}

# Load actual student DB for valid indices
STUDENTS_FILE = parent_dir / "data" / "student_details.json"
with open(STUDENTS_FILE, "r") as f:
    student_data = json.load(f)

valid_indices = set()
for s in student_data.values():
    try:
        valid_indices.add(int(s["idx"]))
    except:
        pass

print(f"Loaded {len(valid_indices)} valid students from DB.")

# Test files
# Test files
files = ["EN1054.pdf", "EN1971.pdf"]
base_dir = parent_dir / "data" / "results" / "sem2"

for f in files:
    path = base_dir / f
    print(f"\nTesting extraction for {f}...")
    try:
        results = extract_results_from_pdf(path, valid_indices)
        print(f"Extracted {len(results)} records.")
        if len(results) > 0:
            print(f"Sample: {results[:3]}")
    except Exception as e:
        print(f"Error: {e}")
