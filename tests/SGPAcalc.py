import sys
import os
from pathlib import Path
import json

# Add project root to sys.path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.append(str(PROJECT_ROOT))

# Import from main
import main
from main import (
    calculate_gpa, load_grades, load_students, load_semester_config, 
    load_all_module_results, load_corrections, 
    GRADES_FILE, STUDENTS_FILE, CORRECTIONS_FILE, SEMESTER_CONFIG_DIR,
    truncate
)

def main_cli():
    # 1. Initialize environment
    print("Initializing...")
    # Patch global GRADES in main module so functions in main.py can use it
    main.GRADES = load_grades(GRADES_FILE) 
    
    # 2. Get inputs
    if len(sys.argv) >= 3:
        student_idx_input = sys.argv[1]
        semester_input = sys.argv[2]
        print(f"Using CLI args: Index={student_idx_input}, Sem={semester_input}")
    else:
        print("\n--- SGPA Calculator ---")
        student_idx_input = input("Enter Student Index (e.g. 230012): ").strip()
        semester_input = input("Enter Semester (e.g. 1, 2): ").strip()

    try:
        student_idx = int(student_idx_input)
    except ValueError:
        print("Invalid index format. Must be a number.")
        return

    # 3. Load Student Data
    students = load_students(STUDENTS_FILE)
    if student_idx not in students:
        print(f"Student index {student_idx} not found in student_details.json")
        return
    
    student_name = students[student_idx].get("name", "Unknown")
    print(f"Student: {student_name} ({student_idx})")

    # 4. Load Semester Config
    sem_config_file = SEMESTER_CONFIG_DIR / f"sem{semester_input}.json"
    if not sem_config_file.exists():
        print(f"Semester config not found: {sem_config_file}")
        # Try full listing
        print("Available semesters:")
        for f in SEMESTER_CONFIG_DIR.glob("*.json"):
            print(f"  - {f.name}")
        return

    semester_config = load_semester_config(sem_config_file)
    sem_name = semester_config.get("semester_name", f"Semester {semester_input}")

    # 5. Extract Results
    corrections = load_corrections(CORRECTIONS_FILE)
    
    # We filter valid_indices to just this student for safety/speed in validation logic,
    # though PDF extraction still scans rows.
    course_info = {
        "students": {student_idx: students[student_idx]} 
    }
    
    print(f"Loading results for {sem_name}...")
    # Temporarily suppress print statements from main functions if desired, 
    # but keeping them is fine for feedback.
    results, available_modules, module_stats = load_all_module_results(
        semester_config, course_info, corrections
    )

    if student_idx not in results:
        print(f"No results found for student {student_idx} in {sem_name}")
        return

    student_modules = results[student_idx]
    
    # 6. Calculate & Display
    print(f"\nResults for {student_name} in {sem_name}:")
    print("-" * 80)
    print(f"{'Module':<10} | {'Grade':<5} | {'Credits':<7} | {'GPA Value':<9} | {'Weighted Points':<15}")
    print("-" * 80)
    
    total_credits = 0
    total_weighted = 0
    
    for module in sorted(student_modules.keys()):
        grade = student_modules[module]
        if module not in module_stats:
            continue
            
        credits = module_stats[module]["credits"]
        
        gpa_val = 0.0
        if grade in main.GRADES:
             gpa_val = main.GRADES[grade]["gpa_4_0"]
        else:
             # If grade is not in GRADES (e.g. "I-we"), it might have 0 value or be ignored in main GPA calc?
             # main.py calculate_gpa checks `if grade in GRADES`.
             # If not in GRADES, it skips contributing to sum and credits.
             # Wait, main.py lines 324-330:
             # if module_code in module_stats and grade in GRADES:
             #    ...
             # So if not in GRADES, it is IGNORED completely (credits not added).
             # Let's mimic that behavior.
             pass
        
        # Determine if we should count it
        if grade in main.GRADES:
            weighted = credits * gpa_val
            total_credits += credits
            total_weighted += weighted
            print(f"{module:<10} | {grade:<5} | {credits:<7} | {gpa_val:<9} | {weighted:<15.2f}")
        else:
            print(f"{module:<10} | {grade:<5} | {credits:<7} | {'N/A':<9} | {'Ignored':<15}")
    
    print("-" * 80)
    print(f"{'TOTAL':<10} | {'':<5} | {total_credits:<7} | {'':<9} | {total_weighted:<15.2f}")
    print("-" * 80)
    
    print(f"\nTotal Weighted Points: {total_weighted:.3f}")
    print(f"Total Credits:         {total_credits}")
    
    # Use the actual function to calculate/test
    sgpa_via_func = calculate_gpa(student_modules, module_stats, scale="4_0")
    print(f"Final SGPA (via calculate_gpa): {sgpa_via_func}")

if __name__ == "__main__":
    main_cli()