"""
Scalable University GPA Analysis System
========================================
Modular system for processing semester results from PDFs
"""

import glob
import json
from pathlib import Path
import tabula
import os.path
import re
import xlsxwriter

# ============================================================================
# CONFIGURATION
# ============================================================================

BASE_DIR = Path(__file__).resolve().parent

STUDENTS_FILE = BASE_DIR/"data"/"student_details.json"
GRADES_FILE = BASE_DIR/"config"/"grades.json"
CORRECTIONS_FILE = BASE_DIR/"config"/"corrections.json"
SEMESTER_CONFIG_DIR = BASE_DIR/"config"/"semesters"
RESULTS_FOLDER = BASE_DIR/"data"/"results"  # Folder containing PDF files
OUTPUT_FOLDER = BASE_DIR/"output/"

GRADES = {}



# ============================================================================
# DATA LOADING FUNCTIONS
# ============================================================================

def load_grades(filepath):
    """Load grades from JSON file"""
    print(f"# Loading grades from '{filepath}'...")
    with open(filepath, 'r') as f:
        return json.load(f)

def load_corrections(filepath):
    """Load grade corrections from JSON file"""
    if not os.path.exists(filepath):
        print(f"# No corrections file found at '{filepath}'. Skipping.")
        return {}
        
    print(f"# Loading corrections from '{filepath}'...")
    with open(filepath, 'r') as f:
        return json.load(f)

def load_students(filepath):
    """Load student details from JSON file and index by int(idx)"""
    print(f"# Loading student data from '{filepath}'...")
    with open(filepath, 'r') as f:
        raw_data = json.load(f)
        
    # Re-index by integer ID for matching with PDF results
    processed_db = {}
    for key, student in raw_data.items():
        try:
            # Assuming 'idx' field exists and is the numeric part
            idx_int = int(student.get("idx", 0))
            if idx_int > 0:
                processed_db[idx_int] = student
        except ValueError:
            continue
            
    return processed_db

def get_semester_config_files():
    """Get list of semester config files"""
    pattern = str(SEMESTER_CONFIG_DIR / "*.json")
    return sorted(glob.glob(pattern))

def select_semester_config():
    """Prompt user to select a semester config file"""
    files = get_semester_config_files()
    
    if not files:
        print("! Error: No semester configuration files found in 'config/semesters/'")
        return None
        
    if len(files) == 1:
        print(f"# Auto-selecting only available config: {Path(files[0]).name}")
        return files[0]
        
    print("\nAvailable Semesters:")
    for i, f in enumerate(files):
        print(f"  {i+1}. {Path(f).name}")
        
    while True:
        try:
            choice = input("\nSelect semester (1-{}): ".format(len(files)))
            idx = int(choice) - 1
            if 0 <= idx < len(files):
                return files[idx]
            print("! Invalid selection. Please try again.")
        except ValueError:
            print("! Invalid input. Please enter a number.")

def load_semester_config(filepath):
    """Load semester configuration and normalize structure"""
    print(f"# Loading semester config from '{filepath}'...")
    with open(filepath, 'r') as f:
        config = json.load(f)
        
    # Normalize 'sem_name' to 'semester_name'
    if "sem_name" in config:
        config["semester_name"] = config["sem_name"]
        
    # Normalize 'courses' list to 'modules' dict if necessary
    if "courses" in config and "modules" not in config:
        config["modules"] = {
            m["code"]: m for m in config["courses"]
        }
        
    return config

# ============================================================================
# PDF EXTRACTION FUNCTIONS
# ============================================================================

def extract_results_from_pdf(pdf_path, valid_indices):
    """
    Extract index and grade pairs from a PDF file
    Returns: list of tuples [(index, grade), ...]
    """
    print(f"  - Processing '{pdf_path}'...")
    
    grade_tables = tabula.read_pdf(pdf_path, pages="all", pandas_options={'header': None})
    index_grade_pairs = []
    
    for tbl in grade_tables:
        if not tbl.empty and len(tbl.columns) > 1:
            try:
                # Check if tbl keys are integers (header=None produces int cols)
                if 0 in tbl.columns and 1 in tbl.columns:
                    # Convert to list ignoring the first header row if it was scraped as data
                    if pdf_path.name in ["EN1020.pdf", "EN1971.pdf"]:
                        idxs = tbl[1].tolist()
                        grades = tbl[6].tolist()
                    else:
                        idxs = tbl[0].tolist()
                        grades = tbl[1].tolist()
                    pairs = list(zip(idxs, grades))
                    index_grade_pairs.extend(pairs)
            except Exception as e:
                print(f"DEBUG: Error extracting from table: {e}")
    
    # Filter valid entries against student database
    valid_results = []
    for idx_str, grade in index_grade_pairs:
        if str(idx_str) != "nan" and str(idx_str).strip() != "Index No.":
            try:
                # Extract digits only (handle '230012U', '230012/U', etc.)
                clean_idx_str = re.sub(r'\D', '', str(idx_str))
                if not clean_idx_str:
                    continue
                    
                idx = int(clean_idx_str)
                if idx in valid_indices:
                    valid_results.append((idx, grade))
            except ValueError:
                continue
    
    return valid_results

def load_all_module_results(semester_config, course_info, corrections=None):
    """
    Load results for all modules in the semester
    Returns: (results_dict, available_modules, module_stats)
    """
    results = {}
    available_modules = []
    module_stats = {}
    
    valid_indices = set(course_info["students"].keys())
    
    print("\n# Extracting results from PDFs...")
    
    semester_name = semester_config.get("semester_name", "")
    
    for module_code, module_info in semester_config["modules"].items():
        # Look for PDF in results/semester_name/module_code.pdf
        # If semester_name is empty or not found, try root results folder or handle as needed.
        # Assuming structure data/results/sem1/CODE.pdf
        pdf_path = RESULTS_FOLDER / semester_name / f"{module_code}.pdf"
        
        if os.path.isfile(pdf_path):
            available_modules.append(module_code)
            module_stats[module_code] = {
                "credits": module_info["credits"],
                "grade_counts": {}
            }
            
            # Extract results from PDF
            module_results = extract_results_from_pdf(pdf_path, valid_indices)
            
            for idx, grade in module_results:
                if idx not in results:
                    results[idx] = {}
                results[idx][module_code] = grade
                
                # Update grade statistics
                module_stats[module_code]["grade_counts"][grade] = \
                    module_stats[module_code]["grade_counts"].get(grade, 0) + 1
        else:
            print(f"  ! Warning: '{pdf_path}' not found. Skipping module {module_code}.")
            
    # Apply corrections if available
    if corrections:
        print("\n# Applying manual corrections...")
        for module_code, module_corrections in corrections.items():
            if module_code in module_stats:
                for idx_str, new_grade in module_corrections.items():
                    try:
                        idx = int(idx_str)
                        if idx in valid_indices:
                            # Update result
                            if idx not in results:
                                results[idx] = {}
                            
                            old_grade = results[idx].get(module_code, "N/A")
                            results[idx][module_code] = new_grade
                            print(f"  - Corrected {idx} in {module_code}: {old_grade} -> {new_grade}")
                            
                            # Adjust stats (simple approach: decrement old, increment new)
                            # Note: This might be slightly inaccurate if we didn't count the old grade originally
                            # but ensures the new grade is counted.
                            if old_grade in module_stats[module_code]["grade_counts"]:
                                module_stats[module_code]["grade_counts"][old_grade] -= 1
                            
                            module_stats[module_code]["grade_counts"][new_grade] = \
                                module_stats[module_code]["grade_counts"].get(new_grade, 0) + 1
                    except ValueError:
                        continue
    
    return results, available_modules, module_stats

# ============================================================================
# GPA CALCULATION FUNCTIONS
# ============================================================================

def calculate_gpa(student_results, module_stats, scale="4_0"):
    """
    Calculate GPA for a student based on available modules
    scale: "4_0" or "4_2"
    """
    total_credits = 0
    weighted_sum = 0
    
    for module_code, grade in student_results.items():
        if module_code in module_stats and grade in GRADES:
            credits = module_stats[module_code]["credits"]
            gpa_value = GRADES[grade][f"gpa_{scale}"]
            
            weighted_sum += credits * gpa_value
            total_credits += credits
    
    if total_credits == 0:
        return 0.0
    
    return round(weighted_sum / total_credits, 2)

def calculate_max_possible_gpa(student_results, module_stats, semester_config):
    """
    Calculate maximum possible GPA if student gets A in all remaining modules
    """
    # Current weighted sum (4.0 scale)
    current_sum = 0
    current_credits = 0
    
    for module_code, grade in student_results.items():
        if module_code in module_stats and grade in GRADES:
            credits = module_stats[module_code]["credits"]
            gpa_value = GRADES[grade]["gpa_4_0"]
            current_sum += credits * gpa_value
            current_credits += credits
    
    # Total credits for all modules in semester
    total_credits = sum(m["credits"] for m in semester_config["modules"].values())
    
    # Maximum possible sum (assuming A in remaining modules)
    max_sum = current_sum + (total_credits - current_credits) * 4.0
    
    if total_credits == 0:
        return 0.0
    
    return round(max_sum / total_credits, 2)

# ============================================================================
# RANKING AND SORTING
# ============================================================================

def rank_students(results, module_stats, semester_config, available_modules):
    """
    Calculate GPAs and rank students
    Returns: sorted list of (idx, student_data) tuples
    """
    print("\n# Calculating GPAs and rankings...")
    
    student_data = {}
    
    for idx, module_grades in results.items():
        gpa_4_0 = calculate_gpa(module_grades, module_stats, "4_0")
        gpa_4_2 = calculate_gpa(module_grades, module_stats, "4_2")
        max_gpa = calculate_max_possible_gpa(module_grades, module_stats, semester_config)
        
        student_data[idx] = {
            "modules": module_grades,
            "gpa_4_0": gpa_4_0,
            "gpa_4_2": gpa_4_2,
            "max_gpa": max_gpa,
            "module_count": len(module_grades)
        }
    
    # Sort by: GPA (4.0), then GPA (4.2), then individual module GPAs, then index
    def sort_key(item):
        idx, data = item
        gpa_4_0 = data["gpa_4_0"]
        gpa_4_2 = data["gpa_4_2"]
        
        # Get GPA values for each available module (for tie-breaking)
        module_gpas = []
        for module in available_modules:
            if module in data["modules"] and data["modules"][module] in GRADES:
                module_gpas.append(GRADES[data["modules"][module]]["gpa_4_2"])
            else:
                module_gpas.append(0.0)
        
        return (gpa_4_0, gpa_4_2, *module_gpas, -idx)
    
    sorted_students = sorted(student_data.items(), key=sort_key, reverse=True)
    
    # Assign ranks
    prev_gpa_4_0 = None
    prev_gpa_4_2 = None
    rank = 1
    rank_4_2 = 1
    rank_gap = 0
    rank_4_2_gap = 0
    
    for i, (idx, data) in enumerate(sorted_students):
        # Rank by 4.0 GPA
        if prev_gpa_4_0 == data["gpa_4_0"]:
            data["rank"] = rank
            rank_gap += 1
        else:
            rank += rank_gap
            data["rank"] = rank
            rank_gap = 1
            prev_gpa_4_0 = data["gpa_4_0"]
        
        # Rank by 4.2 GPA (tie-breaker)
        if prev_gpa_4_2 == data["gpa_4_2"]:
            data["rank_4_2"] = rank_4_2
            rank_4_2_gap += 1
        else:
            rank_4_2 += rank_4_2_gap
            data["rank_4_2"] = rank_4_2
            rank_4_2_gap = 1
            prev_gpa_4_2 = data["gpa_4_2"]
    
    return sorted_students

# ============================================================================
# EXCEL EXPORT FUNCTIONS
# ============================================================================

def export_to_excel(sorted_students, students_db, available_modules, module_stats, 
                    semester_config, course_name):
    """Export results to Excel files"""
    
    print("\n# Exporting to Excel files...")
    
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    
    semester_name = semester_config["semester_name"]
    n_modules = len(available_modules)
    total_modules = len(semester_config["modules"])
    
    # ========== File 1: Basic Results ==========
    filename1 = OUTPUT_FOLDER / f"Results - {semester_name}.xlsx"
    workbook1 = xlsxwriter.Workbook(filename1)
    ws1 = workbook1.add_worksheet("Results")
    
    # Headers
    ws1.write(0, 0, "Rank")
    ws1.write(0, 1, "Index")
    
    for i, module in enumerate(available_modules):
        ws1.write(0, i + 2, module)
    
    col = n_modules + 2
    if n_modules != total_modules:
        ws1.write(0, col, "Current SGPA")
        ws1.write(0, col + 1, "Max Possible SGPA")
    else:
        ws1.write(0, col, "SGPA")
    
    # Student data
    for row, (idx, data) in enumerate(sorted_students, start=1):
        student_info = students_db.get(idx, {})
        
        ws1.write(row, 0, data["rank"])
        ws1.write(row, 1, student_info.get("idx", idx))
        
        for col, module in enumerate(available_modules, start=2):
            grade = data["modules"].get(module, "-")
            ws1.write(row, col, grade)
        
        col = n_modules + 2
        ws1.write(row, col, data["gpa_4_0"])
        if n_modules != total_modules:
            ws1.write(row, col + 1, data["max_gpa"])
    
    # Grade statistics
    col_offset = n_modules + 6 if n_modules == total_modules else n_modules + 5
    
    for i, module in enumerate(available_modules):
        ws1.write(0, col_offset + i, module)
    
    row = 1
    for grade in GRADES.keys():
        ws1.write(row, col_offset - 1, grade)
        
        for i, module in enumerate(available_modules):
            count = module_stats[module]["grade_counts"].get(grade, 0)
            total = sum(module_stats[module]["grade_counts"].values())
            percentage = (count / total * 100) if total > 0 else 0
            ws1.write(row, col_offset + i, f"{count}({percentage:.1f}%)")
        
        row += 1
    
    workbook1.close()
    print(f"  [OK] Created '{filename1}'")
    

    
    # ========== File 2: Extended Results ==========
    filename2 = OUTPUT_FOLDER / f"Results - {semester_name} (Extended).xlsx"
    workbook2 = xlsxwriter.Workbook(filename2)
    ws2 = workbook2.add_worksheet("Results")
    
    # Headers
    ws2.write(0, 0, "Rank")
    ws2.write(0, 1, "Index")
    ws2.write(0, 2, "Name")
    
    for i, module in enumerate(available_modules):
        ws2.write(0, i + 3, module)
    
    col = n_modules + 3
    if n_modules != total_modules:
        ws2.write(0, col, "Current SGPA")
        ws2.write(0, col + 1, "Max Possible SGPA")
        ws2.write(0, col + 2, "Rank (4.2 scale)")
    else:
        ws2.write(0, col, "SGPA")
        ws2.write(0, col + 1, "Rank (4.2 scale)")
    
    # Student data
    for row, (idx, data) in enumerate(sorted_students, start=1):
        student_info = students_db.get(idx, {})
        
        ws2.write(row, 0, data["rank"])
        ws2.write(row, 1, student_info.get("idx", idx))
        ws2.write(row, 2, student_info.get("name", "Unknown"))
        
        for col, module in enumerate(available_modules, start=3):
            grade = data["modules"].get(module, "-")
            ws2.write(row, col, grade)
        
        col = n_modules + 3
        ws2.write(row, col, data["gpa_4_0"])
        if n_modules != total_modules:
            ws2.write(row, col + 1, data["max_gpa"])
            ws2.write(row, col + 2, data["rank_4_2"])
        else:
            ws2.write(row, col + 1, data["rank_4_2"])
    
    # Grade statistics
    col_offset = n_modules + 8 if n_modules == total_modules else n_modules + 7
    
    for i, module in enumerate(available_modules):
        ws2.write(0, col_offset + i, module)
    
    row = 1
    for grade in GRADES.keys():
        ws2.write(row, col_offset - 1, grade)
        
        for i, module in enumerate(available_modules):
            count = module_stats[module]["grade_counts"].get(grade, 0)
            total = sum(module_stats[module]["grade_counts"].values())
            percentage = (count / total * 100) if total > 0 else 0
            ws2.write(row, col_offset + i, f"{count}({percentage:.1f}%)")
        
        row += 1
    
    workbook2.close()
    print(f"  [OK] Created '{filename2}'")

# ============================================================================
# MAIN EXECUTION
# ============================================================================

def main():
    print("=" * 70)
    print("University GPA Analysis System (v5.00)")
    print("=" * 70)
    
    # Load configuration

    global GRADES
    GRADES = load_grades(GRADES_FILE)
    
    corrections = load_corrections(CORRECTIONS_FILE)
    
    students_db = load_students(STUDENTS_FILE)
    
    semester_config_path = select_semester_config()
    if not semester_config_path:
        return
        
    semester_config = load_semester_config(semester_config_path)
    
    # Calculate index range from students_db
    student_indices = list(students_db.keys())
    if not student_indices:
        print("! Error: No valid students found in database.")
        return
        
    min_idx = min(student_indices)
    max_idx = max(student_indices)
    course_info = {"index_range": (min_idx, max_idx), "students": students_db}
    
    semester_name = semester_config.get("semester_name", "Unknown Semester")
    print(f"\nSemester: {semester_name}")
    print(f"Index Range: {min_idx} - {max_idx}")
    print(f"Total Modules: {len(semester_config.get('modules', {}))}")
    
    # Load results from PDFs
    results, available_modules, module_stats = load_all_module_results(
        semester_config, course_info, corrections
    )
    
    print(f"\n[OK] Found results for {len(available_modules)} modules")
    print(f"[OK] Total students with results: {len(results)}")
    
    # Calculate GPAs and rank
    sorted_students = rank_students(results, module_stats, semester_config, available_modules)
    
    # Export to Excel
    export_to_excel(
        sorted_students, 
        students_db, 
        available_modules, 
        module_stats, 
        semester_config,
        "Results"
    )
    
    print("\n" + "=" * 70)
    print("[OK] Finished successfully!")
    print("=" * 70)

if __name__ == "__main__":
    main()