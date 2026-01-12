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
    Extract index and grade pairs from a PDF file using robust column detection.
    Supports multi-column layouts where multiple Index/Grade pairs exist in a single row.
    Returns: list of tuples [(index, grade), ...]
    """
    print(f"  - Processing '{pdf_path}'...")
    
    # Use stream=True which works better for the observed PDF formats
    grade_tables = tabula.read_pdf(pdf_path, pages="all", stream=True, pandas_options={'header': None})
    index_grade_pairs = []
    
    for tbl in grade_tables:
        if tbl.empty:
            continue
            
        # Clean the table: drop all-NaN rows/cols if any
        df = tbl.dropna(how='all').dropna(axis=1, how='all')
        
        # Reset column names to integer index to avoid confusion
        df.columns = range(df.shape[1])
        
        # Heuristics to find pairs of Index and Grade columns
        # We need to find ALL pairs, not just one.
        
        # 1. Start scanning from where valid data seems to begin (skip headers)
        start_row = 0
        for row_idx in range(min(5, len(df))):
            row_vals = [str(x).strip().lower() for x in df.iloc[row_idx]]
            if any("index" in x for x in row_vals) or any("grade" in x for x in row_vals):
                start_row = row_idx + 1
                break
        
        valid_rows_for_analysis = df.iloc[start_row:].head(20)
        
        # Identify columns by type: 'index', 'grade', or 'unknown'
        col_types = {}
        
        for col_idx in df.columns:
            col_data = valid_rows_for_analysis[col_idx].astype(str).tolist()
            
            grade_matches = 0
            index_matches = 0
            
            for cell in col_data:
                cell = cell.strip()
                if not cell or cell.lower() == "nan": continue
                
                # Check for Index pattern anywhere in string
                if re.search(r'\d{6}[A-Z]?', cell):
                    index_matches += 1
                
                # Check for Grade
                if cell in GRADES or cell in ["F", "I-we", "I-ca", "ab"]:
                     grade_matches += 1
            
            # Determine type based on dominance
            if index_matches > 0 and index_matches >= len(col_data) * 0.3:
                col_types[col_idx] = 'index'
            elif grade_matches > 0 and grade_matches >= len(col_data) * 0.3:
                col_types[col_idx] = 'grade'
            else:
                col_types[col_idx] = 'unknown'

        # Pairing strategy:
        # Sort columns left-to-right. For each 'index' column, pair it with the 
        # nearest 'grade' column to its right that hasn't been used.
        
        used_cols = set()
        sorted_cols = sorted(col_types.keys())
        
        for i, idx_col in enumerate(sorted_cols):
            if col_types[idx_col] == 'index' and idx_col not in used_cols:
                # Look for nearest grade col to the right
                grade_col = -1
                
                for j in range(i + 1, len(sorted_cols)):
                    candidate_col = sorted_cols[j]
                    if col_types[candidate_col] == 'grade' and candidate_col not in used_cols:
                        grade_col = candidate_col
                        break
                
                if grade_col != -1:
                    # Found a pair
                    used_cols.add(idx_col)
                    used_cols.add(grade_col)
                    
                    # Extract from this pair
                    subset = df.iloc[start_row:, [idx_col, grade_col]].values
                    
                    for row_dat in subset:
                        idx_raw = str(row_dat[0]).strip()
                        grade = str(row_dat[1]).strip()
                        
                        if not idx_raw: continue

                        try:
                            # Extract numeric part from anywhere in the string
                            numeric_part_match = re.search(r'(\d{6})', idx_raw)
                            if numeric_part_match:
                                idx_int = int(numeric_part_match.group(1))
                                
                                if idx_int in valid_indices:
                                    if grade and grade.lower() != "nan":
                                        index_grade_pairs.append((idx_int, grade))
                        except ValueError:
                            continue

    return index_grade_pairs

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
                        #print(f"Index with corrections:{idx_str}")
                        idx = int(idx_str)
                        print(idx)
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
    
    return round(weighted_sum / total_credits, 3)

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
    
    return round(max_sum / total_credits, 3)

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
# CGPA CALCULATION FUNCTIONS
# ============================================================================

def process_semester_for_cgpa(semester_config_path, student_indices, students_db, corrections):
    """
    Process a single semester for CGPA calculation.
    Returns: (semester_name, semester_results_per_student)
    semester_results_per_student = {student_idx: {'sgpa': float, 'credits': int, 'weighted_points': float}}
    """
    semester_config = load_semester_config(semester_config_path)
    semester_name = semester_config.get("semester_name", "Unknown")
    
    print(f"\n# Processing {semester_name}...")
    
    course_info = {"index_range": (0, 0), "students": students_db} # Range not strictly needed here
    
    # Load results
    results, available_modules, module_stats = load_all_module_results(
        semester_config, course_info, corrections
    )
    
    processed_data = {}
    
    for idx in student_indices:
        student_results = results.get(idx, {})
        
        # Calculate SGPA variables
        total_credits = 0
        weighted_sum = 0
        
        for module_code, grade in student_results.items():
            if module_code in module_stats and grade in GRADES:
                credits = module_stats[module_code]["credits"]
                gpa_value = GRADES[grade]["gpa_4_0"] # Using 4.0 scale for calculation
                
                weighted_sum += credits * gpa_value
                total_credits += credits
        
        sgpa = 0.0
        if total_credits > 0:
            sgpa = round(weighted_sum / total_credits, 3)
            
        processed_data[idx] = {
            "sgpa": sgpa,
            "credits": total_credits,
            "weighted_points": weighted_sum
        }
        
    return semester_name, processed_data

def calculate_cgpa_flow(students_db, corrections):
    """Execute CGPA calculation flow"""
    print("\n" + "=" * 40)
    print("      CGPA CALCULATION MODE")
    print("=" * 40)
    
    config_files = get_semester_config_files()
    if not config_files:
        print("! Error: No semester configurations found.")
        return

    student_indices = list(students_db.keys())
    
    # Data structure: {student_idx: {'semesters': {sem_name: sgpa}, 'total_credits': 0, 'total_points': 0}}
    cgpa_data = {idx: {'semesters': {}, 'total_credits': 0, 'total_points': 0} for idx in student_indices}
    semester_names = []
    
    # Process each semester
    for config_file in config_files:
        sem_name, sem_results = process_semester_for_cgpa(config_file, student_indices, students_db, corrections)
        semester_names.append(sem_name)
        
        for idx, data in sem_results.items():
            if idx in cgpa_data:
                cgpa_data[idx]['semesters'][sem_name] = data['sgpa']
                cgpa_data[idx]['total_credits'] += data['credits']
                cgpa_data[idx]['total_points'] += data['weighted_points']
    
    # Calculate Final CGPA
    final_results = []
    
    print("\n# Calculating Final CGPA...")
    for idx, data in cgpa_data.items():
        total_credits = data['total_credits']
        total_points = data['total_points']
        
        cgpa = 0.0
        if total_credits > 0:
            cgpa = round(total_points / total_credits, 3)
            
        final_results.append({
            "idx": idx,
            "name": students_db.get(idx, {}).get("name", "Unknown"),
            "semesters": data['semesters'],
            "cgpa": cgpa
        })
        
    # Sort by CGPA descending
    final_results.sort(key=lambda x: x['cgpa'], reverse=True)
    
    # Assign Ranks
    print("\n# Exporting CGPA Results...")
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    filename = OUTPUT_FOLDER / "CGPA_Results.xlsx"
    workbook = xlsxwriter.Workbook(filename)
    ws = workbook.add_worksheet("CGPA")
    
    # Headers
    headers = ["Rank", "Index", "Name"] + semester_names + ["CGPA"]
    for i, h in enumerate(headers):
        ws.write(0, i, h)
        
    # Data
    for rank, student in enumerate(final_results, start=1):
        row = rank
        ws.write(row, 0, rank)
        ws.write(row, 1, student['idx'])
        ws.write(row, 2, student['name'])
        
        col = 3
        for sem in semester_names:
            ws.write(row, col, student['semesters'].get(sem, 0.0))
            col += 1
            
        ws.write(row, col, student['cgpa'])
        
    workbook.close()
    print(f"  [OK] Created '{filename}'")

def calculate_sgpa_flow(students_db, corrections):
    """Execute standard SGPA calculation flow"""
    semester_config_path = select_semester_config()
    if not semester_config_path:
        return
        
    semester_config = load_semester_config(semester_config_path)
    
    # Calculate index range from students_db
    student_indices = list(students_db.keys())
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
    
    if not students_db:
        print("! Error: No valid students found in database.")
        return

    while True:
        print("\nSelect Mode:")
        print("  1. Calculate SGPA (Single Semester)")
        print("  2. Calculate CGPA (All Semesters)")
        print("  q. Quit")
        
        choice = input("\nEnter choice (1/2/q): ").strip().lower()
        
        if choice == '1':
            calculate_sgpa_flow(students_db, corrections)
        elif choice == '2':
            calculate_cgpa_flow(students_db, corrections)
        elif choice == 'q':
            print("Exiting...")
            break
        else:
            print("Invalid choice.")

if __name__ == "__main__":
    main()