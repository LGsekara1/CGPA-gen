import os
import json
import tabula
import configparser
import xlsxwriter
import jpype
jpype.startJVM(r"C:\\Program Files\\Java\\jdk-24\\bin\server\\jvm.dll")

# ---------------- CONFIG LOADERS ----------------

def load_grades(path):
    cfg = configparser.ConfigParser()
    cfg.optionxform = str
    cfg.read(path)
    return {k: float(v) for k, v in cfg["GRADES"].items()}

def load_semester(path, student_index_list):
    cfg = configparser.ConfigParser()
    cfg.optionxform = str
    cfg.read(path)

    modules = {m: int(c) for m, c in cfg["MODULES"].items()}
    all_modules = list(modules.keys())

    students = {}
    raw_students = dict(cfg["STUDENT_MODULES"])

    # Case 1: ALL = *
    if raw_students.get("ALL") == "*":
        for idx in student_index_list:
            students[idx] = all_modules.copy()

    # Case 2: Explicit per-student mapping
    else:
        for idx, mods in raw_students.items():
            if mods.strip() == "*":
                students[idx] = all_modules.copy()
            else:
                students[idx] = [m.strip() for m in mods.split(",")]

    return cfg["SEMESTER"]["name"], modules, students

# ---------------- PDF PARSER ----------------

def parse_pdfs(result_dir):
    results = {}
    for pdf in os.listdir(result_dir):
        if not pdf.endswith(".pdf"):
            continue

        module = pdf.replace(".pdf", "")
        tables = tabula.read_pdf(
            os.path.join(result_dir, pdf),
            pages="all",
            pandas_options={"header": None}
        )

        for tbl in tables:
            for idx, grade in zip(tbl[0][1:], tbl[1][1:]):
                if str(idx) == "nan":
                    continue
                index = idx[:-1]
                results.setdefault(index, {})[module] = grade
    return results

# ---------------- GPA CALCULATIONS ----------------

def calculate_sgpa(student_modules, student_grades, credits, grade_map):
    total_credits = 0
    quality_points = 0

    for module in student_modules:
        grade = student_grades.get(module, "W")
        gp = grade_map.get(grade, 0)
        cr = credits[module]

        total_credits += cr
        quality_points += cr * gp

    return round(quality_points / total_credits, 3), total_credits

# ---------------- EXCEL EXPORT ----------------

def export_sgpa_excel(path, semester, students, sgpa_data):
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("SGPA")

    headers = ["Index", "SGPA", "Credits"]
    for c, h in enumerate(headers):
        ws.write(0, c, h)

    for r, idx in enumerate(students, start=1):
        ws.write(r, 0, idx)
        ws.write(r, 1, sgpa_data[idx]["sgpa"])
        ws.write(r, 2, sgpa_data[idx]["credits"])

    wb.close()

def export_cgpa_excel(cgpa_tracker):
    wb = xlsxwriter.Workbook("output/cgpa/CGPA.xlsx")
    ws = wb.add_worksheet("CGPA")

    headers = ["Index", "CGPA", "Total Credits"]
    for c, h in enumerate(headers):
        ws.write(0, c, h)

    for r, (idx, data) in enumerate(cgpa_tracker.items(), start=1):
        cgpa = round(data["qp"] / data["cr"], 3)
        ws.write(r, 0, idx)
        ws.write(r, 1, cgpa)
        ws.write(r, 2, data["cr"])

    wb.close()

# ---------------- MAIN PIPELINE ----------------

def main():
    with open("data/student_details.json") as f:
        student_meta = json.load(f)

    student_index_list = list(student_meta.keys())
    grades = load_grades("config/grades.ini")

    cgpa_tracker = {}

    semester_dir = "config/semesters"
    semester_files = sorted(
        f for f in os.listdir(semester_dir) if f.endswith(".ini")
    )

    for sem_file in semester_files:
        sem_path = os.path.join(semester_dir, sem_file)

        sem_name, modules, student_modules = load_semester(
            sem_path, student_index_list
        )

        result_dir = f"data/results/{sem_file[:-4]}"
        pdf_results = parse_pdfs(result_dir)

        sgpa_results = {}

        for idx, mods in student_modules.items():
            sgpa, credits = calculate_sgpa(
                mods,
                pdf_results.get(idx, {}),
                modules,
                grades
            )

            sgpa_results[idx] = {
                "sgpa": sgpa,
                "credits": credits
            }

            # CGPA accumulation
            cgpa_tracker.setdefault(idx, {"qp": 0, "cr": 0})
            cgpa_tracker[idx]["qp"] += sgpa * credits
            cgpa_tracker[idx]["cr"] += credits

        export_sgpa_excel(
            f"output/sgpa/{sem_name}.xlsx",
            sem_name,
            student_modules,
            sgpa_results
        )

    export_cgpa_excel(cgpa_tracker)



if __name__ == "__main__":
    main()
