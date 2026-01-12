## UOM ENTC GPA (SGPA & CGPA) generator
- Currently CGPA-gen/output contains sem1, sem2, sem3 and CGPA results spreadsheet.
- To get the output run the following and enter semester to calculate
  ```bash
  python -m main
  Select Mode:
  1. Calculate SGPA (Single Semester)
  2. Calculate CGPA (All Semesters)
  q. Quit
  ```
- For SGPA calculation input 1 and corresponding semester number,
  Available Semesters:
  ```bash
  1. sem1.json
  2. sem2.json
  3. sem3.json
  Select semester (1-3):
  ```
- For CGPA calculation input 2, sitback and await results.

  
- The excel files would be generated in data/output
- Current generated outputs sem1, sem2 and sem 3, CGPA file.
- The extended results file would contain ranked data to a scale of 4.2


❤️ Inspired by original work of [@Zunehfu](https://github.com/Zunehfu) at [uom-1st-sem-rankGen](https://github.com/LGsekara1/uom-1st-sem-rankGen.git)
