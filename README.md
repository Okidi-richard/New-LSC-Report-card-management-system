# Uganda Secondary School Automatic Report Card System

**Aligned with the NCDC / UNEB Competency-Based Curriculum (CBC)**  
Lower Secondary (S1 – S4 / UCE)

Now available as a **web app** — open it in any browser. No command-line skills required for daily use.

### Want to use it from a phone?
See **DEPLOY.md** for free online hosting instructions. After one-time setup, any phone can open the system in Chrome and generate report cards.

---

## Two ways to use the system

### Option A — Web App (Recommended)

```bash
# Install dependencies (once)
pip install flask pandas openpyxl reportlab

# Start the app
python run_app.py
```

Then open your browser at:

**http://127.0.0.1:5000**

You will see a clean interface where you can:
1. Download the Excel template
2. Upload the filled Excel
3. Generate all report cards
4. Preview and download individual PDFs or a ZIP of the whole class

### Option B — Command Line

```bash
python report_card_system.py --create-template
# Fill the Excel file…
python report_card_system.py --excel templates/marks_entry_template.xlsx --output generated_reports
```

---

## What this system does

- Automatically calculates final subject scores using the official formula:  
  **Final Score = (Formative × 20%) + (Summative × 80%)**
- Assigns the correct letter grade (A–E) with official descriptors
- Generates professional, printable **PDF report cards** for every learner
- Produces a class summary for quick overview
- Uses a simple Excel template that teachers can fill without any coding knowledge

### Official Grading Scale (UNEB / NCDC)

| Score   | Grade | Descriptor   | Meaning |
|---------|-------|--------------|---------|
| 80–100  | **A** | Exceptional  | Demonstrates extraordinary competency by applying knowledge and skills innovatively and creatively in real-life situations |
| 70–79   | **B** | Outstanding  | Demonstrates a high level of competency by effectively applying acquired knowledge and skills in real-life situations |
| 60–69   | **C** | Satisfactory | Demonstrates an adequate level of competency in applying knowledge and skills in real-life situations |
| 50–59   | **D** | Basic        | Demonstrates a minimum level of competency in applying knowledge and skills in real-life situations |
| 0–49    | **E** | Elementary   | Demonstrates below the basic level of competency in applying knowledge and skills in real-life situations |

---

## Excel Template Sheets

| Sheet            | Purpose |
|------------------|---------|
| **School Settings** | School name, motto, address, term, year, report date |
| **Students**        | Admission number, names, class, stream, attendance, comments |
| **Marks**           | Admission No + Subject + Formative (out of 100) + Summative (out of 100) + optional comment |
| **Grading Guide**   | Official scale for reference |

---

## Folder Structure

```
uganda_report_card_system/
├── run_app.py                     # ← Start the web app with this
├── report_card_system.py          # Core grading & PDF engine
├── app/
│   ├── app.py                     # Flask web application
│   ├── templates/                 # HTML pages
│   ├── uploads/                   # Uploaded Excel files
│   └── outputs/                   # Generated PDFs
├── templates/
│   └── marks_entry_template.xlsx  # Data entry workbook
├── generated_reports/             # CLI output folder
└── README.md
```

---

## Notes for Ugandan Schools

- The system follows the current **Lower Secondary Competency-Based Curriculum** reporting style.
- Project work / Activities of Integration should be included in the **Formative** mark.
- At S1–S2 learners normally take 11 compulsory subjects + 1 elective; at S3–S4 they take 7 compulsory + up to 2 electives. The system handles any number of subjects.
- For A-Level (S5–S6) the grading is still transitioning; this version is optimised for O-Level (S1–S4).

---

**Generated with care for Uganda’s education system.**
