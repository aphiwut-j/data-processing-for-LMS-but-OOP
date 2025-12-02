# Data Processing for LMS (OOP Approach)

This repository provides a suite of Python scripts designed for data processing tasks relevant to Learning Management Systems (LMS), with a focus on object-oriented programming principles. The toolkit facilitates the handling, merging, transformation, and reporting of student result data typically exported from systems such as CANVAS and STARS.

## Contents

- `main.py`: Orchestrates the workflow by running student report generation, merging result files, and converting merged scores to formatted Excel output.
- `dataModelling.py`: Builds different data sheets (such as enrolment report, students on leave/finished, outstanding fees report) from the STARS CSV export and exports them to an Excel file.
- `mergeFile.py`: Merges the results from CANVAS and STARS by student ID, cleans and aligns related columns, filters for final score columns, and exports the merged results to CSV.
- `convertTable.py`: Converts the merged CSV of scores to a "long" format Excel file, applies LMS outcome logic (pass, fail, credit, etc.), and automatically formats the results with colors for easy review.

## Features

- **Student Report Generation:** Creates comprehensive Excel reports including current enrolments, leave/finished status, and students with high outstanding fees.
- **Result Merging:** Combines learning results from separate sources and aligns based on student identifiers.
- **Automated Outcome Calculation:** Converts raw scores into LMS outcomes ("Pass", "Fail", "Credit Transfer", etc.) and visually formats results for rapid understanding.
- **Excel Export with Color Formatting:** Output files are professionally styled for easy navigation and interpretation.

## Setup Guidelines

Follow these steps to set up your environment for running the scripts:

### 1. **Clone the Repository**
```bash
git clone https://github.com/aphiwut-j/data-processing-for-LMS-but-OOP.git
cd data-processing-for-LMS-but-OOP
```

### 2. **Create a Python Virtual Environment** (optional but recommended)
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

### 3. **Install Dependencies**
Install the required Python packages using pip:
```bash
pip install pandas openpyxl
```

### 4. **Configure Input Paths**
Update the following variables in `main.py` to the correct paths for your data CSV files:
```python
BOOK1 = r"Path/to/Book1.csv" # STARS export
MARKS = r"Path/to/Marks.csv" # CANVAS export
```

### 5. **Run the Workflow**
Start data processing by running:
```bash
python main.py
```
This will generate the following output files (in the project root):
- `student_report.xlsx` (Excel report)
- `merged_final_scores_2.csv` (merged scores)
- `output.xlsx` (final results with status color formatting)

## Example Workflow

1. Generate student report.
2. Merge result files from CANVAS and STARS.
3. Convert merged scores to Excel, automatically assigning LMS outcome statuses.

## License

Add licensing information as appropriate.

## Author

Developed by [aphiwut-j](https://github.com/aphiwut-j)