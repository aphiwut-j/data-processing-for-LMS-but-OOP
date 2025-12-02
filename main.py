from dataModelling import DataModeler
from mergeFile import FileMerger
from convertTable import TableConverter

# Book1, csv from STARS
BOOK1 = r"C:\Users\aphiwut.j\Documents\Learning\resulting processing\Book1.csv"
# Marks, csv from CANVAS
MARKS = r"C:\Users\aphiwut.j\Documents\Learning\resulting processing\2025-11-26T1522_Marks-GHC_25-T5_SIT40521_G-CKM.csv"

if __name__ == "__main__":
    # 1 — create student_report.xlsx
    DataModeler(BOOK1).load_data().export()

    # 2 — merge CANVAS + STAR
    FileMerger(MARKS, BOOK1).load_data().merge().export()

    # 3 — convert merged scores to Excel with colours
    TableConverter("merged_final_scores_2.csv").load_data().unpivot().apply_outcomes().export()
