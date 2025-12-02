import pandas as pd

class DataModeler:
    def __init__(self, csv_stars, output_file="student_report.xlsx"):
        self.csv_stars = csv_stars
        self.output_file = output_file
        self.df = None
        self.base_columns = ["Student No", "Enrol No", "First Name", "Last Name"]

    def load_data(self):
        self.df = pd.read_csv(
            self.csv_stars,
            keep_default_na=False,
            na_values=[],
        )
        self.df.columns = self.df.columns.str.strip().str.replace("\ufeff", "")
        return self

    def build_sheet1(self):
        sheet1 = self.df[self.base_columns + ["Status"]]
        sheet1 = sheet1[
            ~sheet1["Status"].str.lower().isin(["on leave", "finished"])
        ]
        return sheet1[self.base_columns]

    def build_sheet2(self):
        sheet2 = self.df[self.base_columns + ["Status"]]
        sheet2 = sheet2[
            sheet2["Status"].str.lower().isin(["on leave", "finished"])
        ]
        return sheet2

    def build_sheet3(self):
        sheet3 = self.df[self.base_columns + ["Outstanding Fees"]].copy()

        sheet3["Outstanding Fees"] = (
            sheet3["Outstanding Fees"]
            .replace("-", "0")
            .astype(str)
            .str.replace("$", "", regex=False)
            .str.replace(",", "", regex=False)
            .str.strip()
        )

        sheet3["Outstanding Fees"] = pd.to_numeric(sheet3["Outstanding Fees"], errors="coerce")
        return sheet3[sheet3["Outstanding Fees"] >= 500]

    def export(self):
        with pd.ExcelWriter(self.output_file, engine="openpyxl") as writer:
            self.build_sheet1().to_excel(writer, "Enrolment Report", index=False)
            self.build_sheet2().to_excel(writer, "OnLeave and Finished", index=False)
            self.build_sheet3().to_excel(writer, "Outstanding_500_Plus", index=False)
        print(f"✅ {self.output_file} created successfully!")

