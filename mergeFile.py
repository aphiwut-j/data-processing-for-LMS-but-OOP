import pandas as pd
import re

class FileMerger:
    def __init__(self, csv_canvas, csv_stars, output_file="merged_final_scores_2.csv"):
        self.csv_canvas = csv_canvas
        self.csv_stars = csv_stars
        self.output_file = output_file
        self.df1 = None
        self.df2 = None
        self.merged_df = None

    @staticmethod
    def clean_id(series):
        return (
            series.astype(str)
            .str.strip()
            .str.replace(r"\.0$", "", regex=True)
            .str.replace(r"\s+", "", regex=True)
            .str.replace(r"[^\d]", "", regex=True)
        )

    def load_data(self):
        self.df1 = pd.read_csv(self.csv_canvas)
        self.df2 = pd.read_csv(self.csv_stars)

        self.df1.columns = self.df1.columns.str.strip()
        self.df2.columns = self.df2.columns.str.strip()

        self.df1["SIS User ID"] = self.clean_id(self.df1["SIS User ID"])
        self.df2["Student No"] = self.clean_id(self.df2["Student No"])

        self.df2 = self.df2.rename(columns={"Student No": "SIS User ID"})
        return self

    def merge(self):
        self.merged_df = pd.merge(
            self.df1,
            self.df2,
            on="SIS User ID",
            how="inner"
        )
        return self

    def filter_final_scores(self):
        base_columns = ["Enrol No", "SIS User ID", "SIS Login ID", "Section"]

        final_score_cols = [
            col for col in self.merged_df.columns
            if "final score" in col.lower()
            and "unposted" not in col.lower()
            and "roll call" not in col.lower()
            and col.lower() != "final score"
        ]

        output_columns = base_columns + final_score_cols

        final_output = self.merged_df[output_columns].copy()

        for col in final_output.columns:
            if "final" in col.lower():
                new_name = col.split()[0]
                final_output.rename(columns={col: new_name}, inplace=True)

        return final_output

    def export(self):
        df_out = self.filter_final_scores()
        df_out.to_csv(self.output_file, index=False)
        print(f"✅ Merge complete: {self.output_file}")

