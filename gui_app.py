import tkinter as tk
from tkinter import filedialog, messagebox
from tkcalendar import DateEntry
from dataModelling import DataModeler
from mergeFile import FileMerger
from convertTable import TableConverter
import os


class ProcessingGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Student Data Processor")
        self.root.geometry("500x420")  # slightly taller for date picker

        self.stars_file = None
        self.canvas_file = None
        self.finish_date = None

        # --- Title ---
        tk.Label(root, text="Student Report Generator", font=("Arial", 16, "bold")).pack(pady=10)

        # --- STARS upload ---
        tk.Label(root, text="Select STARS CSV (Book1.csv):", font=("Arial", 12)).pack()
        tk.Button(root, text="Choose STARS File", command=self.load_stars).pack()
        self.stars_label = tk.Label(root, text="No file selected", fg="gray")
        self.stars_label.pack(pady=5)

        # --- CANVAS upload ---
        tk.Label(root, text="Select CANVAS CSV:", font=("Arial", 12)).pack()
        tk.Button(root, text="Choose CANVAS File", command=self.load_canvas).pack()
        self.canvas_label = tk.Label(root, text="No file selected", fg="gray")
        self.canvas_label.pack(pady=5)

        # --- Date selection ---
        tk.Label(root, text="Select Finish Date:", font=("Arial", 12)).pack(pady=10)
        self.date_picker = DateEntry(
            root,
            date_pattern="dd/mm/yyyy",
            width=12,
            background="darkblue",
            foreground="white",
            borderwidth=2
        )
        self.date_picker.pack()
        
        tk.Label(root, text="(This date will be inserted into the Finish Date column.)", fg="gray")\
            .pack(pady=3)

        # --- Run button ---
        tk.Button(root, text="Run Processing", font=("Arial", 14, "bold"),
                  bg="#4CAF50", fg="white", command=self.run_processing).pack(pady=20)

    def load_stars(self):
        file_path = filedialog.askopenfilename(
            title="Select STARS CSV",
            filetypes=[("CSV files", "*.csv")]
        )
        if file_path:
            self.stars_file = file_path
            self.stars_label.config(text=os.path.basename(file_path), fg="black")

    def load_canvas(self):
        file_path = filedialog.askopenfilename(
            title="Select CANVAS CSV",
            filetypes=[("CSV files", "*.csv")]
        )
        if file_path:
            self.canvas_file = file_path
            self.canvas_label.config(text=os.path.basename(file_path), fg="black")

    def run_processing(self):
        if not self.stars_file or not self.canvas_file:
            messagebox.showerror("Missing Files", "Please upload BOTH STARS and CANVAS files.")
            return

        # read selected date
        finish_date = self.date_picker.get()

        try:
            # 1 — Student report
            DataModeler(self.stars_file).load_data().export()

            # 2 — Merge files
            merger = FileMerger(self.canvas_file, self.stars_file)
            merger.load_data().merge().export()

            # 3 — Convert merged CSV and add finish date
            TableConverter("merged_final_scores_2.csv") \
                .load_data() \
                .unpivot() \
                .apply_outcomes() \
                .export(finish_date=finish_date)  # <-- DATE PASSED HERE

            messagebox.showinfo(
                "Success",
                "Processing complete!\n\nFiles generated:\n- student_report.xlsx\n- merged_final_scores_2.csv\n- output.xlsx"
            )

        except Exception as e:
            messagebox.showerror("Error", f"Something went wrong:\n{e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ProcessingGUI(root)
    root.mainloop()
