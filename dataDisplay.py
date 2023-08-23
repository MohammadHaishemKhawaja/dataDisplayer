import tkinter as tk
from tkinter import filedialog
import pandas as pd

class ExcelProgram:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel-like Program")

        self.load_button = tk.Button(self.root, text="Load Spreadsheet", command=self.load_spreadsheet)
        self.load_button.pack(pady=10)

        self.data_table = None
        self.edited_data = None

    def load_spreadsheet(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            try:
                self.data_table = pd.read_excel(file_path, engine='openpyxl')
                self.edited_data = self.data_table.copy()
                self.display_data()
            except Exception as e:
                print(f"An error occurred: {e}")

    def display_data(self):
        if self.data_table is not None:
            if hasattr(self, "canvas"):
                self.canvas.destroy()

            self.canvas = tk.Canvas(self.root)
            self.canvas.pack(fill=tk.BOTH, expand=True)

            self.scrollbar = tk.Scrollbar(self.root, orient=tk.VERTICAL, command=self.canvas.yview)
            self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            self.canvas.configure(yscrollcommand=self.scrollbar.set)

            self.data_frame = tk.Frame(self.canvas)
            self.canvas.create_window((0, 0), window=self.data_frame, anchor=tk.NW)

            rows, cols = self.edited_data.shape

            for i in range(rows + 1):
                for j in range(cols):
                    if i == 0:
                        label = tk.Label(self.data_frame, text=self.edited_data.columns[j], relief=tk.RIDGE)
                        label.grid(row=i, column=j, sticky="nsew")
                    else:
                        entry = tk.Entry(self.data_frame)
                        entry.insert(0, str(self.edited_data.iloc[i - 1, j]))
                        entry.grid(row=i, column=j, sticky="nsew")

            self.data_frame.update_idletasks()
            self.canvas.config(scrollregion=self.canvas.bbox("all"))

            save_button = tk.Button(self.root, text="Save Changes", command=self.save_changes)
            save_button.pack()

    def save_changes(self):
        if self.edited_data is not None:
            for i in range(self.data_frame.grid_size()[0] - 1):
                for j in range(self.edited_data.shape[1]):
                    entry = self.data_frame.grid_slaves(row=i + 1, column=j)[0]
                    self.edited_data.iat[i, j] = entry.get()

            self.edited_data.to_excel("edited_data.xlsx", index=False)
            print("Changes saved.")

def main():
    root = tk.Tk()
    app = ExcelProgram(root)
    root.mainloop()

if __name__ == "__main__":
    main()
