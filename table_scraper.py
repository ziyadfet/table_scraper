import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
import pandas as pd

# Define functions for getting webpage, extracting tables, and exporting data
def get_webpage(url):
    try:
        chrome_options = Options()
        chrome_options.add_argument("--headless")
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
        driver.get(url)
        page_source = driver.page_source
        driver.quit()
        return page_source
    except Exception as e:
        messagebox.showerror("Webpage Error", f"Failed to access the webpage. Error: {str(e)}")
        return None

def extract_tables(page_source):
    try:
        df_list = pd.read_html(page_source)
        return df_list
    except ValueError:
        messagebox.showwarning("No Tables Found", "No tables were found on the provided webpage.")
        return []

def export(df_list, selected_indices):
    if not selected_indices:
        messagebox.showwarning("Selection Error", "Please select at least one table to export")
        return
    
    if len(selected_indices) == 1:
        # Export as CSV
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if file_path:
            df_list[selected_indices[0]].to_csv(file_path, index=False)
            messagebox.showinfo("Export", f"Data exported successfully as CSV to {file_path}")
    else:
        # Export as Excel
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                for i in selected_indices:
                    sheet_name = f'Sheet{i+1}'
                    df_list[i].to_excel(writer, sheet_name=sheet_name, index=False)
            messagebox.showinfo("Export", f"Data exported successfully as Excel to {file_path}")

# Define GUI application
class TableExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Table Extractor")

        self.url_label = ttk.Label(root, text="Enter URL:")
        self.url_label.pack(pady=5)

        self.url_entry = ttk.Entry(root, width=50)
        self.url_entry.pack(pady=5)

        self.button_frame = ttk.Frame(root)
        self.button_frame.pack(pady=10)

        self.fetch_button = ttk.Button(self.button_frame, text="Fetch Tables", command=self.fetch_tables)
        self.fetch_button.pack(side=tk.LEFT, padx=5)

        self.export_button = ttk.Button(self.button_frame, text="Export Selected Tables", command=self.export_tables)
        self.export_button.pack(side=tk.LEFT, padx=5)

        self.table_frame = ttk.LabelFrame(root, text="Extracted Tables")
        self.table_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.prev_button = ttk.Button(root, text="<< Previous", command=self.show_previous_table)
        self.prev_button.pack(side=tk.LEFT, padx=5, pady=5)

        self.next_button = ttk.Button(root, text="Next >>", command=self.show_next_table)
        self.next_button.pack(side=tk.RIGHT, padx=5, pady=5)

        self.df_list = []
        self.table_vars = []
        self.current_table_index = 0

    def fetch_tables(self):
        url = self.url_entry.get()
        if not url:
            messagebox.showwarning("Input Error", "Please enter a URL")
            return
        page_source = get_webpage(url)
        if not page_source:
            return
        self.df_list = extract_tables(page_source)
        
        if not self.df_list:
            return

        for widget in self.table_frame.winfo_children():
            widget.destroy()

        self.table_vars = []
        for i, df in enumerate(self.df_list):
            var = tk.IntVar()
            chk = ttk.Checkbutton(self.table_frame, text=f"Table {i+1}", variable=var)
            self.table_vars.append(var)
        
        self.current_table_index = 0
        self.show_table()

    def show_table(self):
        for widget in self.table_frame.winfo_children():
            widget.destroy()

        if not self.df_list:
            return

        df = self.df_list[self.current_table_index]
        var = self.table_vars[self.current_table_index]

        chk = ttk.Checkbutton(self.table_frame, text=f"Table {self.current_table_index + 1}", variable=var)
        chk.pack(anchor='w')

        tree_frame = ttk.Frame(self.table_frame)
        tree_frame.pack(fill="both", expand=True)

        tree_scroll = ttk.Scrollbar(tree_frame)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        tree = ttk.Treeview(tree_frame, yscrollcommand=tree_scroll.set)
        tree.pack(fill="both", expand=True)
        tree_scroll.config(command=tree.yview)

        tree["columns"] = list(df.columns)
        tree["show"] = "headings"
        for col in df.columns:
            tree.heading(col, text=col)
            tree.column(col, width=100)

        for row in df.itertuples(index=False):
            tree.insert("", "end", values=row)

    def show_previous_table(self):
        if self.current_table_index > 0:
            self.current_table_index -= 1
            self.show_table()

    def show_next_table(self):
        if self.current_table_index < len(self.df_list) - 1:
            self.current_table_index += 1
            self.show_table()

    def export_tables(self):
        selected_indices = [i for i, var in enumerate(self.table_vars) if var.get() == 1]
        export(self.df_list, selected_indices)

# Create and run the application
if __name__ == "__main__":
    root = tk.Tk()
    app = TableExtractorApp(root)
    root.mainloop()
