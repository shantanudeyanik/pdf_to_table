import tkinter as tk
from tkinter import filedialog
import pandas as pd
import tabula
import os
import time


def browse_pdf():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    pdf_entry.delete(0, tk.END)
    pdf_entry.insert(0, file_path)


def browse_output_dir():
    output_dir = filedialog.askdirectory()
    output_entry.delete(0, tk.END)
    output_entry.insert(0, output_dir)


def export_data():
    pdf_path = pdf_entry.get()
    output_dir = output_entry.get()
    if pdf_path:
        if output_dir:
            try:
                start_time = time.time()
                file_name = os.path.splitext(os.path.basename(pdf_path))[0]

                df = tabula.read_pdf(
                    pdf_path,
                    pages="all",
                    multiple_tables=True,
                    stream=True,
                    guess=True,
                    lattice=True
                )
                num_tables = len(df)
                total_rows = 0

                for i, table in enumerate(df, start=1):
                    # Process the table and export data
                    final_df = pd.concat(table)
                    total_rows += final_df.shape[0]

                    export_format = export_var.get()
                    output_path = os.path.join(output_dir, f"{file_name}_table_{i}.{export_format.lower()}")

                    if export_format == "Excel":
                        final_df.to_excel(output_path, index=False)
                    elif export_format == "CSV":
                        final_df.to_csv(output_path, index=False)
                    elif export_format == "JSON":
                        final_df.to_json(output_path, orient="records")

                elapsed_time = round(time.time() - start_time, 2)
                status_label.config(
                    text=f"Export completed. Elapsed Time: {elapsed_time} seconds. Tables: {num_tables}. Rows: {total_rows}",
                    fg="green"
                )
            except Exception as e:
                error_label.config(text=str(e), fg="red")
        else:
            error_label.config(text="Please select an output directory.", fg="red")
    else:
        error_label.config(text="Please select a PDF file.", fg="red")


# Create the main window
window = tk.Tk()
window.title("PDF to Data Exporter")

# Set the fixed window size
window.geometry("400x350")

# Create a label and entry for PDF path
pdf_label = tk.Label(window, text="PDF File:")
pdf_label.pack()
pdf_entry = tk.Entry(window, width=50)
pdf_entry.pack()
browse_pdf_button = tk.Button(window, text="Browse", command=browse_pdf)
browse_pdf_button.pack()

# Create a label and entry for the output directory
output_label = tk.Label(window, text="Output Directory:")
output_label.pack()
output_entry = tk.Entry(window, width=50)
output_entry.pack()
browse_output_button = tk.Button(window, text="Browse", command=browse_output_dir)
browse_output_button.pack()

# Create an export format selection
format_label = tk.Label(window, text="Export Format:")
format_label.pack()
export_var = tk.StringVar(value="Excel")
export_option_menu = tk.OptionMenu(window, export_var, "Excel", "CSV", "JSON")
export_option_menu.pack()

# Create an export button
export_button = tk.Button(window, text="Export", command=export_data)
export_button.pack()

# Create a label for status messages
status_label = tk.Label(window, text="")
status_label.pack()

# Create a label for error messages
error_label = tk.Label(window, text="", fg="red")
error_label.pack()

# Run the main window loop
window.mainloop()
