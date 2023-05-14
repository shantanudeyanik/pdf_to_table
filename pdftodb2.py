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

                # Extract the source filename without extension
                file_name = os.path.splitext(os.path.basename(pdf_path))[0]

                # Read PDF and concatenate tables
                df = tabula.read_pdf(
                    pdf_path,
                    pages="all",
                    multiple_tables=True,
                    stream=True,
                    guess=True,
                    lattice=True
                )
                final_df = pd.concat(df)

                export_format = export_var.get()

                # Set the output path based on the selected format and source filename
                if export_format == "Excel":
                    output_path = os.path.join(output_dir, f"{file_name}.xlsx")
                    final_df.to_excel(output_path, index=False)
                elif export_format == "CSV":
                    output_path = os.path.join(output_dir, f"{file_name}.csv")
                    final_df.to_csv(output_path, index=False)
                elif export_format == "JSON":
                    output_path = os.path.join(output_dir, f"{file_name}.json")
                    final_df.to_json(output_path, orient="records")

                elapsed_time = round(time.time() - start_time, 2)

                # Update the status label with export success message and elapsed time
                status_label.config(
                    text=f"{export_format} file exported successfully. Elapsed Time: {elapsed_time} seconds.",
                    fg="green"
                )
            except Exception as e:
                # Display error message in the error label
                error_label.config(text=str(e), fg="red")
        else:
            error_label.config(text="Please select an output directory.", fg="red")
    else:
        error_label.config(text="Please select a PDF file.", fg="red")


window = tk.Tk()
window.title("PDF to Data Exporter")
window.geometry("400x350")

# PDF file label and entry
pdf_label = tk.Label(window, text="PDF File:")
pdf_label.pack()
pdf_entry = tk.Entry(window, width=50)
pdf_entry.pack()
browse_pdf_button = tk.Button(window, text="Browse", command=browse_pdf)
browse_pdf_button.pack()

# Output directory label and entry
output_label = tk.Label(window, text="Output Directory:")
output_label.pack()
output_entry = tk.Entry(window, width=50)
output_entry.pack()
browse_output_button = tk.Button(window, text="Browse", command=browse_output_dir)
browse_output_button.pack()

# Export format selection
format_label = tk.Label(window, text="Export Format:")
format_label.pack()
export_var = tk.StringVar(value="Excel")
export_option_menu = tk.OptionMenu(window, export_var, "Excel", "CSV", "JSON")
export_option_menu.pack()

# Export button
export_button = tk.Button(window, text="Export", command=export_data)
export_button.pack()

# Status label for success messages
status_label = tk.Label(window, text="")
status_label.pack()

# Error label for displaying error messages
error_label = tk.Label(window, text="", fg="red")
error_label.pack()

window.mainloop()
