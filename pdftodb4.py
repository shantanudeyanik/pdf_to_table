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
    update_output_path()


def update_output_path():
    output_dir = output_entry.get()
    output_name = output_name_entry.get()
    output_path_label.config(text=f"Output Path: {os.path.join(output_dir, output_name)}")


def export_data():
    pdf_path = pdf_entry.get()
    output_dir = output_entry.get()
    output_name = output_name_entry.get()
    export_format = export_var.get()

    if pdf_path:
        if output_dir and output_name:
            try:
                start_time = time.time()

                file_name = os.path.splitext(output_name)[0]

                df = tabula.read_pdf(
                    pdf_path,
                    pages="all",
                    multiple_tables=True,
                    stream=True,
                    guess=True,
                    lattice=True
                )
                final_df = pd.concat(df)

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

                status_label.config(
                    text=f"{export_format} file exported successfully.\nElapsed Time: {elapsed_time} seconds.",
                    fg="green"
                )

                output_path_label.config(text=f"Output Path: {os.path.abspath(output_path)}")

            except Exception as e:
                error_label.config(text=str(e), fg="red")
        else:
            error_label.config(text="Please select an output directory and provide a file name.", fg="red")
    else:
        error_label.config(text="Please select a PDF file.", fg="red")


window = tk.Tk()
window.title("PDF to Data Exporter")
window.resizable(False, False)
window.geometry("400x450")

pdf_label = tk.Label(window, text="PDF File:")
pdf_label.pack()
pdf_entry = tk.Entry(window, width=50)
pdf_entry.pack()
browse_pdf_button = tk.Button(window, text="Browse PDF", command=browse_pdf, width=25)
browse_pdf_button.pack()

output_label = tk.Label(window, text="Output Directory:")
output_label.pack(pady=(20,0))
output_entry = tk.Entry(window, width=50)
output_entry.pack()
browse_output_button = tk.Button(window, text="Browse Output Directory", command=browse_output_dir, width=25)
browse_output_button.pack()

output_name_label = tk.Label(window, text="Output File Name:")
output_name_label.pack(pady=(20,0))
output_name_entry = tk.Entry(window, width=50)
output_name_entry.pack()
output_name_entry.bind('<KeyRelease>', lambda event: update_output_path())

format_label = tk.Label(window, text="Export Format:")
format_label.pack(pady=(10,0))
export_var = tk.StringVar(value="Excel")
export_option_menu = tk.OptionMenu(window, export_var, "Excel", "CSV", "JSON")
export_option_menu.pack()

export_button = tk.Button(window, text="Export", command=export_data)
export_button.pack()

status_label = tk.Label(window, text="", wraplength=380)
status_label.pack()

output_path_label = tk.Label(window, text="", wraplength=380)
output_path_label.pack()

error_label = tk.Label(window, text="", fg="red", wraplength=380)
error_label.pack()

footer_label = tk.Label(window, text="Prepared by Shantanu\nInfrastructure & Cyber Security Team.")
footer_label.pack(side=tk.BOTTOM)

window.mainloop()
