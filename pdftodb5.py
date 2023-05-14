import tkinter as tk
from tkinter import filedialog
import pandas as pd
import tabula
import os
import time
import threading


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
            export_button.config(state=tk.DISABLED)  # Disable the export button

            try:
                export_thread = threading.Thread(target=export_data_thread, args=(pdf_path, output_dir))
                export_thread.start()
            except Exception as e:
                error_label.config(text=str(e), fg="red")
                export_button.config(state=tk.NORMAL)  # Enable the export button

            # Update status label to display "Converting, please wait..."
            status_label.config(text="Converting, please wait...", fg="blue")

        else:
            error_label.config(text="Please select an output directory.", fg="red")
    else:
        error_label.config(text="Please select a PDF file.", fg="red")


def export_data_thread(pdf_path, output_dir):
    try:
        start_time = time.time()

        file_name = os.path.splitext(os.path.basename(pdf_path))[0]

        df = tabula.read_pdf(pdf_path, pages="all", multiple_tables=True, stream=True, guess=True, lattice=True)
        final_df = pd.concat(df)

        export_format = export_var.get()

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

        # Update GUI elements from the main thread
        window.after(0, update_status_label, export_format, elapsed_time)
        window.after(0, update_output_path_label, output_path)
        window.after(0, enable_export_button)

    except Exception as e:
        window.after(0, update_error_label, str(e))
        window.after(0, enable_export_button)


def update_status_label(export_format, elapsed_time):
    status_label.config(
        text=f"{export_format} file exported successfully.\nElapsed Time: {elapsed_time} seconds.",
        fg="green"
    )


def update_output_path_label(output_path):
    output_path_label.config(text=f"Output Path: {os.path.abspath(output_path)}")


def update_error_label(error_message):
    error_label.config(text=error_message, fg="red")


def enable_export_button():
    export_button.config(state=tk.NORMAL)


window = tk.Tk()
window.title("PDF to Data Exporter")
window.resizable(False, False)
window.geometry("400x400")

pdf_label = tk.Label(window, text="PDF File:")
pdf_label.pack()
pdf_entry = tk.Entry(window, width=50)
pdf_entry.pack()
browse_pdf_button = tk.Button(window,text="Browse PDF", command=browse_pdf, width=25)
browse_pdf_button.pack()

output_label = tk.Label(window, text="Output Directory:")
output_label.pack(pady=(20, 0))
output_entry = tk.Entry(window, width=50)
output_entry.pack()
browse_output_button = tk.Button(window, text="Browse Output Directory", command=browse_output_dir, width=25)
browse_output_button.pack()

format_label = tk.Label(window, text="Export Format:")
format_label.pack(pady=(10, 0))
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

conversion_label = tk.Label(window, text="", fg="blue", wraplength=380)
conversion_label.pack()

window.mainloop()
