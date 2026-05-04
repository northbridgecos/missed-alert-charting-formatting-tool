import tkinter as tk
from tkinter import filedialog, messagebox
import os
from process import process

def run():
    try:
        input_file = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx")]
        )

        if not input_file:
            return

        output_file = filedialog.asksaveasfilename(
            title="Save Output File",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"Formatted-{os.path.basename(input_file)}"
        )

        if not output_file:
            return

        process(input_file, output_file)

        messagebox.showinfo("Success", "File processed successfully!")

    except Exception as e:
        messagebox.showerror("Error", str(e))

# ---------------- GUI ----------------

root = tk.Tk()
root.title("Missed Alert Charting Formatting Tool")
root.geometry("400x100")

btn = tk.Button(root, text="Select File & Run", command=run)
btn.pack(pady=20)

root.mainloop()