import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

def convert_excel_to_unicode():
    root = tk.Tk()
    root.withdraw()
    
    file_path = filedialog.askopenfilename(
        title="اختر ملف Excel",
        filetypes=[("Excel Files", "*.xlsx;*.xls")]
    )
    
    if not file_path:
        return

    try:
        excel_file = pd.ExcelFile(file_path)
        base_dir = os.path.dirname(file_path)
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        converted_files = []

        for sheet_name in excel_file.sheet_names:
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
            out_name = f"{base_name}_{sheet_name}_utf8.csv" if len(excel_file.sheet_names) > 1 else f"{base_name}_utf8.csv"
            out_path = os.path.join(base_dir, out_name)
            df.to_csv(out_path, index=False, encoding='utf-8-sig')
            converted_files.append(out_name)

        messagebox.showinfo("نجاح", f"تم تحويل الملف بنجاح:\n" + "\n".join(converted_files))
    except Exception as e:
        messagebox.showerror("خطأ", f"حدث خطأ أثناء التحويل:\n{str(e)}")

if __name__ == "__main__":
    convert_excel_to_unicode()
