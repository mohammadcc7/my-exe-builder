import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

def convert_excel_to_txt_batch():
    root = tk.Tk()
    root.withdraw()
    
    # السماح باختيار أكثر من ملف في نفس الوقت
    file_paths = filedialog.askopenfilenames(
        title="اختر ملفات Excel (يمكنك تحديد أكثر من ملف)",
        filetypes=[("Excel Files", "*.xlsx;*.xls")]
    )
    
    if not file_paths:
        return

    converted_files = []
    errors = []

    for file_path in file_paths:
        try:
            excel_file = pd.ExcelFile(file_path)
            base_dir = os.path.dirname(file_path)
            base_name = os.path.splitext(os.path.basename(file_path))[0]

            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(excel_file, sheet_name=sheet_name)
                out_name = f"{base_name}_{sheet_name}_utf8.txt" if len(excel_file.sheet_names) > 1 else f"{base_name}_utf8.txt"
                out_path = os.path.join(base_dir, out_name)
                
                # حفظ كملف نصي TXT بفاصل Tab وترقيم UTF-8
                df.to_csv(out_path, sep='\t', index=False, encoding='utf-8-sig')
                converted_files.append(out_name)
        except Exception as e:
            errors.append(f"{os.path.basename(file_path)}: {str(e)}")

    # إظهار النتيجة النهائية
    msg = ""
    if converted_files:
        msg += f"تم تحويل {len(converted_files)} ملف/ورقة بنجاح:\n" + "\n".join(converted_files)
    if errors:
        msg += f"\n\nحدثت أخطاء في الملفات التالية:\n" + "\n".join(errors)

    if msg:
        messagebox.showinfo("النتيجة", msg)

if __name__ == "__main__":
    convert_excel_to_txt_batch()
