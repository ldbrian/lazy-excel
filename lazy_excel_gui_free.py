import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import json
from datetime import datetime

# Create main window
root = tk.Tk()
root.title("Lazy Excel Toolbox (Free Trial)")  # 修改标题
root.geometry("600x500")  # 调整窗口高度

# File list display box
file_listbox = tk.Listbox(root, selectmode=tk.MULTIPLE, width=60, height=10)
file_listbox.pack(pady=10)

# Select files button
def select_files():
    file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")])
    for path in file_paths:
        if path not in file_listbox.get(0, tk.END):
            file_listbox.insert(tk.END, path)
    print(f"Files added: {file_listbox.get(0, tk.END)}")  # 输出文件列表到终端

# 清空文件列表和文本框的函数
def clear_file_list():
    file_listbox.delete(0, tk.END)
    messagebox.showinfo("Info", "File list cleared.")

# 按钮容器框架
button_frame = tk.Frame(root)
button_frame.pack(pady=10)

# 选择文件按钮
tk.Button(button_frame, text="Select Excel Files", command=select_files).pack(side="left", padx=10)

# 清空文件列表按钮
tk.Button(button_frame, text="Clear File List", command=clear_file_list).pack(side="left", padx=10)

# Feature selection (checkboxes)
features = {
    "merge": tk.BooleanVar(),
    "clean": tk.BooleanVar(),
    "format_adjust": tk.BooleanVar(),
    "rename_columns": tk.BooleanVar(),
    "generate_summary": tk.BooleanVar()
}

tk.Checkbutton(root, text="Merge Files", variable=features["merge"]).pack(anchor="w", padx=20)
tk.Checkbutton(root, text="Clean Data (Remove Empty Rows/Columns)", variable=features["clean"]).pack(anchor="w", padx=20)
tk.Checkbutton(root, text="Quick Format Adjustment (Bold Header, Auto Column Width)", variable=features["format_adjust"]).pack(anchor="w", padx=20)

# Batch Rename Columns Checkbox and Input Box
def toggle_rename_entry():
    if features["rename_columns"].get():
        rename_entry.pack(pady=5, padx=40, anchor="w")  # Show input box
    else:
        rename_entry.pack_forget()  # Hide input box

rename_frame = tk.Frame(root)  # Create a container frame
rename_frame.pack(anchor="w", padx=20)

tk.Checkbutton(rename_frame, text="Batch Rename Columns (One-Click Replace)", variable=features["rename_columns"], command=toggle_rename_entry).pack(anchor="w")

# Input Box: Column Mapping Rules
rename_entry = tk.Entry(rename_frame, width=50)
rename_entry.insert(0, "OldColumn1:NewColumn1,OldColumn2:NewColumn2")  # Provide default hint

tk.Checkbutton(root, text="Generate Summary Template (Sum, Average, Count)", variable=features["generate_summary"]).pack(anchor="w", padx=20)

# File processing limit logic
def load_daily_limit():
    """Load daily file processing limit from a JSON file."""
    limit_file = "daily_limit.json"
    if not os.path.exists(limit_file):
        return {"date": str(datetime.now().date()), "processed_count": 0}
    with open(limit_file, "r") as f:
        return json.load(f)

def save_daily_limit(data):
    """Save daily file processing limit to a JSON file."""
    with open("daily_limit.json", "w") as f:
        json.dump(data, f)

# Function Implementation
def process_files():
    files = file_listbox.get(0, tk.END)
    if not files:
        messagebox.showwarning("Warning", "Please select Excel files first")
        return

    # Load daily limit
    daily_limit = load_daily_limit()
    current_date = str(datetime.now().date())

    # Reset count if the date has changed
    if daily_limit["date"] != current_date:
        daily_limit = {"date": current_date, "processed_count": 0}

    # Check if daily limit is exceeded
    if daily_limit["processed_count"] + len(files) > 10:
        remaining = 10 - daily_limit["processed_count"]
        messagebox.showwarning("Warning", f"Free Trial version supports up to 10 files per day. You can process {remaining} more files today.")
        return

    try:
        # Update processed count
        daily_limit["processed_count"] += len(files)
        save_daily_limit(daily_limit)

        # Debug output
        print(f"Daily processed count: {daily_limit['processed_count']}")

        # 添加调试信息以确认文件数量限制逻辑是否被触发
        print(f"Selected files: {len(files)}")  # 输出文件数量到终端

        if features["merge"].get():
            try:
                merged_df = pd.DataFrame()
                for file in files:
                    if not os.path.exists(file):
                        messagebox.showerror("Error", f"File not found:\n{file}")
                        return
                    try:
                        df = pd.read_excel(file)
                    except Exception as e:
                        messagebox.showerror("Error", f"Failed to read file:\n{file}\nError: {str(e)}")
                        return
                    merged_df = pd.concat([merged_df, df], ignore_index=True, sort=False)

                if features["clean"].get():
                    merged_df.dropna(how='all', axis=0, inplace=True)
                    merged_df.dropna(how='all', axis=1, inplace=True)

                save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                         filetypes=[("Excel files", "*.xlsx")],
                                                         title="Save Merged File As")
                if not save_path:
                    messagebox.showwarning("Warning", "No save path selected")
                    return

                merged_df.to_excel(save_path, index=False)
                messagebox.showinfo("Success", f"Merged file saved as:\n{save_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Error during merge:\n{str(e)}")
            return

        if features["clean"].get():
            try:
                for file in files:
                    if not os.path.exists(file):
                        messagebox.showerror("Error", f"File not found:\n{file}")
                        return
                    try:
                        df = pd.read_excel(file)
                    except Exception as e:
                        messagebox.showerror("Error", f"Failed to read file:\n{file}\nError: {str(e)}")
                        return

                    df.dropna(how='all', axis=0, inplace=True)
                    df.dropna(how='all', axis=1, inplace=True)

                    cleaned_path = os.path.splitext(file)[0] + "_cleaned.xlsx"
                    df.to_excel(cleaned_path, index=False)

                messagebox.showinfo("Success", "Data cleaning completed, files saved in original directories")
            except Exception as e:
                messagebox.showerror("Error", f"Error during cleaning:\n{str(e)}")
            return

        if features["format_adjust"].get():
            try:
                for file in files:
                    df = pd.read_excel(file)
                    writer = pd.ExcelWriter(file.replace(".xlsx", "_formatted.xlsx"), engine='xlsxwriter')
                    df.to_excel(writer, index=False, sheet_name='Sheet1')
                    workbook = writer.book
                    worksheet = writer.sheets['Sheet1']
                    header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})
                    for col_num, value in enumerate(df.columns.values):
                        column_width = max(df[value].astype(str).map(len).max(), len(value)) + 2
                        worksheet.write(0, col_num, value, header_format)
                        worksheet.set_column(col_num, col_num, column_width)
                    writer.save()
                messagebox.showinfo("Success", "Format adjustment completed, files saved in original directories")
            except Exception as e:
                messagebox.showerror("Error", f"Error during format adjustment:\n{str(e)}")
            return

        if features["rename_columns"].get():
            try:
                mapping_text = rename_entry.get()
                if not mapping_text:
                    messagebox.showwarning("Warning", "Please provide column mapping rules")
                    return
                column_mapping = dict(item.split(":") for item in mapping_text.split(","))
                for file in files:
                    df = pd.read_excel(file)
                    missing_columns = [col for col in column_mapping.keys() if col not in df.columns]
                    if missing_columns:
                        messagebox.showwarning("Warning", f"The following columns are missing:\n{', '.join(missing_columns)}")
                        continue
                    df.rename(columns=column_mapping, inplace=True)
                    renamed_path = file.replace(".xlsx", "_renamed.xlsx")
                    df.to_excel(renamed_path, index=False)
                messagebox.showinfo("Success", "Column renaming completed, files saved in original directories")
            except Exception as e:
                messagebox.showerror("Error", f"Error during column renaming:\n{str(e)}")
            return

        if features["generate_summary"].get():
            try:
                for file in files:
                    df = pd.read_excel(file)
                    numeric_columns = df.select_dtypes(include=['number']).columns
                    if numeric_columns.empty:
                        messagebox.showwarning("Warning", f"No numeric columns found in file:\n{file}")
                        continue
                    summary = pd.DataFrame({
                        "Column Name": df.columns,
                        "Sum": [df[col].sum() if col in numeric_columns else None for col in df.columns],
                        "Average": [df[col].mean() if col in numeric_columns else None for col in df.columns],
                        "Count": [df[col].count() for col in df.columns]
                    })
                    summary_path = file.replace(".xlsx", "_summary.xlsx")
                    summary.to_excel(summary_path, index=False)
                messagebox.showinfo("Success", "Summary template generation completed, files saved in original directories")
            except Exception as e:
                messagebox.showerror("Error", f"Error during summary generation:\n{str(e)}")
            return

        # 清空文件列表和文本框
        clear_file_list()
        messagebox.showinfo("Success", "Files processed successfully. File list cleared.")
    except Exception as e:
        messagebox.showerror("Error", f"Error during file processing:\n{str(e)}")
        # 保留文件列表
        print("File processing failed. File list retained.")

# Adjust "Start Processing" button size
tk.Button(root, text="Start Processing", command=process_files, bg="#4CAF50", fg="white", height=2, width=20).pack(pady=20)

# Main loop
root.mainloop()
