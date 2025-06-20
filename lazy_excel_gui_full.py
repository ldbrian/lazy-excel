import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import matplotlib.pyplot as plt
from fpdf import FPDF

# Create main window
root = tk.Tk()
root.title("Lazy Excel Toolbox (Full)")
root.geometry("600x700")  # 调整窗口高度

# File list display box
file_listbox = tk.Listbox(root, selectmode=tk.MULTIPLE, width=60, height=10)
file_listbox.pack(pady=10)

# Select files button
def select_files():
    file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")])
    for path in file_paths:
        if path not in file_listbox.get(0, tk.END):
            file_listbox.insert(tk.END, path)

# 清空文件列表和文本框的函数
def clear_file_list():
    """清空文件列表，不弹出提示框"""
    file_listbox.delete(0, tk.END)

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
    "generate_summary": tk.BooleanVar(),
    "enhanced_template_export": tk.BooleanVar(),
    "advanced_data_analysis": tk.BooleanVar(),
    "smart_multi_file_merge": tk.BooleanVar(),
    "one_click_format_beautification": tk.BooleanVar(),
    "template_export_with_logo": tk.BooleanVar(),
    "data_analysis_report": tk.BooleanVar(),
    "smart_cross_table_merge": tk.BooleanVar(),
    "enterprise_format_beautification": tk.BooleanVar(),
    "authorization_management": tk.BooleanVar()
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

# Advanced Features Checkboxes
tk.Checkbutton(root, text="Enhanced Template Export", variable=features["enhanced_template_export"]).pack(anchor="w", padx=20)
tk.Checkbutton(root, text="Advanced Data Analysis", variable=features["advanced_data_analysis"]).pack(anchor="w", padx=20)
tk.Checkbutton(root, text="Smart Multi-File Merge (Enhanced)", variable=features["smart_multi_file_merge"]).pack(anchor="w", padx=20)
tk.Checkbutton(root, text="One-Click Format Beautification (Enhanced)", variable=features["one_click_format_beautification"]).pack(anchor="w", padx=20)
tk.Checkbutton(root, text="Template Export (With LOGO + Auto Naming)", variable=features["template_export_with_logo"]).pack(anchor="w", padx=20)
tk.Checkbutton(root, text="Data Analysis Report (Charts + PDF)", variable=features["data_analysis_report"]).pack(anchor="w", padx=20)
tk.Checkbutton(root, text="Smart Cross-Table Merge (Keyword Matching)", variable=features["smart_cross_table_merge"]).pack(anchor="w", padx=20)
tk.Checkbutton(root, text="Enterprise Format Beautification", variable=features["enterprise_format_beautification"]).pack(anchor="w", padx=20)
tk.Checkbutton(root, text="Authorization Management (Team Usage)", variable=features["authorization_management"]).pack(anchor="w", padx=20)

# Function Implementation
def process_files():
    files = file_listbox.get(0, tk.END)
    if not files:
        messagebox.showwarning("Warning", "Please select Excel files first")
        return

    try:
        if features["merge"].get():
            merged_df = pd.DataFrame()
            for file in files:
                df = pd.read_excel(file)
                merged_df = pd.concat([merged_df, df], ignore_index=True, sort=False)
            
            # 允许用户选择保存的文件目录和文件名
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                     filetypes=[("Excel files", "*.xlsx")],
                                                     title="Save Merged File As")
            if not save_path:
                messagebox.showwarning("Warning", "No save path selected")
                return

            merged_df.to_excel(save_path, index=False)
            messagebox.showinfo("Success", f"Merged file saved as:\n{save_path}")

        if features["enhanced_template_export"].get():
            for file in files:
                df = pd.read_excel(file)
                enhanced_template_path = file.replace(".xlsx", "_enhanced_template.xlsx")
                df.to_excel(enhanced_template_path, index=False)
            messagebox.showinfo("Success", "Enhanced template export completed")

        if features["advanced_data_analysis"].get():
            analysis_results = []
            for file in files:
                df = pd.read_excel(file)
                analysis_results.append({
                    "File": file,
                    "Row Count": len(df),
                    "Column Count": len(df.columns),
                    "Numeric Columns": len(df.select_dtypes(include=['number']).columns),
                    "Empty Rows": df.isnull().all(axis=1).sum(),
                    "Empty Columns": df.isnull().all(axis=0).sum()
                })
            analysis_df = pd.DataFrame(analysis_results)
            analysis_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                         filetypes=[("Excel files", "*.xlsx")],
                                                         title="Save Analysis Results As")
            if analysis_path:
                analysis_df.to_excel(analysis_path, index=False)
                messagebox.showinfo("Success", f"Advanced analysis results saved as:\n{analysis_path}")

        if features["smart_multi_file_merge"].get():
            merged_df = pd.DataFrame()
            for file in files:
                df = pd.read_excel(file)
                merged_df = pd.concat([merged_df, df], ignore_index=True, sort=False)
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                     filetypes=[("Excel files", "*.xlsx")],
                                                     title="Save Smart Merged File As")
            if save_path:
                merged_df.to_excel(save_path, index=False)
                messagebox.showinfo("Success", f"Smart merged file saved as:\n{save_path}")

        if features["one_click_format_beautification"].get():
            for file in files:
                df = pd.read_excel(file)
                writer = pd.ExcelWriter(file.replace(".xlsx", "_beautified.xlsx"), engine='xlsxwriter')
                df.to_excel(writer, index=False, sheet_name='Sheet1')
                workbook = writer.book
                worksheet = writer.sheets['Sheet1']
                header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})
                for col_num, value in enumerate(df.columns.values):
                    column_width = max(df[value].astype(str).map(len).max(), len(value)) + 2
                    worksheet.write(0, col_num, value, header_format)
                    worksheet.set_column(col_num, col_num, column_width)
                writer.save()
            messagebox.showinfo("Success", "One-click format beautification completed")

        if features["template_export_with_logo"].get():
            for file in files:
                df = pd.read_excel(file)
                template_path = file.replace(".xlsx", "_template_with_logo.xlsx")
                writer = pd.ExcelWriter(template_path, engine='xlsxwriter')
                df.to_excel(writer, index=False, sheet_name='Sheet1')
                workbook = writer.book
                worksheet = writer.sheets['Sheet1']
                worksheet.insert_image('A1', 'logo.png')  # 插入 LOGO
                writer.save()
            messagebox.showinfo("Success", "Template export with LOGO completed")

        if features["data_analysis_report"].get():
            for file in files:
                df = pd.read_excel(file)
                numeric_columns = df.select_dtypes(include=['number']).columns
                if numeric_columns.empty:
                    messagebox.showwarning("Warning", f"No numeric columns found in file:\n{file}")
                    continue

                # Generate charts
                for col in numeric_columns:
                    plt.figure()
                    df[col].plot(kind='bar', title=f"Analysis of {col}")
                    chart_path = file.replace(".xlsx", f"_{col}_chart.png")
                    plt.savefig(chart_path)
                    plt.close()

                # Generate PDF report
                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("Arial", size=12)
                pdf.cell(200, 10, txt="Data Analysis Report", ln=True, align='C')
                for col in numeric_columns:
                    chart_path = file.replace(".xlsx", f"_{col}_chart.png")
                    pdf.image(chart_path, x=10, y=None, w=180)
                report_path = file.replace(".xlsx", "_analysis_report.pdf")
                pdf.output(report_path)
            messagebox.showinfo("Success", "Data analysis report generated")

        if features["smart_cross_table_merge"].get():
            merged_df = pd.DataFrame()
            for file in files:
                df = pd.read_excel(file)
                if "OrderID" in df.columns:  # 假设关键词为 "OrderID"
                    merged_df = pd.merge(merged_df, df, on="OrderID", how="outer") if not merged_df.empty else df
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                     filetypes=[("Excel files", "*.xlsx")],
                                                     title="Save Smart Cross-Table Merged File As")
            if save_path:
                merged_df.to_excel(save_path, index=False)
                messagebox.showinfo("Success", f"Smart cross-table merged file saved as:\n{save_path}")

        if features["enterprise_format_beautification"].get():
            for file in files:
                df = pd.read_excel(file)
                writer = pd.ExcelWriter(file.replace(".xlsx", "_enterprise_beautified.xlsx"), engine='xlsxwriter')
                df.to_excel(writer, index=False, sheet_name='Sheet1')
                workbook = writer.book
                worksheet = writer.sheets['Sheet1']
                header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#D9EAD3'})
                for col_num, value in enumerate(df.columns.values):
                    column_width = max(df[value].astype(str).map(len).max(), len(value)) + 2
                    worksheet.write(0, col_num, value, header_format)
                    worksheet.set_column(col_num, col_num, column_width)
                writer.save()
            messagebox.showinfo("Success", "Enterprise format beautification completed")

        if features["authorization_management"].get():
            messagebox.showinfo("Info", "Authorization management is enabled. Please contact the administrator for team usage.")

        # 清空文件列表和文本框
        clear_file_list()
        messagebox.showinfo("Success", "Files processed successfully.")
    except Exception as e:
        messagebox.showerror("Error", f"Error during file processing:\n{str(e)}")
        # 保留文件列表
        print("File processing failed. File list retained.")

# Adjust "Start Processing" button size
tk.Button(root, text="Start Processing", command=process_files, bg="#4CAF50", fg="white", height=2, width=20).pack(pady=20)

# Main loop
root.mainloop()
