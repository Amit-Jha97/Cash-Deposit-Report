# Cash-Deposit-Report
Python-based tool to generate ASM-wise Excel reports from moneyrequest and ASM mapping files, with overall and Sunday performance, daily averages, and conditional formatting Automation reports.

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import sys


if getattr(sys, 'frozen', False):
    script_dir = os.path.dirname(sys.executable)
else:
    script_dir = os.path.dirname(os.path.abspath(__file__))

output_file = os.path.join(script_dir, "final_report.xlsx")

# GUI setup
root = tk.Tk()
root.withdraw()

# File selection
money_file = filedialog.askopenfilename(
    title="Select moneyrequest.xlsx",
    filetypes=[("Excel files", "*.xlsx")]
)
if not money_file:
    messagebox.showerror("Error", "Moneyrequest file not selected.")
    exit()

asm_file = filedialog.askopenfilename(
    title="Select asm.xlsx",
    filetypes=[("Excel files", "*.xlsx")]
)
if not asm_file:
    messagebox.showerror("Error", "ASM file not selected.")
    exit()

# Load Excel files
money_df = pd.read_excel(money_file)
asm_df = pd.read_excel(asm_file)

# Clean up code fields
money_df['code'] = money_df['code'].astype(str).str.strip().str.upper()
asm_df['code'] = asm_df['code'].astype(str).str.strip().str.upper()
asm_df = asm_df.drop_duplicates(subset='code')


money_df['type'] = money_df['type'].astype(str).str.lower()
money_df['status'] = money_df['status'].astype(str).str.lower()


filtered_df = money_df[
    (money_df['status'] == 'accepted') &
    (money_df['type'] != 'online')
].copy()


merged_df = pd.merge(filtered_df, asm_df, on='code', how='left')
merged_df['amount'] = pd.to_numeric(merged_df['amount'], errors='coerce').fillna(0)


merged_df['deposit_date'] = pd.to_datetime(merged_df['deposit_date'], errors='coerce')
merged_df = merged_df[merged_df['deposit_date'].notnull()]


merged_df['day_of_week'] = merged_df['deposit_date'].dt.day_name()
merged_df['deposit_date_only'] = merged_df['deposit_date'].dt.date


def generate_report_df(df_source):
    
    days_count = df_source['deposit_date_only'].nunique()
    days_count = max(days_count, 1)

 
    summary = df_source.groupby(['head', 'asmname'], dropna=False)['amount'].sum().reset_index()
    summary.rename(columns={'amount': 'total_amount'}, inplace=True)
    summary['average_per_day'] = (summary['total_amount'] / days_count).round(0).astype(int)

    
    latest_date = df_source['deposit_date_only'].max()
    latest_df = df_source[df_source['deposit_date_only'] == latest_date]
    latest_summary = latest_df.groupby(['head', 'asmname'], dropna=False)['amount'].sum().reset_index()
    latest_summary.rename(columns={'amount': 'latest_date_total'}, inplace=True)

 
    final = pd.merge(summary, latest_summary, on=['head', 'asmname'], how='left')
    final['latest_date_total'] = final['latest_date_total'].fillna(0).astype(int)

   
    final['remaining_amount'] = final['average_per_day'] - final['latest_date_total']

    
    final['latest_date_avg_days'] = final.apply(
        lambda row: (row['latest_date_total'] / row['average_per_day']) * 100
        if row['average_per_day'] != 0 else 0,
        axis=1
    ).round(0)
    
    return final


overall_report_df = generate_report_df(merged_df)


sunday_df = merged_df[merged_df['day_of_week'] == 'Sunday'].copy()
if not sunday_df.empty:
    sunday_report_df = generate_report_df(sunday_df)
else:
    sunday_report_df = pd.DataFrame(columns=overall_report_df.columns)


with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    overall_report_df.to_excel(writer, sheet_name="Overall_Report", index=False)
    sunday_report_df.to_excel(writer, sheet_name="Sunday_Report", index=False)


def format_excel_sheet(ws, df_source):
    percent_col_index = list(df_source.columns).index('latest_date_avg_days') + 1
    red_fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")

    for row in range(2, ws.max_row + 1):
        percent_cell = ws.cell(row=row, column=percent_col_index)
        percent_value = percent_cell.value

        if isinstance(percent_value, (int, float)):
            percent_cell.value = percent_value / 100
            percent_cell.number_format = '0%'
            if percent_cell.value < 0.9:
                for col in range(1, ws.max_column + 1):
                    ws.cell(row=row, column=col).fill = red_fill

wb = load_workbook(output_file)
ws_overall = wb["Overall_Report"]
ws_sunday = wb["Sunday_Report"]

format_excel_sheet(ws_overall, overall_report_df)
format_excel_sheet(ws_sunday, sunday_report_df)

wb.save(output_file)


messagebox.showinfo("Success", f"✅ Final report saved with two sheets in:\n{output_file}")
