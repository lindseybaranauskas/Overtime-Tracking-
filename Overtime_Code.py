import pandas as pd
import os
import platform
import subprocess

# ----------------------------
# Utility Function: Open File
# ----------------------------
def open_file(filepath: str):
    """Opens a file using the default application."""
    system = platform.system()
    try:
        if system == "Windows":
            os.startfile(filepath)
        elif system == "Darwin":
            subprocess.call(["open", filepath])
        else:
            subprocess.call(["xdg-open", filepath])
    except Exception as e:
        print(f"Unable to open {filepath}: {e}")

# ----------------------------
# Data Loading Function
# ----------------------------
def load_data(file_path: str) -> pd.DataFrame:
    """
    Loads the Excel file and reorders the columns to:
    Date, EmpID, Hours, Pay Code, Location, Regular Hours.
    Converts Date to datetime and Hours and Regular Hours to numeric.
    Standardizes Pay Code.
    """
    df = pd.read_excel(file_path)
    desired_order = ["Date", "EmpID", "Hours", "Pay Code", "Location", "Regular Hours"]
    df = df[desired_order]
    if pd.api.types.is_numeric_dtype(df['Date']):
        df['Date'] = pd.to_datetime(df['Date'], unit='D', origin='1899-12-30')
    else:
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df['Hours'] = pd.to_numeric(df['Hours'], errors='coerce')
    df['Regular Hours'] = pd.to_numeric(df['Regular Hours'], errors='coerce')
    df['Pay Code'] = df['Pay Code'].str.lower().str.strip()
    return df

# ----------------------------
# Timecard Report Function
# ----------------------------
def generate_timecard_report(df: pd.DataFrame, output_file: str = "timecard_report.csv", week_freq: str = 'W-SAT'):
    """
    Generates a calendar-style timecard report with one row per employee-week.
    
    For each day, a string is produced in the format:
       "Mon 03/14: 8.5 hrs (Reg: X, OT Paid: Y, Proj OT: Z)"
    
    where:
      - Total Hours is the sum of all hours for that day.
      - Reg is computed by summing hours from rows with pay code "regular."
      - OT Paid is computed by summing hours from rows with pay code "overtime."
      - Sched_Reg is taken directly from the "Regular Hours" column in the data file 
        (using the average for that week).
      - Proj OT for the day = Total Hours - Sched_Reg.
    
    Weekly aggregates are computed as follows:
      - Weekly Regular Hours: Sum of hours for rows with pay code "regular" for the week.
      - Weekly OT Paid: Sum of hours for rows with pay code "overtime" for the week.
      - Weekly Total Hours = Weekly Regular Hours + Weekly OT Paid.
      - Weekly Proj OT = Sum of daily Proj OT values.
      - Overtime Owed is defined as:
            • If Weekly Total Hours ≤ 40, then = Weekly Proj OT.
            • If Weekly Total Hours > 40, then:
                  if (Weekly OT Paid - Weekly Proj OT) > 0, then = (Weekly OT Paid - Weekly Proj OT), else 0.
    """
    # Compute the average scheduled regular hours for each week.
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df['Week Start'] = df['Date'].dt.to_period(week_freq).apply(lambda r: r.start_time)
    weekly_avg_reg = df.groupby(['Location','EmpID','Week Start'])['Regular Hours'].mean()\
                        .reset_index().rename(columns={'Regular Hours': 'Sched_Reg'})
    
    # --- Daily Calculations ---
    # Total Hours for each day.
    daily_total = df.groupby(['Location','EmpID','Week Start','Date'])['Hours']\
                    .sum().reset_index().rename(columns={'Hours': 'Total_Hours'})
    
    # Reg: Sum of hours for rows with pay code "regular".
    daily_reg = df[df['Pay Code'] == 'regular'].groupby(['Location','EmpID','Week Start','Date'])['Hours']\
                  .sum().reset_index().rename(columns={'Hours': 'Reg'})
    
    # OT Paid: Sum of hours for rows with pay code "overtime".
    daily_ot = df[df['Pay Code'] == 'overtime'].groupby(['Location','EmpID','Week Start','Date'])['Hours']\
                 .sum().reset_index().rename(columns={'Hours': 'OT_Paid'})
    
    # Merge daily aggregates.
    daily = pd.merge(daily_total, daily_reg, on=['Location','EmpID','Week Start','Date'], how='left')
    daily = pd.merge(daily, daily_ot, on=['Location','EmpID','Week Start','Date'], how='left')
    daily['Reg'] = daily['Reg'].fillna(0)
    daily['OT_Paid'] = daily['OT_Paid'].fillna(0)
    
    # Merge in the weekly average scheduled regular hours for that week.
    daily = pd.merge(daily, weekly_avg_reg, on=['Location','EmpID','Week Start'], how='left')
    
    # Compute Proj OT for each day.
    daily['Proj_OT'] = daily.apply(lambda row: max(0, row['Total_Hours'] - row['Sched_Reg']), axis=1)
    
    # Construct the daily info string.
    daily['Day Info'] = daily['Date'].dt.strftime('%a %m/%d') + ": " + daily['Total_Hours'].astype(str) + " hrs (Reg: " \
                          + daily['Reg'].astype(str) + ", OT Paid: " + daily['OT_Paid'].astype(str) + ", Proj OT: " \
                          + daily['Proj_OT'].astype(str) + ")"
    
    # Pivot the daily info so that each day becomes its own column.
    pivot = daily.pivot_table(index=['Location','EmpID','Week Start'],
                              columns=daily['Date'].dt.day_name(),
                              values='Day Info', aggfunc='first')
    pivot.reset_index(inplace=True)
    day_order = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday']
    for day in day_order:
        if day not in pivot.columns:
            pivot[day] = ""
    
    # --- Weekly Aggregates ---
    # Weekly Regular Hours: Sum of hours for rows with pay code "regular" for the week.
    reg_week = df[df['Pay Code'] == 'regular'].groupby(['Location','EmpID','Week Start'])['Hours']\
                .sum().reset_index().rename(columns={'Hours': 'Weekly Regular Hours'})
    
    # Weekly OT Paid: Sum of hours for rows with pay code "overtime" for the week.
    ot_week = df[df['Pay Code'] == 'overtime'].groupby(['Location','EmpID','Week Start'])['Hours']\
               .sum().reset_index().rename(columns={'Hours': 'Weekly OT Paid'})
    
    weekly_agg = pd.merge(reg_week, ot_week, on=['Location','EmpID','Week Start'], how='outer')
    weekly_agg['Weekly Regular Hours'] = weekly_agg['Weekly Regular Hours'].fillna(0)
    weekly_agg['Weekly OT Paid'] = weekly_agg['Weekly OT Paid'].fillna(0)
    weekly_agg['Weekly Total Hours'] = weekly_agg['Weekly Regular Hours'] + weekly_agg['Weekly OT Paid']
    
    # Weekly Proj OT: Sum of daily Proj OT.
    weekly_proj_ot = daily.groupby(['Location','EmpID','Week Start'])['Proj_OT']\
                          .sum().reset_index().rename(columns={'Proj_OT': 'Weekly Proj OT'})
    weekly_agg = pd.merge(weekly_agg, weekly_proj_ot, on=['Location','EmpID','Week Start'], how='left')
    weekly_agg['Weekly Proj OT'] = weekly_agg['Weekly Proj OT'].fillna(0)
    
    # Compute Overtime Owed.
    # If Weekly Total Hours <= 40, then Overtime Owed = Weekly Proj OT.
    # If Weekly Total Hours > 40:
    #    if (Weekly OT Paid - Weekly Proj OT) > 0, then Overtime Owed = (Weekly OT Paid - Weekly Proj OT), else 0.
    def compute_ot_owed(row):
        if row['Weekly Total Hours'] <= 40:
            value = row['Weekly Proj OT']
        else:
            diff = row['Weekly OT Paid'] - row['Weekly Proj OT']
            value = 0 if diff > 0 else abs(diff)
        return value
    
    weekly_agg['Overtime Owed'] = weekly_agg.apply(compute_ot_owed, axis=1)
    
     
    # Merge the daily pivot with weekly aggregates.
    report = pd.merge(pivot, weekly_agg, on=['Location','EmpID','Week Start'], how='left')
    
    final_cols = ['EmpID', 'Location', 'Week Start'] + day_order + \
                 ['Weekly Regular Hours', 'Weekly OT Paid', 'Weekly Total Hours', 'Weekly Proj OT', 'Overtime Owed']
    report = report[final_cols]
    
    report.to_csv(output_file, index=False)
    return output_file

# ----------------------------
# Main Script
# ----------------------------
def main():
    file_path = "/Users/lindseybaranauskas/Documents/Work/ALL_PCs_Audit_revising.xlsx"
    df = load_data(file_path)
    timecard_file = generate_timecard_report(df)
    print("Timecard report output to:", timecard_file)
    open_file(timecard_file)

if __name__ == "__main__":
    main()
