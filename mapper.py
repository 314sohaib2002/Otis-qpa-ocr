import openpyxl
import pandas as pd

def clean_val(val):
    if val is None or val == "-":
        return ""
    return str(val)

def process_timesheet(input_excel_path, df_admin, df_absence, df_overtime):
    # Load the workbook (data_only=True ensures we get the value, not formulas)
    wb = openpyxl.load_workbook(input_excel_path, data_only=True)
    sheet = wb.active

    # ==========================================
    # 1. PROCESS ADMIN TEMPLATE
    # ==========================================
    employee_id = clean_val(sheet["F2"].value)
    employee_name = clean_val(sheet["B2"].value)
    
    total_overtime = clean_val(sheet["F36"].value)

    # For G36/H2+I2 logic
    raw_overtime_payout = sheet["G36"].value 
    if raw_overtime_payout is None or raw_overtime_payout == "-":
        try:
            h2 = sheet["H2"].value if isinstance(sheet["H2"].value, (int, float)) else 0
            i2 = sheet["I2"].value if isinstance(sheet["I2"].value, (int, float)) else 0
            overtime_payout = str(h2 + i2)
        except:
            overtime_payout = "0"
    else:
        overtime_payout = clean_val(raw_overtime_payout)

    overtime_50 = clean_val(sheet["H36"].value)
    overtime_100 = clean_val(sheet["I36"].value)

    # Time balance transferred
    time_balance = clean_val(sheet["E40"].value)

    # Vacation days and remaining
    days_taken = clean_val(sheet["K36"].value)
    remaining_holiday = clean_val(sheet["K41"].value)

    # Timebank logic
    e5_val = sheet["E5"].value if isinstance(sheet["E5"].value, (int, float)) else 0
    e40_val = sheet["E40"].value if isinstance(sheet["E40"].value, (int, float)) else 0

    timebank_math_result = str(e5_val - e40_val)
    timebank_payout = timebank_math_result
    time_off_withdrawal = timebank_math_result

    overtime_for_time_off = clean_val(sheet["E38"].value)

    # Advanced Parsing for Dates (Rows 4 through 34)
    vacation_dates = []
    self_report_dates = []
    sick_leave_dates = []

    for row in range(4, 35):
        day_val = sheet[f"A{row}"].value
        if not isinstance(day_val, (int, float)):
            continue
            
        day_int = int(day_val)
        
        # Check Vacation (Column K)
        k_val = sheet[f"K{row}"].value
        if isinstance(k_val, (int, float)) and k_val > 0:
            vacation_dates.append(str(day_int))
            
        # Check Sick Leave / Absence (Column J) and Comments (Column D)
        j_val = sheet[f"J{row}"].value
        d_val = sheet[f"D{row}"].value
        
        if isinstance(j_val, (int, float)) and j_val > 0:
            if d_val and str(d_val).strip() != "":
                sick_leave_dates.append(str(day_int))
            else:
                self_report_dates.append(str(day_int))

    vacation_dates_str = ", ".join(vacation_dates)
    self_report_dates_str = ", ".join(self_report_dates)
    sick_leave_dates_str = ", ".join(sick_leave_dates)

    cols = df_admin.columns
    
    new_admin_row = {
        cols[0]: employee_id,
        cols[1]: employee_name,
        cols[5]: total_overtime,
        cols[6]: overtime_payout,
        cols[7]: overtime_50,
        cols[8]: overtime_100,
        cols[9]: overtime_for_time_off,
        cols[10]: time_off_withdrawal,
        cols[11]: timebank_payout,
        cols[12]: time_balance,
        cols[13]: vacation_dates_str,
        cols[14]: days_taken,
        cols[15]: remaining_holiday,
        cols[16]: self_report_dates_str,
        cols[17]: sick_leave_dates_str
    }
    df_admin = pd.concat([df_admin, pd.DataFrame([new_admin_row])], ignore_index=True)


    # ==========================================
    # 2. PROCESS ABSENCE TEMPLATE
    # ==========================================
    def group_consecutive_days(days_list):
        if not days_list:
            return []
        days_list = sorted([int(d) for d in days_list])
        groups = []
        start_day = days_list[0]
        prev_day = days_list[0]
        count = 1
        for day in days_list[1:]:
            if day == prev_day + 1:
                count += 1
                prev_day = day
            else:
                groups.append({"from": start_day, "to": prev_day, "days": count})
                start_day = day
                prev_day = day
                count = 1
        groups.append({"from": start_day, "to": prev_day, "days": count})
        return groups

    vacation_groups = group_consecutive_days(vacation_dates)
    self_report_groups = group_consecutive_days(self_report_dates)
    cert_sick_groups = group_consecutive_days(sick_leave_dates)

    absence_rows = []
    month_str = clean_val(sheet["I2"].value) # e.g. "Jun"
    year_str = clean_val(sheet["J2"].value)  # e.g. "2026"

    def format_date(day):
        return f"{day}-{month_str}-{year_str}"

    for grp in vacation_groups:
        row_data = [None] * len(df_absence.columns)
        row_data[1] = employee_id
        row_data[2] = employee_name
        row_data[4] = "Ferie (Vacation)"
        row_data[5] = "100%"
        row_data[6] = format_date(grp["from"])
        row_data[7] = format_date(grp["to"])
        row_data[8] = grp["days"]
        absence_rows.append(row_data)

    for grp in self_report_groups:
        row_data = [None] * len(df_absence.columns)
        row_data[1] = employee_id
        row_data[2] = employee_name
        row_data[4] = "2"
        row_data[5] = "100%"
        row_data[6] = format_date(grp["from"])
        row_data[7] = format_date(grp["to"])
        row_data[8] = grp["days"]
        absence_rows.append(row_data)

    for grp in cert_sick_groups:
        row_data = [None] * len(df_absence.columns)
        row_data[1] = employee_id
        row_data[2] = employee_name
        row_data[4] = "Syk med legeerklæring"
        row_data[5] = "100%"
        row_data[6] = format_date(grp["from"])
        row_data[7] = format_date(grp["to"])
        row_data[8] = grp["days"]
        absence_rows.append(row_data)

    if absence_rows:
        new_abs_df = pd.DataFrame(absence_rows, columns=df_absence.columns)
        df_absence = pd.concat([df_absence, new_abs_df], ignore_index=True)


    # ==========================================
    # 3. PROCESS OVERTIME TEMPLATE
    # ==========================================
    admin_cols = df_admin.columns
    header_118 = admin_cols[6]
    header_425 = admin_cols[7]
    header_426 = admin_cols[8]
    header_61  = admin_cols[9]

    val_118 = float(overtime_payout) if overtime_payout and str(overtime_payout).replace('.','').isdigit() else 0
    val_425 = float(overtime_50) if overtime_50 and str(overtime_50).replace('.','').isdigit() else 0
    val_426 = float(overtime_100) if overtime_100 and str(overtime_100).replace('.','').isdigit() else 0
    val_61  = float(overtime_for_time_off) if overtime_for_time_off and str(overtime_for_time_off).replace('.','').isdigit() else 0

    wage_entries = [
        (118, header_118, val_118),
        (425, header_425, val_425),
        (426, header_426, val_426),
        (61,  header_61,  val_61)
    ]

    overtime_rows = []
    columns_length = len(df_overtime.columns)

    for code, header_text, value in wage_entries:
        if value != 0: # Only add if value > 0
            row_data = [None] * columns_length
            row_data[0] = employee_id         
            row_data[1] = employee_name       
            row_data[2] = code                
            row_data[3] = header_text         
            row_data[4] = header_text         
            row_data[5] = value               
            overtime_rows.append(row_data)

    if overtime_rows:
        new_ovt_df = pd.DataFrame(overtime_rows, columns=df_overtime.columns)
        df_overtime = pd.concat([df_overtime, new_ovt_df], ignore_index=True)

    return df_admin, df_absence, df_overtime
