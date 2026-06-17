import gspread
from oauth2client.service_account import ServiceAccountCredentials
import openpyxl
from datetime import datetime
from copy import copy
from openpyxl.styles import Alignment

def generate_excel_report(target_month, target_year):
    print(f"⏳ กำลังดึงข้อมูลและแยกประเภทน้ำมัน ประจำเดือน {target_month}/{target_year}...")
    
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
    client = gspread.authorize(creds)
    
    sheet_name = "Master Database ระบบจัดการน้ำมัน (LINE Bot)" 
    sheet = client.open(sheet_name).sheet1
    all_data = sheet.get_all_values()
    
    fuel_data = {
        "ดีเซล": {"records": [], "open_bal": 0.0, "tot_in": 0.0, "tot_out": 0.0},
        "เบนซิน": {"records": [], "open_bal": 0.0, "tot_in": 0.0, "tot_out": 0.0}
    }
    
    for row in all_data[1:]:
        if len(row) < 10: continue
        
        date_str = str(row[1]).split(" ")[0].strip()
        if not date_str: continue
        
        try:
            if "/" in date_str:
                parts = date_str.split("/")
                date_val = datetime(int(parts[2]), int(parts[1]), int(parts[0]))
            elif "-" in date_str:
                parts = date_str.split("-")
                date_val = datetime(int(parts[0]), int(parts[1]), int(parts[2]))
            else:
                continue
        except Exception:
            continue
            
        row_year = date_val.year
        if row_year > 2500: row_year -= 543
        row_month = date_val.month
        
        type_val = str(row[2]).strip()
        
        target_fuel = None
        if "ดีเซล" in type_val:
            target_fuel = "ดีเซล"
        elif "เบนซิน" in type_val or "แก๊สโซฮอล์" in type_val:
            target_fuel = "เบนซิน"
            
        if not target_fuel: continue
        
        # 🌟 ถอดโค้ดแปลงภาษาไทยออก ดึงข้อความมาตรงๆ จาก Sheet เลยครับ
        license_plate = str(row[3]).strip() if str(row[3]).strip() else "-"
        location = str(row[4]).strip() if str(row[4]).strip() else "-"
        mission = str(row[5]).strip() if str(row[5]).strip() else "-"
        
        in_str = str(row[6]).replace(',', '').strip()
        out_str = str(row[7]).replace(',', '').strip()
        in_amount = float(in_str) if in_str else 0.0
        out_amount = float(out_str) if out_str else 0.0
        
        price_str = str(row[8]).replace(',', '').strip()
        price = float(price_str) if price_str else "-"
        
        note = str(row[9]).strip() if len(row) > 9 and str(row[9]).strip() else ""
        
        if row_year < target_year or (row_year == target_year and row_month < target_month):
            fuel_data[target_fuel]["open_bal"] += (in_amount - out_amount)
            
        elif row_year == target_year and row_month == target_month:
            fuel_data[target_fuel]["tot_in"] += in_amount
            fuel_data[target_fuel]["tot_out"] += out_amount
            
            thai_year = target_year + 543
            thai_yy_2digits = str(thai_year)[-2:]
            formatted_date = f"{date_val.day}/{date_val.month}/{thai_yy_2digits}"
            
            fuel_data[target_fuel]["records"].append({
                'date': formatted_date, 'license': license_plate, 'location': location,
                'mission': mission, 'price': price, 'in_amount': in_amount,
                'out_amount': out_amount, 'note': note
            })
            
    try:
        wb = openpyxl.load_workbook("template_oil.xlsx")
    except FileNotFoundError:
        print("❌ ไม่พบไฟล์ template_oil.xlsx ในโฟลเดอร์ครับ")
        return
        
    print("✅ โหลด Template สำเร็จ! กำลังจัดรูปแบบและล็อกการจัดหน้ากระดาษ...")
    
    center_align = Alignment(horizontal='center', vertical='center')
    left_align = Alignment(horizontal='left', vertical='center')
    right_align = Alignment(horizontal='right', vertical='center')
    
    def fill_sheet(ws, data, month, year):
        thai_months = ["", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", 
                       "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        thai_month_name = thai_months[month]
        thai_year = year + 543
        
        try:
            ws['B4'].value = f"ประจำเดือน {thai_month_name} {thai_year}"
        except AttributeError:
            pass
        
        start_row = 7
        current_balance = round(data["open_bal"], 2)
        
        ws.cell(row=start_row, column=3, value="ยอดยกมาจากเดือนก่อน").alignment = right_align
        c_ob_val = ws.cell(row=start_row, column=10, value=current_balance)
        c_ob_val.number_format = '#,##0.00'
        c_ob_val.alignment = right_align
        
        current_row = start_row + 1
        idx = 1
        
        template_styles = {}
        for c in range(1, 11):
            t_cell = ws.cell(row=8, column=c)
            template_styles[c] = {
                'font': copy(t_cell.font),
                'border': copy(t_cell.border)
            }

        for r in data["records"]:
            in_amt = r['in_amount']
            out_amt = r['out_amount']
            
            current_balance = current_balance + in_amt - out_amt
            current_balance = round(current_balance, 2)
            
            final_note = f"{current_balance:,.2f}"
            if r['note'] and r['note'] != "-": final_note += f" ({r['note']})"
                
            cells_to_write = [
                (1, idx), (2, r['date']), (3, r['license']), (4, r['location']),
                (5, r['mission']), (6, r['price']), (7, "-"),
                (8, in_amt if in_amt > 0 else "-"),
                (9, out_amt if out_amt > 0 else "-"),
                (10, final_note)
            ]
            
            for col, val in cells_to_write:
                c = ws.cell(row=current_row, column=col, value=val)
                
                style = template_styles[col]
                c.font = copy(style['font'])
                c.border = copy(style['border'])
                
                if col in [3, 5]: 
                    c.alignment = left_align
                elif col in [6, 8, 9, 10]: 
                    c.alignment = right_align
                else: 
                    c.alignment = center_align
                
                if col in [8, 9] and isinstance(val, float):
                    c.number_format = '#,##0.00'
            
            current_row += 1
            idx += 1
            
        ws.cell(row=current_row, column=3, value="รวมใช้ประจำเดือน").alignment = right_align
        c_tot_out = ws.cell(row=current_row, column=9, value=data["tot_out"] if data["tot_out"] > 0 else "-")
        c_tot_out.alignment = right_align
        if data["tot_out"] > 0: c_tot_out.number_format = '#,##0.00'
        
        ws.cell(row=current_row+1, column=3, value="รับเพิ่มประจำเดือน").alignment = right_align
        c_tot_in = ws.cell(row=current_row+1, column=8, value=data["tot_in"] if data["tot_in"] > 0 else "-")
        c_tot_in.alignment = right_align
        if data["tot_in"] > 0: c_tot_in.number_format = '#,##0.00'
        
        ws.cell(row=current_row+2, column=3, value="ยอดยกไปเดือนหน้า").alignment = right_align
        c_bal = ws.cell(row=current_row+2, column=10, value=current_balance)
        c_bal.alignment = right_align
        c_bal.number_format = '#,##0.00'
        
        for r_idx in range(current_row, current_row + 3):
            for c_idx in range(1, 11):
                c = ws.cell(row=r_idx, column=c_idx)
                if c_idx in template_styles:
                    style = template_styles[c_idx]
                    if not c.font: c.font = copy(style['font'])
                    if not c.border: c.border = copy(style['border'])

    if "ดีเซล" in wb.sheetnames:
        fill_sheet(wb["ดีเซล"], fuel_data["ดีเซล"], target_month, target_year)
    
    if "เบนซิน" in wb.sheetnames:
        fill_sheet(wb["เบนซิน"], fuel_data["เบนซิน"], target_month, target_year)

    output_filename = f"รายงานน้ำมัน_เดือน_{target_month}_{target_year}.xlsx"
    wb.save(output_filename)
    print(f"🎉 เสร็จเรียบร้อย! สร้างไฟล์รายงาน 2 ชีท ชื่อ '{output_filename}' ให้แล้วครับ!")

if __name__ == "__main__":
    now = datetime.now()
    report_month = now.month - 1
    report_year = now.year
    
    if report_month == 0:
        report_month = 12
        report_year -= 1
        
    print(f"📅 วันนี้คือเดือน {now.month}/{now.year}")
    print(f"🚀 กำลังเริ่มสร้างรายงานสรุปยอดของเดือน {report_month}/{report_year} ให้อัตโนมัติ...")
    
    generate_excel_report(report_month, report_year)