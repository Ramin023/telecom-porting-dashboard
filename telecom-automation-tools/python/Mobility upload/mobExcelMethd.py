from datetime import date, datetime
from openpyxl.styles import Font, Border, Side
import os

class mobExcelMethd:
    def saveEx(path, fname, uname):
        save_path = f"{path}\\Mobility_{fname}_{uname}_{date.today().strftime('%Y%m%d')}.xlsx"
        filename, extension = os.path.splitext(save_path)
        counter = 2

        while os.path.exists(save_path):
            save_path = f"{filename}_v{counter}{extension}"
            counter += 1
        return save_path


    def convertDateToStr(date):
        if isinstance(date, str):
            date_parts = date.strip().split('/')
            if len(date_parts) == 2:
                date_obj = datetime(datetime.now().year, int(date_parts[0]), int(date_parts[1]))
            elif len(date_parts) == 3:
                date_obj = datetime(int(date_parts[2])%2000+2000, int(date_parts[0]), int(date_parts[1]))
            else:
                return "Invalid Date"
        else:
            date_obj = date

        return date_obj.strftime("%m/%d/%Y")


    def font_header(wb, header):
        wb.append(header)
        for row in wb.iter_rows(min_row=1, max_row=1):
            for cell in row:
                cell.font = Font(bold=True)
                cell.border = Border(left=Side(style='thin'),
                                     right=Side(style='thin'),
                                     top=Side(style='thin'),
                                     bottom=Side(style='thin'))