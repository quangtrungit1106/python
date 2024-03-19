import xlwings as xw
import fileTaoSheetExcel
def main(wb, sheet2, ten_phan3, dataSharpeMax):
    # Tổ hợp chập 3 của các công ty
    
    fileTaoSheetExcel.TaoSheet(wb, ten_phan3)
    for sh in wb.sheets:
        if sh.name == ten_phan3:
            sheet = sh
    fileTaoSheetExcel.GiaoDienSheet3(sheet)

    
    data_range = sheet.used_range
    # Tự động điều chỉnh chiều cao của hàng
    data_range.rows.autofit()
    # Tự động điều chỉnh chiều rộng của cột
    data_range.columns.autofit()
