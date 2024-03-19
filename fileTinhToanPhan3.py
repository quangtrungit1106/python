import xlwings as xw
import fileTaoSheetExcel
import math

A = 0

def in_thong_tin(sheet, data):
    sheet.range("A8:T8").value = data
    sheet.range("A8:A8").value = 1
    sheet.range("D12:D12").value = A
    sheet.range("G13:G13").value = 1
    sheet.range("K11:K11").value = 100
    sheet.range("I14:I14").value = data[1]
    sheet.range("I15:I15").value = data[2]
    sheet.range("I16:I16").value = data[3]
 
def tinh_toan(sheet, data):
    y = (data[17] - data[4]) / (A*data[18]*data[18])
    e_rc = data[4] + y*(data[17] - data[18])
    o_c = math.sqrt(y*y * data[18]*data[18])
    U = e_rc - 0.5*A*o_c*o_c
    #in ra file
    sheet.range("G11:G11").value = y
    sheet.range("G12:G12").value = 1-y
    sheet.range("G15:G15").value = e_rc
    sheet.range("G16:G16").value = o_c
    sheet.range("G17:G17").value = U
    sheet.range("K12:K12").value = 100 - y*100
    sheet.range("K13:K13").value = y*100
    sheet.range("K14:K14").value = y*100*data[14]
    sheet.range("K15:K15").value = y*100*data[15]
    sheet.range("K16:K16").value = y*100*data[16]
    
def main(wb, sheet2, ten_phan3, dataSharpeMax):
    # Tổ hợp chập 3 của các công ty
    
    fileTaoSheetExcel.TaoSheet(wb, ten_phan3)
    for sh in wb.sheets:
        if sh.name == ten_phan3:
            sheet = sh
    fileTaoSheetExcel.GiaoDienSheet3(sheet)
    in_thong_tin(sheet, dataSharpeMax)
    tinh_toan(sheet, dataSharpeMax)
    data_range = sheet.used_range
    # Tự động điều chỉnh chiều cao của hàng
    data_range.rows.autofit()
    # Tự động điều chỉnh chiều rộng của cột
    data_range.columns.autofit()
