import xlwings as xw
import itertools
import math
import scipy.optimize as optimize
import fileTaoSheetExcel
rf = 0
def in_thong_tin(sheet, sheet1, congty, tohop_chap3, wb):
    
    data_rimean = sheet1.range("AG68:AP68").value
    data_return = sheet1.range("M68:V70").value
    data_cov = sheet1.range("M76:V85").value
    row = 8
    stt_dong = 1
    for tohop in tohop_chap3:
        cov = 0
        cor = 0
        sheet.range(row, 1).value = stt_dong  # Cột 1
        sheet.range(row, 2).value = tohop[0]  # Cột 2
        sheet.range(row, 3).value = tohop[1]  # Cột 3
        sheet.range(row, 4).value = tohop[2]  # Cột 4
        sheet.range(row, 5).value = rf / 12#Cột 5
        stt1 = 0
        for x in range(0, 3):
            stt = congty.index(tohop[x])
            sheet.range(row, 6+x).value = data_return[0][stt]  # Cột 6 7 8
            sheet.range(row, 12+x).value = data_return[2][stt]   # Cột 12 13 14
            for y in range(x+1, 3):
                if x != y:
                    i = congty.index(tohop[x])
                    j = congty.index(tohop[y])
                    sheet.range(row, 9+stt1).value = data_cov[i][j] # Cột 9 10 11 
                    stt1+=1
        row +=1
        stt_dong += 1
# Tính toán Output

def tinh_sharpe(x, row_data):
    w1 = x[0]
    w2 = x[1]
    w3 = x[2]
    e_rp = w1*row_data[5]+w2*row_data[6]+w3*row_data[7]
    o_p = math.sqrt((w1*w1)*(row_data[11]*row_data[11])
                        + (w2*w2)*(row_data[12]*row_data[12])
                        + (w3*w3)*(row_data[13]*row_data[13])
                        + 2*w1*w2*row_data[8]
                        + 2*w1*w3*row_data[9]
                        + 2*w3*w2*row_data[10])
    sharpe = (e_rp-row_data[4])/o_p
    return -sharpe

def dieu_kien_1(x):
    return x[0] + x[1] + x[2] - 1  # w1 + w2 + w3 = 1

def dieu_kien_2(x):
    min_percentage = 0.05  # 5%
    return min(x) - min_percentage #

def solver(row_data):
    initial_guess = [0.5, 0.3, 0.2]
    constraints = [{'type': 'eq', 'fun': dieu_kien_1},
                   {'type': 'ineq', 'fun': dieu_kien_2}]
    result = optimize.minimize(tinh_sharpe, initial_guess,args=(row_data,), constraints=constraints)
    w1, w2, w3 = result.x
    return -result.fun, w1, w2, w3

def tinh_toan_output(sheet, congty):
    data = sheet.range("A8:T127").value
    row = 8
    max_sharpe = 0
    max_row = 8
    
    for row_data in data:
        
        sharpe, w1, w2, w3 = solver(row_data)
        e_rp = w1*row_data[5]+w2*row_data[6]+w3*row_data[7]
        o_p = math.sqrt((w1*w1)*(row_data[11]*row_data[11])
                        + (w2*w2)*(row_data[12]*row_data[12])
                        + (w3*w3)*(row_data[13]*row_data[13])
                        + 2*w1*w2*row_data[8]
                        + 2*w1*w3*row_data[9]
                        + 2*w3*w2*row_data[10])
        if (sharpe > max_sharpe):
            max_sharpe = sharpe
            max_row = row
            
        # Ghi file
        sheet.range(row, 15).value = w1
        sheet.range(row, 16).value = w2
        sheet.range(row, 17).value = w3
        sheet.range(row, 18).value = e_rp
        sheet.range(row, 19).value = o_p
        sheet.range(row, 20).value = sharpe
        row += 1
    pham_vi = str("A" +str(max_row) +":T" +str(max_row))
    sheet.range(pham_vi).color = (255, 217, 102)
    max_rowdata = sheet.range(pham_vi).value
    sheet.range(pham_vi).api.Font.Bold = True
    return max_rowdata

def main(wb, sheet1, ten_phan2, congty):
    # Tổ hợp chập 3 của các công ty
    tohop_chap3 = list(itertools.combinations(congty, 3))
    fileTaoSheetExcel.TaoSheet(wb, ten_phan2)
    for sh in wb.sheets:
        if sh.name == ten_phan2:
            sheet = sh
    fileTaoSheetExcel.GiaoDienSheet2(sheet, congty)
    in_thong_tin(sheet, sheet1, congty, tohop_chap3, wb)
    data = tinh_toan_output(sheet, congty)
    data_range = sheet.used_range
    # Tự động điều chỉnh chiều cao của hàng
    data_range.rows.autofit()
    # Tự động điều chỉnh chiều rộng của cột
    data_range.columns.autofit()

    return data
