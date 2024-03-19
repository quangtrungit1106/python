import xlwings as xw
import itertools
import math
import fileTaoSheetExcel

pham_vi_data_dasapxep = "B7:L67"
pham_vi_data_return = "M8:V67"
pham_vi_data_rimean = "W8:AF67"
pham_vi_data_rimeanbinh = "AG8:AP67"
pham_vi_data_tong1 =  "M68:V70"

# Lấy dữ liệu nguồn 
def lay_data_nguon(wb, congty):
    sheet = wb.sheets[0]
    data_congty = {}
    data = sheet.range("A1:G700").value
    for cty in congty:
        data_cty = []
        for row in data:
            if row[0] == cty:
                data_cty.append(row)
        # Sắp xếp dữ liệu cho công ty hiện tại theo cột thứ hai (số thứ tự)
        data_cty.sort(key=lambda x: x[1])
        data_congty[cty] = data_cty
    return data_congty

# In thông tin
def in_thong_tin_nguon(data, new_sheet, congty):
    print("Đang đọc dữ liệu, Vui lòng đợi")
    for cty, data_cty in data.items():
        # Dùng để ghi vào file excel
        row = 7
        stt_dong = 1
        stt_cot = congty.index(cty) + 3
        for row_data in data_cty:
            return_moi = row_data[5]
            # Kiểm tra xem row_data[5] có phải là chuỗi không
            if isinstance(row_data[5], str):
                # Xóa dấu chấm ngăn cách phần ngàn
                return_moi = row_data[5].replace(".", "")
                # Nếu chuỗi vẫn chứa dấu chấm, loại bỏ tất cả dấu chấm còn lại
                if return_moi.count('.') > 1:
                    return_moi = return_moi.replace(".", "", return_moi.count('.') - 1)
            new_sheet.range(row, 1).value = stt_dong  # Cột STT
            new_sheet.range(row, 2).value = row_data[1]  # Cột Thời gian
            new_sheet.range(row, stt_cot).value = return_moi # Cột data cua cong ty
            row += 1
            stt_dong += 1
# Tính toán return và mean
def tinh_toan_return_mean(new_sheet):
    data = new_sheet.range(pham_vi_data_dasapxep).value
    print("Đang tính toán return và mean. Vui lòng đợi")
    # Dùng để ghi vào file excel
    row = 7
    
    for cot in range(1, 11):
        stt_cot = 12
        sum_return = 0
        for dong in range (1, len(data)):
            sum_return += (data[dong][cot] - data[dong-1][cot])/data[dong-1][cot]
            new_sheet.range(row+dong, stt_cot+cot).value = (data[dong][cot] - data[dong-1][cot])/data[dong-1][cot]
        new_sheet.range(row+61, stt_cot+cot).value = sum_return / 60    # mean
    
    print("Đã xong tính toán return")

# Tính toán rimean
def tinh_toan_rimean_var(new_sheet):
    data = new_sheet.range(pham_vi_data_return).value
    data_tong = new_sheet.range(pham_vi_data_tong1).value
    print("Đang tính toán ri-mean, Vui lòng đợi")
    # Dùng để ghi vào file excel
    row = 8
    
    for cot in range(0, 10):
        stt_cot = 13
        sum_return = 0
        for dong in range (0, len(data)):
            rimean = data[dong][cot] - data_tong[0][cot]
            sum_return += rimean * rimean
            new_sheet.range(row+dong, stt_cot+10+cot).value = rimean # ri-mean
            new_sheet.range(row+dong, stt_cot+20+cot).value = rimean * rimean # ri-mean binh
        new_sheet.range(row+dong+1, stt_cot+20+cot).value = sum_return # tong
        new_sheet.range(row+dong+2, stt_cot+cot).value = sum_return / 60 #var
        new_sheet.range(row+dong+3, stt_cot+cot).value = math.sqrt(sum_return / 60)
    print("Đã xong tính toán return") 

# Tính toán cov cor
def tinh_toan_cov_cor(new_sheet, tohop_chap3):
    print("Đang tính toán ri-mean, Vui lòng đợi")
    data = new_sheet.range(pham_vi_data_rimean).value
    data_tong = new_sheet.range(pham_vi_data_tong1).value
    
    tong = 0
    row = 76
    for x in range(0, 10):
        col = 13
        for y in range (0, 10):
            tong = 0
            cov = 0
            cor = 0
            for row_data in data:
                tong += row_data[x] * row_data[y]
            cov = tong / 60
            cor = cov / data_tong[2][x] / data_tong[2][y]
            new_sheet.range(row, col + y).value = cov
            new_sheet.range(row+15, col + y).value = cor
        row += 1
    
            
    print("Đã xong tính toán cov, cor")
def main(wb, ten_phan1, congty):
    # Tổ hợp chập 3 của các công ty
    tohop_chap3 = list(itertools.combinations(congty, 3))
    fileTaoSheetExcel.TaoSheet(wb, ten_phan1)
    for sh in wb.sheets:
        if sh.name == ten_phan1:
            sheet = sh
    fileTaoSheetExcel.GiaoDienSheet1(sheet, congty)
    data_nguon = lay_data_nguon(wb, congty)
    in_thong_tin_nguon(data_nguon, sheet, congty)
    tinh_toan_return_mean(sheet)
    tinh_toan_rimean_var(sheet)
    tinh_toan_cov_cor(sheet, tohop_chap3)
    data_range = sheet.used_range
    # Tự động điều chỉnh chiều cao của hàng
    data_range.rows.autofit()
    # Tự động điều chỉnh chiều rộng của cột
    data_range.columns.autofit()
