import xlwings as xw
import fileTaoSheetExcel
import fileTinhToanPhan1
import fileTinhToanPhan2
import fileTinhToanPhan3

congty = ["FPT", "HPG", "BMP", "GAS", "MWG", "PTB", "IMP", "MSN", "VCB", "VNM"]
file_name = ""
ten_phan1 = "Cucaltor"
ten_phan2 = "DataSolver"
ten_phan3 = "TheOptimalRiskyPorfolio"
ten_phan4 = "TheEfficidentFrontierAndCAL"

if __name__ == "__main__":
    while True:
        file_name = input("Nhập tên file: ")
        file_name = file_name + ".xlsx"
        try:
            wb = xw.Book(file_name)
            break
        except Exception as e:
            print("Lỗi không tìm thấy file")
    while True:
        check = input("Nhập định dạng số thập phân trong Excel của bạn, (Nhập . / ,): ")
        if check == '.':
            fileTaoSheetExcel.str1 = "0.0000%"
            break
        if check == ',':
            fileTaoSheetExcel.str1 = "0,0000%"
            break
    rf = float(input("Nhập rf = "))
    fileTinhToanPhan2.rf = rf / 100
    A = int(input("Nhập A = "))
    fileTinhToanPhan3.A = A
    
    print("Đang tính toán. Vui lòng đợi")
    for sheet in wb.sheets:
        if sheet.name == ten_phan1 or sheet.name == ten_phan2 or sheet.name == ten_phan3:
            sheet.delete()
            
    fileTinhToanPhan1.main(wb, ten_phan1, congty)
    for sheet in wb.sheets:
        if sheet.name == ten_phan1:
            sheet1 = sheet
    dataSharpeMax = fileTinhToanPhan2.main(wb, sheet1 ,ten_phan2, congty)
    for sheet in wb.sheets:
        if sheet.name == ten_phan2:
            sheet2 = sheet
    fileTinhToanPhan3.main(wb, sheet2 ,ten_phan3, dataSharpeMax)
    print("Hoàn tất")
    wb.save()
