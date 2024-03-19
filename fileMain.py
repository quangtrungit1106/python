import xlwings as xw
import fileTinhToanPhan1
import fileTinhToanPhan2
import fileTinhToanPhan3

congty = ["FPT", "HPG", "BMP", "GAS", "MWG", "PTB", "IMP", "MSN", "VCB", "VNM"]
file_name = "data.xlsx"
ten_phandata = "Data"
ten_phan1 = "Cucaltor"
ten_phan2 = "DataSolver"
ten_phan3 = "TheOptimalRiskyPorfolio"
ten_phan4 = "TheEfficidentFrontierAndCAL"

if __name__ == "__main__":
    wb = xw.Book(file_name)
    for sheet in wb.sheets:
        if sheet.name != ten_phandata:
            sheet.delete()
    fileTinhToanPhan1.main(wb, ten_phan1, congty)
    for sheet in wb.sheets:
        if sheet.name == ten_phan1:
            sheet1 = sheet
    dataSharpeMax = fileTinhToanPhan2.main(wb, sheet1 ,ten_phan2, congty)
    '''for sheet in wb.sheets:
        if sheet.name == ten_phan2:
            sheet2 = sheet
    fileTinhToanPhan3.main(wb, sheet2 ,ten_phan3, dataSharpeMax)
    '''
    print("Hoan tat")
    wb.save()
