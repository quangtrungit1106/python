import xlwings as xw
import win32com.client
# Khai báo các hằng số từ thư viện win32com.client.constants
xlEdgeBottom = win32com.client.constants.xlEdgeBottom
xlEdgeLeft = win32com.client.constants.xlEdgeLeft
xlEdgeRight = win32com.client.constants.xlEdgeRight
xlEdgeTop = win32com.client.constants.xlEdgeTop
xlInsideHorizontal = win32com.client.constants.xlInsideHorizontal
xlInsideVertical = win32com.client.constants.xlInsideVertical

str1 = ""


def TaoSheet(wb, tensheet):
    for sheet in wb.sheets:
        if sheet.name == tensheet:
            sheet.delete()
            new_sheet = wb.sheets.add(tensheet, after=wb.sheets[-1])
            return
    new_sheet = wb.sheets.add(tensheet, after=wb.sheets[-1])

def GiaoDienSheet1(sheet, congty):
    #Bảng1: return, ri-mean, ri-mean bình
    bang1 = sheet.range("A5:AP67")
    sheet.range("A5:A6").api.Merge()
    sheet.range("A5:A6").value = "TT"
    sheet.range("B5:B6").api.Merge()
    sheet.range("B5:B6").value = "Date"
    sheet.range("C5:C6").api.Merge()
    sheet.range("D5:D6").api.Merge()
    sheet.range("E5:E6").api.Merge()
    sheet.range("F5:F6").api.Merge()
    sheet.range("G5:G6").api.Merge()
    sheet.range("H5:H6").api.Merge()
    sheet.range("I5:I6").api.Merge()
    sheet.range("J5:J6").api.Merge()
    sheet.range("K5:K6").api.Merge()
    sheet.range("L5:L6").api.Merge()
    sheet.range("C5:L5").value = congty
    sheet.range("A5:L6").color = (248, 203, 173) # cam
    sheet.range("A5:AP6").api.Font.Bold = True
    sheet.range("M5:AP5").color = (255, 217, 102) # vang
    sheet.range("M6:AP6").color = (198, 224, 180)   #xanh
    sheet.range("M6:V6").value = congty
    sheet.range("W6:AF6").value = congty
    sheet.range("AG6:AP6").value = congty
    sheet.range("M5:V5").api.Merge()
    sheet.range("M5:V5").value = "RETURN"
    sheet.range("W5:AF5").api.Merge()
    sheet.range("W5:AF5").value = "RI-MEAN"
    sheet.range("AG5:AP5").api.Merge()
    sheet.range("AG5:AP5").value = "RI-MEAN ^ 2"
    
    bang1.api.HorizontalAlignment = -4108
    bang1.api.VerticalAlignment = -4108
    bang1.api.Borders(7).LineStyle = 1  # Đường viền bên trái
    bang1.api.Borders(8).LineStyle = 1  # Đường viền trên cùng
    bang1.api.Borders(9).LineStyle = 1  # Đường viền dưới cùng
    bang1.api.Borders(10).LineStyle = 1  # Đường viền bên phải
    bang1.api.Borders(11).LineStyle = 1  # Đường viền nội bộ dọc
    bang1.api.Borders(12).LineStyle = 1  # Đường viền nội bộ ngang
    for row_index in range(7, 68, 2):  # Bắt đầu từ hàng thứ 2 và tăng 2 hàng mỗi lần
        for col_index in range(1, 43):
            sheet.cells(row_index, col_index).color = (189, 215, 238)
    
    # Bảng 2: Mean, Var, S.D
    bang2 = sheet.range("K68:V70")
    sheet.range("K68:L68").api.Merge()
    sheet.range("K68:L68").value = "Mean:"
    sheet.range("K69:L69").api.Merge()
    sheet.range("K69:L69").value = "Var:"
    sheet.range("K70:L70").api.Merge()
    sheet.range("K70:L70").value = "S.D:"
    sheet.range("K68:L70").api.Font.Bold = True

    bang2.api.HorizontalAlignment = -4108
    bang2.api.VerticalAlignment = -4108
    bang2.api.Borders(7).LineStyle = 1  # Đường viền bên trái
    bang2.api.Borders(8).LineStyle = 1  # Đường viền trên cùng
    bang2.api.Borders(9).LineStyle = 1  # Đường viền dưới cùng
    bang2.api.Borders(10).LineStyle = 1  # Đường viền bên phải
    bang2.api.Borders(11).LineStyle = 1  # Đường viền nội bộ dọc
    bang2.api.Borders(12).LineStyle = 1  # Đường viền nội bộ ngang

    # Bảng 3: Tổng ri-mean^2
    bang3 = sheet.range("AF68:AP68")
    sheet.range("AF68").api.Font.Bold = True
    sheet.range("AF68").value = "Tổng:"
    
    bang3.api.HorizontalAlignment = -4108
    bang3.api.VerticalAlignment = -4108
    bang3.api.Borders(7).LineStyle = 1  # Đường viền bên trái
    bang3.api.Borders(8).LineStyle = 1  # Đường viền trên cùng
    bang3.api.Borders(9).LineStyle = 1  # Đường viền dưới cùng
    bang3.api.Borders(10).LineStyle = 1  # Đường viền bên phải
    bang3.api.Borders(11).LineStyle = 1  # Đường viền nội bộ dọc
    bang3.api.Borders(12).LineStyle = 1  # Đường viền nội bộ ngang

    # Bảng 4: Bảng Cov
    bang4 = sheet.range("K75:V85")
    sheet.range("K75:V75").api.Font.Bold = True
    sheet.range("K75:V75").color = (248, 203, 173) # cam
    sheet.range("K75:L75").api.Merge()
    sheet.range("K75:L75").value = "COVERIANCE"
    sheet.range("K75:L75").color = (255, 217, 102)
    sheet.range("K76:L76").api.Merge()
    sheet.range("K76:L76").value = congty[0]
    sheet.range("K77:L77").api.Merge()
    sheet.range("K77:L77").value = congty[1]
    sheet.range("K78:L78").api.Merge()
    sheet.range("K78:L78").value = congty[2]
    sheet.range("K79:L79").api.Merge()
    sheet.range("K79:L79").value = congty[3]
    sheet.range("K80:L80").api.Merge()
    sheet.range("K80:L80").value = congty[4]
    sheet.range("K81:L81").api.Merge()
    sheet.range("K81:L81").value = congty[5]
    sheet.range("K82:L82").api.Merge()
    sheet.range("K82:L82").value = congty[6]
    sheet.range("K83:L83").api.Merge()
    sheet.range("K83:L83").value = congty[7]
    sheet.range("K84:L84").api.Merge()
    sheet.range("K84:L84").value = congty[8]
    sheet.range("K85:L85").api.Merge()
    sheet.range("K85:L85").value = congty[9]
    sheet.range("K76:L85").color = (248, 203, 173) # cam
    sheet.range("K75:L85").api.Font.Bold = True
    sheet.range("M75:V75").value = congty
       
    bang4.api.HorizontalAlignment = -4108
    bang4.api.VerticalAlignment = -4108
    bang4.api.Borders(7).LineStyle = 1  # Đường viền bên trái
    bang4.api.Borders(8).LineStyle = 1  # Đường viền trên cùng
    bang4.api.Borders(9).LineStyle = 1  # Đường viền dưới cùng
    bang4.api.Borders(10).LineStyle = 1  # Đường viền bên phải
    bang4.api.Borders(11).LineStyle = 1  # Đường viền nội bộ dọc
    bang4.api.Borders(12).LineStyle = 1  # Đường viền nội bộ ngang
    i = 0
    j = 0
    for row_index in range(76, 86):  # Bắt đầu từ hàng thứ 2 và tăng 2 hàng mỗi lần
        j = 0
        for col_index in range(13, 23):
            if i == j:
                sheet.cells(row_index, col_index).color = (189, 215, 238)
            j += 1
        i += 1

    # Bảng 5: Bảng Cor.
    bang5 = sheet.range("K90:V100")
    sheet.range("K90:V90").api.Font.Bold = True
    sheet.range("K90:V90").color = (248, 203, 173) # cam
    sheet.range("K90:L90").api.Merge()
    sheet.range("K90:L90").value = "CORRELATION"
    sheet.range("K90:L90").color = (255, 217, 102)
    sheet.range("K91:L91").api.Merge()
    sheet.range("K91:L91").value = congty[0]
    sheet.range("K92:L92").api.Merge()
    sheet.range("K92:L92").value = congty[1]
    sheet.range("K93:L93").api.Merge()
    sheet.range("K93:L93").value = congty[2]
    sheet.range("K94:L94").api.Merge()
    sheet.range("K94:L94").value = congty[3]
    sheet.range("K95:L95").api.Merge()
    sheet.range("K95:L95").value = congty[4]
    sheet.range("K96:L96").api.Merge()
    sheet.range("K96:L96").value = congty[5]
    sheet.range("K97:L97").api.Merge()
    sheet.range("K97:L97").value = congty[6]
    sheet.range("K98:L98").api.Merge()
    sheet.range("K98:L98").value = congty[7]
    sheet.range("K99:L99").api.Merge()
    sheet.range("K99:L99").value = congty[8]
    sheet.range("K100:L100").api.Merge()
    sheet.range("K100:L100").value = congty[9]
    sheet.range("K91:L100").color = (248, 203, 173) # cam
    sheet.range("K90:L100").api.Font.Bold = True
    sheet.range("M90:V90").value = congty
       
    bang5.api.HorizontalAlignment = -4108
    bang5.api.VerticalAlignment = -4108
    bang5.api.Borders(7).LineStyle = 1  # Đường viền bên trái
    bang5.api.Borders(8).LineStyle = 1  # Đường viền trên cùng
    bang5.api.Borders(9).LineStyle = 1  # Đường viền dưới cùng
    bang5.api.Borders(10).LineStyle = 1  # Đường viền bên phải
    bang5.api.Borders(11).LineStyle = 1  # Đường viền nội bộ dọc
    bang5.api.Borders(12).LineStyle = 1  # Đường viền nội bộ ngang
    i = 0
    j = 0
    for row_index in range(91, 101):  # Bắt đầu từ hàng thứ 2 và tăng 2 hàng mỗi lần
        j = 0
        for col_index in range(13, 23):
            if i == j:
                sheet.cells(row_index, col_index).color = (189, 215, 238)
            j += 1
        i += 1
    
    # Chọn toàn bộ dữ liệu trong sheet
    data_range = sheet.used_range
    data_range.api.Font.Name = "Times New Roman"
    sheet.range("M8:AP67").number_format = str1
    sheet.range("M68:V70").number_format = str1
    sheet.range("AG68:AP68").number_format = str1
    sheet.range("E76:F195").number_format = str1
    sheet.range("M76:V85").number_format = str1
    sheet.range("M91:V100").number_format = str1[0:6]
def GiaoDienSheet2(sheet, congty):
    # Bảng thông tin
    bang1 = sheet.range("A5:T127")
    sheet.range("A5:A7").api.Merge()
    sheet.range("A5:A7").value = "TT"
    sheet.range("A5:R5").color = (248, 203, 173) # cam
    # stock
    sheet.range("B5:D5").api.Merge()
    sheet.range("B5:D5").value = "Stock"
    sheet.range("B6:B7").api.Merge()
    sheet.range("C6:C7").api.Merge()
    sheet.range("D6:D7").api.Merge()
    sheet.range("B6:D7").value = ["1", "2", "3"]
    sheet.range("B6:D7").color = (198, 224, 180)   #xanh

    # Input    
    sheet.range("E5:N5").api.Merge()
    sheet.range("E5:L5").value = "Input"
    sheet.range("E6:N6").color = (255, 217, 102) # vang
    sheet.range("E6:K6").api.Merge()
    sheet.range("E6:K6").value = "E(r)"
    sheet.range("L6:N6").api.Merge()
    sheet.range("L6:N6").value = "σ"
    sheet.range("E7:N7").color = (198, 224, 180)   #xanh
    sheet.range("E7:L9").value = ["rf", "Stock 1", "Stock 2", "Stock 3", "Cov(1,2)", "Cov(1,3)", "Cov(2,3)", "Stock 1", "Stock 2","Stock 3"]

    # Output
    sheet.range("O5:T5").api.Merge()
    sheet.range("O5:T5").value = "Output"
    sheet.range("O6:O7").api.Merge()
    sheet.range("P6:P7").api.Merge()
    sheet.range("Q6:Q7").api.Merge()
    sheet.range("R6:R7").api.Merge()
    sheet.range("S6:S7").api.Merge()
    sheet.range("T6:T7").api.Merge()
    sheet.range("O6:Q7").color = (198, 224, 180)   #xanh
    sheet.range("R6:T7").color = (248, 203, 173) # cam
    sheet.range("O6:T7").value = ["W1 (Stock 1)", "W2 (Stock 2)", "W3 (Stock 3)", "E(rp)", "σp", "Sharpe ratio"]

    sheet.range("A5:T7").api.Font.Bold = True
    
    bang1.api.HorizontalAlignment = -4108
    bang1.api.VerticalAlignment = -4108
    bang1.api.Borders(7).LineStyle = 1  # Đường viền bên trái
    bang1.api.Borders(8).LineStyle = 1  # Đường viền trên cùng
    bang1.api.Borders(9).LineStyle = 1  # Đường viền dưới cùng
    bang1.api.Borders(10).LineStyle = 1  # Đường viền bên phải
    bang1.api.Borders(11).LineStyle = 1  # Đường viền nội bộ dọc
    bang1.api.Borders(12).LineStyle = 1  # Đường viền nội bộ ngang
    for row_index in range(8, 127, 2):  # Bắt đầu từ hàng thứ 2 và tăng 2 hàng mỗi lần
        for col_index in range(1, 21):
            sheet.cells(row_index, col_index).color = (189, 215, 238)
    # Chọn toàn bộ dữ liệu trong sheet
    data_range = sheet.used_range
    data_range.api.Font.Name = "Times New Roman"
    sheet.range("E8:T127").number_format = str1
    sheet.range("O8:Q127").number_format = "0%"
    
def GiaoDienSheet3(sheet):
    # Bảng thông tin
    bang1 = sheet.range("A5:T8")
    sheet.range("A5:A7").api.Merge()
    sheet.range("A5:A7").value = "TT"
    sheet.range("A5:R5").color = (248, 203, 173) # cam
    # stock
    sheet.range("B5:D5").api.Merge()
    sheet.range("B5:D5").value = "Stock"
    sheet.range("B6:B7").api.Merge()
    sheet.range("C6:C7").api.Merge()
    sheet.range("D6:D7").api.Merge()
    sheet.range("B6:D7").value = ["1", "2", "3"]
    sheet.range("B6:D7").color = (198, 224, 180)   #xanh

    # Input    
    sheet.range("E5:N5").api.Merge()
    sheet.range("E5:L5").value = "Input"
    sheet.range("E6:N6").color = (255, 217, 102) # vang
    sheet.range("E6:K6").api.Merge()
    sheet.range("E6:K6").value = "E(r)"
    sheet.range("L6:N6").api.Merge()
    sheet.range("L6:N6").value = "σ"
    sheet.range("E7:N7").color = (198, 224, 180)   #xanh
    sheet.range("E7:L9").value = ["rf", "Stock 1", "Stock 2", "Stock 3", "Cov(1,2)", "Cov(1,3)", "Cov(2,3)", "Stock 1", "Stock 2","Stock 3"]

    # Output
    sheet.range("O5:T5").api.Merge()
    sheet.range("O5:T5").value = "Output"
    sheet.range("O6:O7").api.Merge()
    sheet.range("P6:P7").api.Merge()
    sheet.range("Q6:Q7").api.Merge()
    sheet.range("R6:R7").api.Merge()
    sheet.range("S6:S7").api.Merge()
    sheet.range("T6:T7").api.Merge()
    sheet.range("O6:Q7").color = (198, 224, 180)   #xanh
    sheet.range("R6:T7").color = (248, 203, 173) # cam
    sheet.range("O6:T7").value = ["W1 (Stock 1)", "W2 (Stock 2)", "W3 (Stock 3)", "E(rp)", "σp", "Sharpe ratio"]
    sheet.range("A5:T7").api.Font.Bold = True
    
    bang1.api.HorizontalAlignment = -4108
    bang1.api.VerticalAlignment = -4108
    bang1.api.Borders(7).LineStyle = 1  # Đường viền bên trái
    bang1.api.Borders(8).LineStyle = 1  # Đường viền trên cùng
    bang1.api.Borders(9).LineStyle = 1  # Đường viền dưới cùng
    bang1.api.Borders(10).LineStyle = 1  # Đường viền bên phải
    bang1.api.Borders(11).LineStyle = 1  # Đường viền nội bộ dọc
    bang1.api.Borders(12).LineStyle = 1  # Đường viền nội bộ ngang

    # Bảng 2: assuming
    bang2 = sheet.range("C11:D12")
    sheet.range("C11:D11").api.Merge()
    sheet.range("C11:D11").value = "Assuming"
    sheet.range("C11:D11").api.Font.Bold = True
    sheet.range("C11:D11").color = (248, 203, 173) #cam
    sheet.range("C12:C12").value = "A = "
    sheet.range("C12:C12").color = (255, 217, 102) #vang
    
    bang2.api.HorizontalAlignment = -4108
    bang2.api.VerticalAlignment = -4108
    bang2.api.Borders(7).LineStyle = 1  # Đường viền bên trái
    bang2.api.Borders(8).LineStyle = 1  # Đường viền trên cùng
    bang2.api.Borders(9).LineStyle = 1  # Đường viền dưới cùng
    bang2.api.Borders(10).LineStyle = 1  # Đường viền bên phải
    bang2.api.Borders(11).LineStyle = 1  # Đường viền nội bộ dọc
    bang2.api.Borders(12).LineStyle = 1  # Đường viền nội bộ ngang

    # Bảng 3: y vs 1-y
    bang3 = sheet.range("F11:G13")
    sheet.range("F11:F11").value = "y"
    sheet.range("F12:F12").value = "1 - y"
    sheet.range("F11:F13").color = (248, 203, 173) #cam
    sheet.range("F11:F13").api.Font.Bold = True

    bang3.api.HorizontalAlignment = -4108
    bang3.api.VerticalAlignment = -4108
    bang3.api.Borders(7).LineStyle = 1  # Đường viền bên trái
    bang3.api.Borders(8).LineStyle = 1  # Đường viền trên cùng
    bang3.api.Borders(9).LineStyle = 1  # Đường viền dưới cùng
    bang3.api.Borders(10).LineStyle = 1  # Đường viền bên phải
    bang3.api.Borders(11).LineStyle = 1  # Đường viền nội bộ dọc
    bang3.api.Borders(12).LineStyle = 1  # Đường viền nội bộ ngang

    # Bảng 4: E(rc), oc, U
    bang4 = sheet.range("F15:G17")
    sheet.range("F15:F15").value = "E(rc)"
    sheet.range("F16:F16").value = "σc"
    sheet.range("F17:F17").value = "U"
    sheet.range("F15:F17").color = (248, 203, 173) #cam
    sheet.range("F15:F17").api.Font.Bold = True
    
    bang4.api.HorizontalAlignment = -4108
    bang4.api.VerticalAlignment = -4108
    bang4.api.Borders(7).LineStyle = 1  # Đường viền bên trái
    bang4.api.Borders(8).LineStyle = 1  # Đường viền trên cùng
    bang4.api.Borders(9).LineStyle = 1  # Đường viền dưới cùng
    bang4.api.Borders(10).LineStyle = 1  # Đường viền bên phải
    bang4.api.Borders(11).LineStyle = 1  # Đường viền nội bộ dọc
    bang4.api.Borders(12).LineStyle = 1  # Đường viền nội bộ ngang

    # Bảng 5: 
    bang5 = sheet.range("I11:K16")
    sheet.range("I11:J11").api.Merge()
    sheet.range("I11:J11").value = "Capital"
    sheet.range("I12:J12").api.Merge()
    sheet.range("I12:J12").value = "The risk-free asset"
    sheet.range("I13:J13").api.Merge()
    sheet.range("I13:J13").value = "The risky portfolio"
    sheet.range("I14:J14").api.Merge()
    sheet.range("I15:J15").api.Merge()
    sheet.range("I16:J16").api.Merge()
    sheet.range("I11:J16").color = (248, 203, 173) #cam
    sheet.range("I11:J16").api.Font.Bold = True

    sheet.range("K14:K16").color = (198, 224, 180)
    
    bang5.api.HorizontalAlignment = -4108
    bang5.api.VerticalAlignment = -4108
    bang5.api.Borders(7).LineStyle = 1  # Đường viền bên trái
    bang5.api.Borders(8).LineStyle = 1  # Đường viền trên cùng
    bang5.api.Borders(9).LineStyle = 1  # Đường viền dưới cùng
    bang5.api.Borders(10).LineStyle = 1  # Đường viền bên phải
    bang5.api.Borders(11).LineStyle = 1  # Đường viền nội bộ dọc
    bang5.api.Borders(12).LineStyle = 1  # Đường viền nội bộ ngang
    
    data_range = sheet.used_range
    data_range.api.Font.Name = "Times New Roman"
    sheet.range("E8:T8").number_format = str1
    sheet.range("G11:G12").number_format = str1
    sheet.range("G15:G17").number_format = str1
    sheet.range("K12:K16").number_format = str1[0:4]
    sheet.range("O8:Q8").number_format = "0%"
    sheet.range("G13:G13").number_format = "0%"
