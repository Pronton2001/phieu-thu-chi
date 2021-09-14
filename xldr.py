# To add a new cell, type '# %%'
# To add a new markdown cell, type '# %% [markdown]'
# %% [markdown]
# ## Đổi từ dạng số sang dạng tiền

# %%
print('Bắt đầu chạy chương trình:')
def VND(number):
    stringNumber = str(int(number))
    money = ''
    for i in range(len(stringNumber)):
        if (len(stringNumber) - i) % 3 == 0 and i != 0:
            money += '.'
        money+=stringNumber[i]
    money += 'đ'
    return money
print("10%")

# %% [markdown]
# ## Đổi số sang chữ - Nguồn https://ideone.com/OZJ7oa và https://daynhauhoc.com/t/challenge-doc-so-va-viet-so/58832/13
# 

# %%
one_digit = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"]
units = ["", " mốt", " hai", " ba", " bốn", " lăm", " sáu", " bảy", " tám", " chín"]
tens = ["lẻ", "mười"] + [x + " mươi" for x in ["hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"]]
hundreds = [x + " trăm" for x in one_digit]
bigger = ["", " nghìn", " triệu", " tỉ", " nghìn tỉ", " triệu tỉ", " tỉ tỉ"]
 
def read_class(s):
	n = int(s)
	if len(s) == 1:
		return one_digit[n]
	elif len(s) == 2:
		if s == '11':
			return tens[n//10] + ' một' 
		return tens[n//10] + units[n%10]
	else:  # len(s) == 3
		if s == '000':
			return ''
		elif (n//10)%10 == 0:  # x0y
			if n % 10 == 0:  # x00
				return hundreds[n//100]
			else:
				return hundreds[n//100] + ' ' + tens[(n//10)%10] + ' ' + one_digit[n%10]
		else:  # xyz
			if n%100 == 11:
				return hundreds[n//100] + ' ' + tens[(n//10)%10] + ' một'
			return hundreds[n//100] + ' ' + tens[(n//10)%10] + units[n%10]
 
def Convert(s):
	if (s == '0'):
		return 'Không đồng'
	s = s.lstrip('0').rstrip('\n')
	classes = []
	for i in range(len(s)-1, 1, -3):
		classes.append(s[i-2:i+1])
	if len(s) % 3 != 0:
		classes.append(s[:len(s)%3])
	res = ''
	for i in range(len(classes)):
		named_class = read_class(classes[i])
		if named_class != '':
			res = named_class + bigger[i] + (' ' * (res != '')) + res
	res_new = res[0].upper() + res[1:] + ' đồng'
	return res_new
print('20%')

# %% [markdown]
# # Danh Sách Tên

# %%
import xlrd
import datetime

data = xlrd.open_workbook('Dữ liệu đầu vào.xls')
sh = data.sheet_by_index(0)
# Ten doanh nghiep, dia chi, Giam doc, Ke toan truong, Nguoi lap Phieu, Thu quy
# arr [0                1       2           3               4               5]
arr = [sh.row(i)[2].value for i in range(1,7)] 
print('30%')

# %% [markdown]
# # Phieu Chi and SO QUY TIEN MĂT 2017

# %%
ex = xlrd.open_workbook('ex.xls', formatting_info=True)
sh2 = ex.sheet_by_index(0)

# sh2.cell(0,0),       # Ten doanh nghiep
# sh2.cell(2, 0),      # Dia chi
# sh2.cell(21, 3),     # Nguoi lap Phieu
# sh2.cell(21, 6),     # Thu quy
# sh2.cell(21, 14),    # Ke toan truong
# sh2.cell(21, 20)     # Giam doc

qt = data.sheet_by_index(1)
print('40%')

# %% [markdown]
# ## Lọc theo tháng

# %%
rows = {}
rows_by_month = {}
for i in range (0, qt.nrows): 
    # Check if this row is UNVALID
    if isinstance (qt.cell(i,0).value, str):
        continue
    date = xlrd.xldate_as_datetime(qt.cell(i, 1).value, datemode = 0)

    phieuId = ''
    # Check for safety if both [phieuThuID] and [phieuChiID] is blank
    if (qt.cell(i,2).value == '' and qt.cell(i,3).value != ''):
        phieuId = qt.cell(i,3).value
    elif (qt.cell(i,2).value != '' and qt.cell(i,3).value == ''):
        phieuId = qt.cell(i,2).value
    
    tienThu = 0
    # Check for safety if both [tienThu] and [tienChi] is blank (convert to zero) or zero
    if isinstance(qt.cell(i,9).value, str):
        if (qt.cell(i,9) == ''):
            tienThu = 0
        else:
            tienThu = int(qt.cell(i,9))
    if isinstance(qt.cell(i,11).value, str):
        if (qt.cell(i,11) == ''):
            tienThu = 0
        else:
            tienThu = int(qt.cell(i,11))

    # If phieuID is available in list [rows.keys()]
    if phieuId in rows.keys(): 
        rows[phieuId]['tien'] += (qt.cell(i,9).value +  qt.cell(i,11).value)
        if qt.cell(i,8).value not in rows[phieuId]['tk']:
            rows[phieuId]['tk'] += (',' + qt.cell(i,8).value)
    else:
        rows[phieuId] = {
            'date'  :   date,
            'lyDo'  :   qt.cell(i,6).value,
            'tk'    :   qt.cell(i,8).value,
            'tien'  :   qt.cell(i,9).value +  qt.cell(i,11).value,
            'kemTheo':   qt.cell(i,15).value,
            'hoTen' :   qt.cell(i,16).value,
            'diaChi':   qt.cell(i,17).value,
            'loai'  :   'phieuChi' if (qt.cell(i,2).value == '') else 'phieuThu'
        }
        
    # Check for the [nextDate] is next month or no valid ('')
    if isinstance(qt.cell(i + 1, 1).value, float):
        nextDate    = xlrd.xldate_as_datetime(qt.cell(i + 1, 1).value, datemode = 0) 
    else:
        nextDate = '' 
    if (nextDate != '') and (date.month != nextDate.month):
        rows_by_month[f'{date.month}-{date.year}'] = rows  # Add [rows] to [rows_by_month]
        rows = {} # Empty [rows]
rows_by_month[f'{date.month}-{date.year}'] = rows  # Add [rows] to [rows_by_month]
rows_by_month.keys()
print('50%')

# %% [markdown]
# # XLWT and copy

# %%
from xlutils.copy import copy # http://pypi.python.org/pypi/xlutils
from copy import deepcopy
import xlwt
from xlwt import Workbook
from datetime import datetime
print('60%')

# %% [markdown]
# ## Write soure \[DanhSachTen\] to destination
# %% [markdown]
# ### Formating and constant

# %%
FONT_SIZE_UNIT = 20

styles = {
    # 'datetime': xlwt.easyxf(num_format_str='yyyy-mm-dd hh:mm:ss'),
    # 'date': xlwt.easyxf(num_format_str='yyyy-mm-dd'),
    'time':  xlwt.easyxf(f"font: name Tahoma, height {9*FONT_SIZE_UNIT}, italic on; align: wrap true, horiz right, vert center" ),
    'time2' : xlwt.easyxf(f"font: name Tahoma, height {9*FONT_SIZE_UNIT}, italic on; align: wrap true, horiz left, vert center" ), 
    'header': xlwt.easyxf(f"font: name Tahoma, height {11*FONT_SIZE_UNIT}, bold on; align: wrap true, horiz left, vert center", num_format_str='#,##0.00'),
    'small_header' : xlwt.easyxf(f"font: name Tahoma, height {9*FONT_SIZE_UNIT}, bold on; align: wrap true, horiz left, vert center", num_format_str='#,##0.00'),
    'small_default' : xlwt.easyxf(f"font: name Tahoma, height {9*FONT_SIZE_UNIT}; align: wrap true, horiz left, vert center", num_format_str='#,##0.00'),
    'name' : xlwt.easyxf(f"font: name Times New Roman, height {9*FONT_SIZE_UNIT}, bold on; align: wrap true, horiz center, vert bottom", num_format_str='#,##0.00'),
    'default': xlwt.easyxf(f"font: name Tahoma, height {11*FONT_SIZE_UNIT}; align: wrap true, horiz left, vert center", num_format_str='#,##0.00'),
    'title': xlwt.easyxf(f"font: name Tahoma, height {15*FONT_SIZE_UNIT}, bold on; align: wrap true, horiz center, vert center", num_format_str='#,##0.00'),
    'footer': xlwt.easyxf(f"font: name Tahoma, height {9*FONT_SIZE_UNIT}; align: wrap true, horiz left, vert center", num_format_str='#,##0.00'),
    'medium_default': xlwt.easyxf(f"font: name Tahoma, height {10*FONT_SIZE_UNIT}; align: wrap true, horiz left, vert center", num_format_str='#,##0.00'),
    'small_header_cent' : xlwt.easyxf(f"font: name Tahoma, height {9*FONT_SIZE_UNIT}, bold on; align: wrap true, horiz center, vert center", num_format_str='#,##0.00'),
    }
print('70%')

# %% [markdown]
# ### Hàm Tạo Phiếu thu/chi

# %%
def TaoPhieu(iD, date_ID, w_sheet):
    #----------------------------- Data from DanhSachTen from sheet 1 in Du lieu dau vao.xls ---------------------------#
    w_sheet.write(0, 0, arr[0], styles['header'])                # Ten doanh nghiep
    w_sheet.write(2, 0, arr[1], styles['small_header']),         # Dia chi
    w_sheet.write(21, 20, arr[2], styles['name']),               # Giam doc
    w_sheet.write(21, 14, arr[3], styles['name']),               # Ke toan truong
    w_sheet.write(21, 3, arr[4], styles['name']),                # Nguoi lap Phieu
    w_sheet.write(21, 6, arr[5], styles['name']),                # Thu quy
    
    date    = rows_by_month[date_ID][iD]['date']
    lyDo    = rows_by_month[date_ID][iD]['lyDo']
    tk      = rows_by_month[date_ID][iD]['tk']
    tien    = rows_by_month[date_ID][iD]['tien']
    kemTheo  = rows_by_month[date_ID][iD]['kemTheo']
    hoTen   = rows_by_month[date_ID][iD]['hoTen']
    diaChi  = rows_by_month[date_ID][iD]['diaChi']
    loai    = rows_by_month[date_ID][iD]['loai']
    
    if loai == 'phieuThu':
        availID = tk
        debtID  = '111'
    else: 
        availID = '111'
        debtID  = tk

    #----------------------------------------------- DAY / MONTH / YEAR -----------------------------------------------#

    zeroDay = '0' if date.day / 10 < 1.0 else ''
    zeroMonth = '0' if date.month / 10 < 1.0 else ''

    w_sheet.write(8,7, f'{zeroDay}{date.day}', styles['time2'])
    w_sheet.write(8,10, f'{zeroMonth}{date.month}', styles['time2'])
    w_sheet.write(8,13, date.year, styles['time2'])
    w_sheet.write(18,14, f'Ngày  {zeroDay}{date.day}  tháng  {zeroMonth}{date.month}  năm  {date.year} ', styles['time'])
    
    #----------------------------------------------- phieuID/ So No/ So Co --------------------------------------------#

    w_sheet.write(5, 19, iD, styles['small_default'])
    w_sheet.write(6, 19, debtID, styles['small_default'])
    w_sheet.write(9, 19, availID, styles['small_default'])

    #------------------------------------  HoTen/ Dia Chi/ Ly do/ So tien/ Kem theo -----------------------------------#

    w_sheet.write(12, 4, hoTen, styles['default'])
    w_sheet.write(13, 1, diaChi, styles['default'])
    if loai == 'phieuThu':
        w_sheet.write(5, 7, 'PHIẾU THU', styles['title'])
        w_sheet.write(0, 21, 'Mẫu số 01-TT', styles['small_header_cent'])
        w_sheet.write(14, 0, f'Lý do nộp: {lyDo}', styles['default'])
        w_sheet.write(12, 0, 'Họ và tên người nộp tiền:', styles['default'])
        w_sheet.write(19, 0, 'Người nộp tiền', styles['small_header_cent'])
    else:
        w_sheet.write(5, 7, 'PHIẾU CHI', styles['title'])
        w_sheet.write(0, 21, 'Mẫu số 02-TT', styles['small_header_cent'])
        w_sheet.write(14, 0, f'Lý do chi: {lyDo}', styles['default'])
        w_sheet.write(12, 0, 'Họ và tên người nhận tiền:', styles['default'])
        w_sheet.write(19, 0, 'Người nhận tiền', styles['small_header_cent'])
        
    if int(tien) == 0:
        print(f'Chu y: {date_ID} {iD} co so tien bang 0!')
    w_sheet.write(15, 1, f'{VND(tien)}. (Viết bằng chữ): {Convert(str(int(tien)))}.', styles['default'])
    w_sheet.write(16, 0, f'Kèm theo {kemTheo} chứng từ gốc', styles['default'])
    #----------------------------------------- Chu Ky nguoi nhan tien ------------------------------------------------#

    w_sheet.write(21, 0, hoTen, styles['name'])

    #---------------------------------------------Da Nhan Du Tien ----------------------------------------------------#

    w_sheet.write(23, 0, f'Đã nhận đủ số tiền(viết bằng chữ): {Convert(str(int(tien)))}.', styles['footer'])

    return loai
print('80%')

# %% [markdown]
# ### Tổng hợp các sheet lại thành 1 file

# %%
def TongHop():
    # try:
        wb_chi = copy(ex) # a writable copy (I can't read values out of this, only write to it)
        wb_thu = copy(ex)

        demChi = 0
        demThu = 0
        stopPhieuThu = False
        stopPhieuChi = False
        for month in rows_by_month:
            for phieu in rows_by_month[month]:
                if (stopPhieuThu == False) and (demThu == ex.nsheets - 1):
                    print("Phieu Thu da toi gioi han, xin hay tao them sheet moi")
                    stopPhieuThu = True
                if (stopPhieuChi == False) and (demChi == ex.nsheets - 1):
                    print("Phieu Chi da toi gioi han, xin hay tao them sheet moi")
                    stopPhieuChi = True
                if (stopPhieuThu == False) and (rows_by_month[month][phieu]['loai'] == 'phieuThu'):
                    w_sheet = wb_thu.get_sheet(demThu)
                    demThu += 1
                elif (stopPhieuChi ==False) and (rows_by_month[month][phieu]['loai'] == 'phieuChi'):
                    w_sheet = wb_chi.get_sheet(demChi)
                    demChi += 1
                else:
                    break
                w_sheet.set_name(month + ' ' + phieu)
                # ? This code works if I fix built-in code of Python3, it is not safe for run in other computers now.
                # w_sheet.header_str = '' 
                # w_sheet.footer_str = ''
                TaoPhieu(phieu, month, w_sheet)        
            if stopPhieuThu == True and stopPhieuChi == True:
                break
        print('90%')
        return wb_thu, wb_chi

# %% [markdown]
# ### Lưu lại file tổng

# %%
try:
    wb_thu, wb_chi = TongHop()
    File_PhieuThu = 'Phiếu Thu.xls'
    wb_thu.save(File_PhieuThu)
    File_PhieuChi = 'Phiếu Chi.xls'
    wb_chi.save(File_PhieuChi)
    print('100%')
    print("Da hoan tat chuong trinh")
except Exception as e:
    if 'Permission denied' in str(e):
        print('Loi he thong: Hay kiem tra da dong Phieu Thu.xls va Phieu Chi.xls truoc khi chay chua?')
    else: 
        print('Loi chua xac dinh')
    print('Chua Hoan tat!')

