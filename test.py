import xlwt
import xlrd
def set_style(name,height,bold=False):
    style = xlwt.XFStyle()  # 初始化样式
    font = xlwt.Font()  # 为样式创建字体
    font.name = name # 'Times New Roman'
    font.bold = bold
    font.color_index = 4
    font.height = height
    style.font = font
    return style
if __name__ == '__main__' :
    f = {'a':1,'b':2}
    print(f.keys())
    print(f['a'])
    print(str(1+1))
    workbook = xlrd.open_workbook(r'C:\Users\74599\Desktop\kang\chu.xlsx')#打开出库目录
    sheet1 = workbook.sheet_by_name('Sheet1')
    nrows = sheet1.nrows
    f = xlwt.Workbook()
    sheet = f.add_sheet(u'sheet1', cell_overwrite_ok=True)
    row0 = [u'车次', u'路径']
    for i in range(0, len(row0)):
        sheet.write(0, i, row0[i], set_style('Times New Roman', 220, True))
    car = sheet1.cell(1,0).value
    count = 1
    for i in range(1,nrows,1):
        if sheet1.cell(i,0).value != car:
            sheet.write(count,0,car)
            count = count + 1
            car = sheet1.cell(i,0).value
    f.save('result.xlsx')


