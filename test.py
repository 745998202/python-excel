import xlwt
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
    f = xlwt.Workbook()
    sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)
    row0 = [u'品种', u'历史数量']
    for i in range(0, len(row0)):
        sheet1.write(0, i, row0[i], set_style('Times New Roman', 220, True))
    sheet1.write_merge(7,8,2,4,'ceshi')
    n = "HYPERLINK"
    sheet1.write_merge(9, 9, 2, 8, xlwt.Formula(n + '("http://www.baidu.com")'),
                       set_style('Arial', 300, True))
    f.save('demo.xlsx')