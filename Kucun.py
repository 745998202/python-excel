# -*- coding= utf-8 -*-

import xlrd
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

def test():
    #获取历史记录表
    workbook = xlrd.open_workbook(r'C:\Users\74599\Desktop\kang\lishi.xls')
    sheet2 = workbook.sheet_by_name('Sheet1')
    HRB = {}
    HRBK = {}
    for gangzhong in range(1,977,1):
         if sheet2.cell(gangzhong,1).value=='HRB400':
             if sheet2.cell(gangzhong,2).value in HRB.keys():
                 HRB[sheet2.cell(gangzhong,2).value]+=sheet2.cell(gangzhong,3).value
             else:
                 HRB[sheet2.cell(gangzhong,2).value]=sheet2.cell(gangzhong,3).value
         elif sheet2.cell(gangzhong,1).value=='HRB400E抗震':
             if sheet2.cell(gangzhong,2).value in HRBK.keys():
                 HRBK[sheet2.cell(gangzhong,2).value]+=sheet2.cell(gangzhong,3).value
             else:
                 HRBK[sheet2.cell(gangzhong,2).value]=sheet2.cell(gangzhong,3).value
    f = xlwt.Workbook()
    sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)
    row0 = [u'品种', u'历史数量']
    for i in range(0, len(row0)):
        sheet1.write(0, i, row0[i], set_style('Times New Roman', 220, True))
    num = 1
    for i in HRB.keys():
        sheet1.write(num,0,'HRB400 '+i)
        sheet1.write(num,1,HRB[i])
        num = num+1
    for i in HRBK.keys():
        sheet1.write(num,0,'HRB400E抗震 '+i)
        sheet1.write(num,1,HRBK[i])
        num = num+1



    #获取入库信息
    ruku = xlrd.open_workbook(r'C:\Users\74599\Desktop\kang\ruku.xlsx')
    rukuSheet = ruku.sheet_by_name('Sheet1') #获取入库那张表
    rukuHRB = {}
    rukuHRBK = {}
    ruku_rows = rukuSheet.nrows
    rukudate = rukuSheet.cell(1,0).value
    onedayHRB = {}
    onedayHRBK = {}
    i = 1
    while(i<ruku_rows):
        if rukuSheet.cell(i, 3).value == 'HRB400':  # HRB
            if rukuSheet.cell(i,4).value in onedayHRB.keys():#不是当天的第一个该规格的零件:
                onedayHRB[rukuSheet.cell(i,4).value]+= rukuSheet.cell(i,5).value
            else:
                onedayHRB[rukuSheet.cell(i,4).value] = rukuSheet.cell(i,5).value
        else:
            if rukuSheet.cell(i,4).value in onedayHRBK.keys():#不是当天的第一个该规格的零件:
                onedayHRBK[rukuSheet.cell(i,4).value]+= rukuSheet.cell(i,5).value
            else:
                onedayHRBK[rukuSheet.cell(i,4).value] = rukuSheet.cell(i,5).value
        if i+1<ruku_rows:
            if rukuSheet.cell(i+1,0).value != rukudate:
                rukuHRB[rukudate] = onedayHRB
                rukuHRBK[rukudate] = onedayHRBK
                rukudate = rukuSheet.cell(i+1,0).value
                onedayHRB = {}
                onedayHRBK = {}
        if i == ruku_rows-1:
            rukuHRB[rukudate] = onedayHRB
            rukuHRBK[rukudate] = onedayHRBK
        i= i+1
    #获取出库那张表
    chuku = xlrd.open_workbook(r'C:\Users\74599\Desktop\kang\chuku.xlsx')
    chukuSheet = chuku.sheet_by_name('Sheet1')
    chuku_rows = chukuSheet.nrows
    chukuHRB ={}
    chukuHRBK = {}
    conedayHRB = {}
    conedayHRBK = {}
    chukudate = chukuSheet.cell(1, 0).value[0:10]
    i = 1
    while (i < chuku_rows):
        if chukuSheet.cell(i, 3).value == 'HRB400':  # HRB
            if chukuSheet.cell(i, 4).value in conedayHRB.keys():  # 不是当天的第一个该规格的零件:
                conedayHRB[chukuSheet.cell(i, 4).value] += chukuSheet.cell(i, 5).value
            else:
                conedayHRB[chukuSheet.cell(i, 4).value] = chukuSheet.cell(i, 5).value
        else:
            if chukuSheet.cell(i, 4).value in conedayHRBK.keys():  # 不是当天的第一个该规格的零件:
                conedayHRBK[chukuSheet.cell(i, 4).value] += chukuSheet.cell(i, 5).value
            else:
                conedayHRBK[chukuSheet.cell(i, 4).value] = chukuSheet.cell(i, 5).value
        if i + 1 < chuku_rows:
            if chukuSheet.cell(i + 1, 0).value[0:10] != chukudate:
                chukuHRB[chukudate] = conedayHRB
                chukuHRBK[chukudate] = conedayHRBK
                chukudate = chukuSheet.cell(i + 1, 0).value[0:10]
                conedayHRB = {}
                conedayHRBK = {}
        if i == chuku_rows - 1:
            chukuHRB[chukudate] = conedayHRB
            chukuHRBK[chukudate] = conedayHRBK
        i = i + 1
    #计算每天的数量并记录
    sheet2 = f.add_sheet(u'每天各规格零件数量',cell_overwrite_ok=True)
    row1 = [u'日期', u'品种',u'进库',u'出库',u'数量']
    for i in range(0, len(row1)):
        sheet2.write(0, i, row1[i], set_style('Times New Roman', 220, True))
    num = 1
    for m,p in zip(rukuHRB.keys(),chukuHRB.keys()):#每一天
        for n in HRB.keys(): #每一个规格的零件
            print(m,' HRB 规格为',n,'数量为',HRB[n])
            sheet2.write(num,0,m)
            if n in rukuHRB[m]: #如果当前规格的零件在当天进了货
                print(m,'进规格为',n,rukuHRB[m][n],'个')
                sheet2.write(num,2,+rukuHRB[m][n])
                HRB[n]+=rukuHRB[m][n]
            if n in chukuHRB[p]:
                print(p,'出规格为',n,chukuHRB[p][n],'个')
                sheet2.write(num, 3, -chukuHRB[p][n])
                HRB[n]-=chukuHRB[p][n]
            print(m,'出库入库后HRB规格为',n,'零件数量为',HRB[n])
            sheet2.write(num,1,'HRB400 '+n)
            sheet2.write(num,4,HRB[n])
            num = num+1
        for n in HRBK.keys(): #每一个规格的零件
            sheet2.write(num, 0, m)
            print(m, ' HRBK 规格为', n, '数量为', HRBK[n])
            if n in rukuHRBK[m]: #如果当前规格的零件在当天进了货
                print(m, '进规格为', n, rukuHRBK[m][n], '个')
                sheet2.write(num, 2, +rukuHRBK[m][n])
                HRBK[n]+=rukuHRBK[m][n]
            if n in chukuHRBK[p]:
                print(p, '出规格为', n, chukuHRBK[p][n], '个')
                sheet2.write(num, 3, -chukuHRBK[p][n])
                HRBK[n]-=chukuHRBK[p][n]
            print(m, '出库入库后HRBK规格为',n,'零件数量为', HRBK[n])
            sheet2.write(num, 1, 'HRB400E抗震 ' + n)
            sheet2.write(num, 4, HRBK[n])
            num = num+1
    f.save('历史总数据.xlsx')
if __name__ == '__main__':
    test();