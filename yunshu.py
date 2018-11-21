import xlwt
import xlrd

groupa = {'HRB400 12*9':625,'HRB400 14*9':625,'HRB400 16*9':625,'HRB400 18*9':625,'HRB400 20*9':625,'HRB400 22*9':559,'HRB400 25*9':541,
              'HRB400E抗震 12*9':625,'HRB400E抗震 14*9':625,'HRB400E抗震 16*9':625,'HRB400E抗震 18*9':625,'HRB400E抗震 20*9':625,'HRB400E抗震 22*9':559,'HRB400E抗震 25*9':541}
groupb = {'HRB400 12*12':234,'HRB400 14*12':196,'HRB400 16*12':191,'HRB400 18*12':166,'HRB400 20*12':168,
              'HRB400 22*12':149,'HRB400 25*12':144,'HRB400 28*9':153,'HRB400 28*12':153,'HRB400 32*9':156,'HRB400 32*12':150,
              'HRB400E抗震 12*12':234,'HRB400E抗震 14*12':196,'HRB400E抗震 16*12':191,'HRB400E抗震 18*12':166,'HRB400E抗震 20*12':168,
              'HRB400E抗震 22*12':149,'HRB400E抗震 25*12':144,'HRB400E抗震 28*9':153,'HRB400E抗震 28*12':153,'HRB400E抗震 32*9':156,'HRB400E抗震 32*12':150}
#仓库类
class room:
    Aroom={}
    Broom={}
    adui = 0
    bdui = 0
    def __init__(self):
        self.Aroom ={}
        self.Broom ={}
        self.adui = 0    #含有a类的堆数
        self.bdui = 0    #含有b类的堆数
#堆场控制类
class roomcontrol:
    rooms = []              #五个堆场
    ranum = [12,12,12,10,10]
    rbnum = [14,20,20,30,36]
    #初始化堆场
    def __init__(self):
        room1 = room()
        room2 = room()
        room3 = room()
        room4 = room()
        room5 = room()
        self.rooms = [room1,room2,room3,room4,room5]
    def removeIn(self,bag,num):       #参数2：规格名称   参数3： 此规格的数量
        ran = [2,3,4,1,0]   #3\4\5\2\1
        if bag in groupa.keys():  #货物为A类
            for i in ran:        #遍历五个仓库
                while num != 0:       #数量不为0（还没有把货物完全放入仓库）
                    if self.rooms[i].adui < self.ranum[i]:#当仓库堆数没有满时
                        if bag in self.rooms[i].Aroom:   #当堆内含有这种规格的零件时
                            if self.rooms[i].Aroom[bag]%groupa[bag] < groupa[bag]: #当堆中零件数没有超标时
                                if num > groupa[bag]-self.rooms[i].Aroom[bag]%groupa[bag]:#当一次放不下时
                                    num = num - (groupa[bag] - self.rooms[i].Aroom[bag]%groupa[bag]) #num 减去差值
                                    self.rooms[i].Aroom[bag] = self.rooms[i].Aroom[bag]+(groupa[bag] - self.rooms[i].Aroom[bag]%groupa[bag])     #直接放到最大容量
                                else:#一次可以放下
                                    self.rooms[i].Aroom[bag] = self.rooms[i].Aroom[bag]+num
                                    num = 0
                            else:                  #当堆中零件超标时
                                if num > groupa[bag]:        #一堆放不下的时候
                                    self.rooms[i].Aroom[bag] = self.rooms[i].Aroom[bag]+groupa[bag]
                                    num = num - groupa[bag]
                                else:                        #一堆可以放下的时候
                                    self.rooms[i].Aroom[bag] = self.rooms[i].Aroom[bag]+num
                                    num = 0
                                self.rooms[i].adui = self.rooms[i].adui+1
                        else:                            #当堆内不含这种规格的零件时
                            if num > groupa[bag]:  # 一堆放不下的时候
                                num = num - groupa[bag]
                                self.rooms[i].Aroom[bag] = groupa[bag]
                            else:                  # 一堆可以放下的时候
                                self.rooms[i].Aroom[bag] = num
                                num = 0
                            self.rooms[i].adui = self.rooms[i].adui+1
                    else:                                       #当仓库堆数满了时
                        if bag in self.rooms[i].Aroom:          # 当堆内含有这种规格的零件时
                            if self.rooms[i].Aroom[bag]%groupa[bag] < groupa[bag] and self.rooms[i].Aroom[bag]%groupa[bag]!=0:  # 当堆中零件数没有超标时
                                if num > groupa[bag]-self.rooms[i].Aroom[bag]%groupa[bag]:#当一次放不下时
                                    num = num - (groupa[bag] - self.rooms[i].Aroom[bag]%groupa[bag]) #num 减去差值
                                    self.rooms[i].Aroom[bag] = self.rooms[i].Aroom[bag]+(groupa[bag] - self.rooms[i].Aroom[bag]%groupa[bag])    #直接放到最大容量
                                else:#一次可以放下
                                    self.rooms[i].Aroom[bag] = self.rooms[i].Aroom[bag]+num
                                    num = 0
                            else:
                                break
                        else:                                   #当堆内不含这种规格的零件时
                            break
        else:                         #此规格为B类
            for i in ran:        #遍历五个仓库
                while num != 0:       #数量不为0（还没有把货物完全放入仓库）
                    if self.rooms[i].bdui < self.rbnum[i]:#当仓库堆数没有满时
                        if bag in self.rooms[i].Broom:   #当堆内含有这种规格的零件时
                            if self.rooms[i].Broom[bag]%groupb[bag] < groupb[bag]: #当堆中零件数没有超标时
                                if num > groupb[bag]-self.rooms[i].Broom[bag]%groupb[bag]:#当一次放不下时
                                    num = num - (groupb[bag] - self.rooms[i].Broom[bag]%groupb[bag]) #num 减去差值
                                    self.rooms[i].Broom[bag] = self.rooms[i].Broom[bag]+(groupb[bag] - self.rooms[i].Broom[bag]%groupb[bag])     #直接放到最大容量
                                else:#一次可以放下
                                    self.rooms[i].Broom[bag] = self.rooms[i].Broom[bag]+num
                                    num = 0
                                self.rooms[i].bdui = self.rooms[i].bdui+1
                            else:                  #当堆中零件超标时
                                if num > groupb[bag]:        #一堆放不下的时候
                                    self.rooms[i].Broom[bag] = self.rooms[i].Broom[bag]+groupb[bag]
                                    num = num - groupb[bag]
                                else:                        #一堆可以放下的时候
                                    self.rooms[i].Broom[bag] = self.rooms[i].Broom[bag]+num
                                    num = 0
                                self.rooms[i].bdui = self.rooms[i].bdui + 1
                        else:                            #当堆内不含这种规格的零件时
                            if num > groupb[bag]:  # 一堆放不下的时候
                                num = num - groupb[bag]
                                self.rooms[i].Broom[bag] = groupb[bag]
                            else:                  # 一堆可以放下的时候
                                self.rooms[i].Broom[bag] = num
                                num = 0
                    else:                                       #当仓库堆数满了时
                        if bag in self.rooms[i].Broom:          # 当堆内含有这种规格的零件时
                            if self.rooms[i].Broom[bag]%groupb[bag] < groupb[bag] and self.rooms[i].Broom[bag]%groupb[bag]!=0:  # 当堆中零件数没有超标时
                                if num > groupb[bag]-self.rooms[i].Broom[bag]%groupb[bag]:#当一次放不下时
                                    num = num - (groupb[bag] - self.rooms[i].Broom[bag]%groupb[bag]) #num 减去差值
                                    self.rooms[i].Broom[bag] = self.rooms[i].Broom[bag]+(groupb[bag] - self.rooms[i].Broom[bag]%groupb[bag])    #直接放到最大容量
                                else:#一次可以放下
                                    self.rooms[i].Broom[bag] = self.rooms[i].Broom[bag]+num
                                    num = 0
                            else:
                                break
                        else:                                   #当堆内不含这种规格的零件时
                            break

    def printAll(self):
        for i in range(5):
            print('第',i+1,'个堆场的数据')
            print(self.rooms[i].Aroom)
            print(self.rooms[i].Broom)
    def removeOut(self,bag,num):#出库
        ran = [2, 3, 4, 1, 0]  # 3\4\5\2\1
        cars = ''
        if bag in groupa.keys():  #规格为A类时
            for i in ran:
                while num !=0:        #车没有装完需要的货物
                    if bag in self.rooms[i].Aroom:    #含有该规格的货物
                        if num < self.rooms[i].Aroom[bag]:   #货物源足够
                            cars = cars + '/'+ str(i+1)
                            self.rooms[i].Aroom[bag] = self.rooms[i].Aroom[bag]-num
                            num = 0
                            return cars
                        else:        #货物源不足（一次不够）
                            cars = cars + '/'+str(i+1)
                            num = num - self.rooms[i].Aroom[bag]
                            self.rooms[i].Aroom[bag] = 0
                            break
                    else:
                        break
        else:                     #规格为B类时
            for i in ran:
                while num !=0:        #车没有装完需要的货物
                    if bag in self.rooms[i].Broom:    #含有该规格的货物
                        if num < self.rooms[i].Broom[bag]:   #货物源足够
                            cars = cars + '/'+ str(i+1)
                            self.rooms[i].Broom[bag] = self.rooms[i].Broom[bag]-num
                            num = 0
                            return cars
                        else:        #货物源不足（一次不够）
                            cars = cars + '/'+str(i+1)
                            num = num - self.rooms[i].Broom[bag]
                            self.rooms[i].Broom[bag] = 0
                            break
                    else:
                        break
        return cars


if __name__ == '__main__':

    #创建六个仓库，每个仓库初始不放货物
    Rooms = roomcontrol()
    workbook = xlrd.open_workbook(r'C:\Users\74599\Desktop\kang\历史总数据.xlsx')
    sheet1 = workbook.sheet_by_name('sheet1')
    rows = sheet1.nrows #获取行数

    #历史记录入库
    for i in range(1,rows,1):
        bag = sheet1.cell(i,0).value
        num = sheet1.cell(i,1).value
        Rooms.removeIn(bag, num)
    #Rooms.printAll()

    #入库出库的初始参数设置
    sheet2 = workbook.sheet_by_name('每天各规格零件数量')
    data = ''   #用于获取当前日期
    ruku = ''   #入库的规格
    rukunum = ''#入库的数量
    chuku = ''  #出库的规格
    chukunum = 0#出库的数量
    road = ''   #车辆线路

    #获取出库表
    chukuworkbook = xlrd.open_workbook(r'C:\Users\74599\Desktop\kang\chu.xlsx')
    chusheet1 = chukuworkbook.sheet_by_name('Sheet1')
    chukudata = chusheet1.cell(1,0).value[0:10]   #初始日期
    z = 0                                   #逐行读取出库信息的标记量
    resultWorkbook = xlrd.open_workbook('result.xlsx')
    resultsheet = resultWorkbook.sheet_by_name('sheet1')
    nowcar = resultsheet.cell(1,0).value    #当前车辆
    nowcarcount = 1                         #记录当前车辆的标记量
    #进行写入的sheet
    f = xlwt.Workbook()
    rsheet = f.add_sheet(u'sheet1', cell_overwrite_ok=True)
    row0 = [u'车次', u'路径']
    for i in range(0, len(row0)):
        rsheet.write(0, i, row0[i])

    #一共10天计算
    for i in range(300):#一共10天
        data = sheet2.cell(i*36+1,0).value  #获取日期
        # 每一天入库
        for j in range(36):
            if sheet2.cell(i*36+j+1,2).value!='':
                print('入库',sheet2.cell(i*36+j+1,2).value)
                ruku = sheet2.cell(i*36+j+1,1).value
                rukunum = sheet2.cell(i*36+j+1,2).value
                Rooms.removeIn(ruku,rukunum)
        # 每一天出库
        while (chusheet1.cell(z + 1, 0).value[0:10] == chukudata):  # 还是这一天
            while(chusheet1.cell(z+1,0).value == nowcar):   #还是一辆车
                z = z + 1                                   #指针移到下一行记录
                chuku = chusheet1.cell(z,3).value + ' ' + chusheet1.cell(z,4).value
                chukunum = chusheet1.cell(z,5).value
                road = road + Rooms.removeOut(chuku,chukunum)#获得出库路径
            rsheet.write(nowcarcount, 1, road)
            road = ''
            nowcarcount = nowcarcount + 1
            nowcar = resultsheet.cell(nowcarcount, 0).value  # 当前车辆
        chukudata = chusheet1.cell(z + 1, 0).value[0:10] #如果是第二天，更改当前出库日期
    f.save('haha.xlsx')


    # print(sheet1.nrows)
    # print(sheet1.ncols)



