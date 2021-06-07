import matplotlib.pyplot as plt
import numpy as np
import xlwings as xw



############excel文件路径###############
xlfile = 'C:\\Users\\alanl\\Desktop\\Na.xlsx'
############excel文件路径###############




############数据从excel导入###############
time_x=[]
concentration_y=[]

app=xw.App(visible=False,add_book=False)
wb = app.books.open(xlfile)

nrows=wb.sheets['Sheet1'].range('a1').expand('table').rows.count      #统计数据个数，行数

time_x=wb.sheets['Sheet1'].range((1,1),(nrows,1)).value
concentration_y=wb.sheets['Sheet1'].range((1,2),(nrows,2)).value

app.quit()
############数据从excel导入###############
#print (time_x)
#print (concentration_y)


plt.ion()                                                             #开启interactive mode 成功的关键函数
plt.figure(1)

x=[]
y=[]
for i in range(nrows):
    x.append(time_x[i])#模拟数据增量流入
    y.append(concentration_y[i])#模拟数据增量流入
    plt.plot(x,y,'-r')
    
    plt.pause(0.000001)

    #plt.draw()#注意此函数需要调用
    #time.sleep(0.01)

