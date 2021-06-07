import matplotlib.pyplot as plt
from matplotlib.pyplot import MultipleLocator
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
#print (time_x)
#print (concentration_y)
############数据从excel导入###############



x_major_locator=MultipleLocator(500)                               #把x轴的刻度间隔设置为500，并存在变量里
y_major_locator=MultipleLocator(30)                                #把y轴的刻度间隔设置为30，并存在变量里

plt.rcParams['figure.figsize'] = (8.0, 4.0)
plt.ion()                                                          #开启interactive mode 成功的关键函数
plt.figure(1)

ax=plt.gca()                                                       #ax为两条坐标轴的实例
ax.xaxis.set_major_locator(x_major_locator)
ax.yaxis.set_major_locator(y_major_locator)
plt.xlabel('time(s)',fontsize=14)
plt.ylabel('N$\mathregular{a^+}$ (mM)',fontsize=14)
plt.xlim(0,4001)
plt.ylim(0,120)

for i in range(nrows):
    plt.plot(time_x[i],concentration_y[i],'.',color='red')
    plt.pause(0.000000000000001)
