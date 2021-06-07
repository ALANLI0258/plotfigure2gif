import matplotlib.pyplot as plt
import numpy as np
import xlwings as xw
import imageio


############excel文件路径###############
xlfile = 'C:\\Users\\alanl\\Desktop\\plot_demo\\wjl.xlsx'
############excel文件路径###############

################数据点个数##############
dots_num=100
################数据点个数##############

#############坐标轴数字的大小、字体和间隔################
#x_ticks = np.arange(0, 3501, 500)        #s计时，左闭右开
x_ticks = np.arange(0, 61, 20)        #min计时，左闭右开
#x_lim = (0, 3500)                       #s计时，闭区间
x_lim = (0, 60)                          #min计时，闭区间

NA_ticks = np.arange(0, 121, 30)       
NA_lim = (0, 120)
  
K_ticks = np.arange(0, 33, 8)
K_lim = (0, 32)
 
CA_ticks = np.arange(0, 4.5, 1)   
CA_lim = (0, 4)
PH_ticks = np.arange(4, 8.5, 1)          
PH_lim = (4, 8)                      
#############坐标轴数字的大小、字体和间隔################


######################################数据开始从excel导入##################################
######################################数据开始从excel导入##################################
######################################数据开始从excel导入##################################
time_x=[]
NA=[]
K=[]
CA=[]
PH=[]

app=xw.App(visible=False,add_book=False)
wb = app.books.open(xlfile)

nrows=wb.sheets['Sheet1'].range('a1').expand('table').rows.count      #统计数据个数，行数

time_x=wb.sheets['Sheet1'].range((2,1),(nrows,1)).value
NA=wb.sheets['Sheet1'].range((2,2),(nrows,2)).value
K=wb.sheets['Sheet1'].range((2,4),(nrows,4)).value
CA=wb.sheets['Sheet1'].range((2,6),(nrows,6)).value
PH=wb.sheets['Sheet1'].range((2,8),(nrows,8)).value

app.quit()
#print (time_x)
#print (concentration_y)
####################################数据结束从excel导入##################################
####################################数据结束从excel导入##################################
####################################数据结束从excel导入##################################


#######################################开始画图########################################
#######################################开始画图########################################
#######################################开始画图########################################
fig = plt.figure(figsize=(18,7))
ax1 = fig.add_subplot(221)
ax2 = fig.add_subplot(222)
ax3 = fig.add_subplot(223)
ax4 = fig.add_subplot(224)

plt.ion()                                                          #开启interactive mode 成功的关键函数
#plt.figure(1)

############坐标轴###############
ax1.set_xticks(x_ticks)
ax1.set_yticks(NA_ticks)
ax1.tick_params(axis='x', labelsize= 14)
ax1.tick_params(axis='y', labelsize= 14)
ax1.set_xlim(x_lim)
ax1.set_ylim(NA_lim)
#ax1.set_xlabel('time(s)',fontsize=14)
ax1.set_ylabel('N$\mathregular{a^+}$ (mM)',fontsize=14)

ax2.set_xticks(x_ticks)
ax2.set_yticks(K_ticks)
ax2.tick_params(axis='x', labelsize= 14)
ax2.tick_params(axis='y', labelsize= 14)
ax2.set_xlim(x_lim)
ax2.set_ylim(K_lim)
#ax2.set_xlabel('time(s)',fontsize=14)
ax2.set_ylabel('$\mathregular{K^+}$ (mM)',fontsize=14)

ax3.set_xticks(x_ticks)
ax3.set_yticks(CA_ticks)
ax3.tick_params(axis='x', labelsize= 14)
ax3.tick_params(axis='y', labelsize= 14)
ax3.set_xlim(x_lim)
ax3.set_ylim(CA_lim)
#ax3.set_xlabel('time(s)',fontsize=14)
ax3.set_xlabel('time(min)',fontsize=14)
ax3.set_ylabel('$C\mathregular{a^{2+}}$ (mM)',fontsize=14)
ax3.yaxis.set_label_coords(-0.075,0.5)

ax4.set_xticks(x_ticks)
ax4.set_yticks(PH_ticks)
ax4.tick_params(axis='x', labelsize= 14)
ax4.tick_params(axis='y', labelsize= 14)
ax4.set_xlim(x_lim)
ax4.set_ylim(PH_lim)
#ax4.set_xlabel('time(s)',fontsize=14)
ax4.set_xlabel('time(min)',fontsize=14)
ax4.set_ylabel('pH',fontsize=14)
ax4.yaxis.set_label_coords(-0.07,0.5)
############坐标轴###############
  

#################间隔取点##############
dots=[]
dots=np.linspace(0,nrows,num=dots_num,endpoint=False, dtype = int)
#print (dots)
#################间隔取点##############

#################图片展示并保存##############
pngnames_list = []
for i in range(len(dots)):
    ax1.plot(time_x[dots[i]]/60,NA[dots[i]],'.',color='#1A6FDF')
    ax2.plot(time_x[dots[i]]/60,K[dots[i]],'.',color='#F14040')  
    ax3.plot(time_x[dots[i]]/60,CA[dots[i]],'.',color='#FB6501')  
    ax4.plot(time_x[dots[i]]/60,PH[dots[i]],'.',color='#9900FF') 
    pngname='png/'+str(i)+'.png'                                                           #图片名
    fig.savefig(pngname, dpi=300, bbox_inches = 'tight', format='png')                     #保存
    pngnames_list.append(pngname)                                                          #图片名列表，做gif用
    plt.pause(0.01)

plt.ioff()
#plt.show()
#################图片展示并保存##############


#######################################结束画图########################################
#######################################结束画图########################################
#######################################结束画图########################################


#######################################开始做gif########################################
#######################################开始做gif########################################
#######################################开始做gif########################################
gif_images = []
for png in pngnames_list:
    gif_images.append(imageio.imread(png))

imageio.mimsave("test.gif",gif_images,fps=5)        #fps参数越大播放的速率越大，fps越小播放的速度越慢
#######################################结束gif########################################
#######################################结束gif########################################
#######################################结束gif########################################

