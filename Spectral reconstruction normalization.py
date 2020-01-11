# -*- coding: utf-8 -*-
"""
Created on Tue Jul 23 15:20:02 2019

@author: slw
"""
import numpy as np
import xlrd
import xlwt

def excel2m(path):
    data = xlrd.open_workbook(path)
    table = data.sheets()[0]
    nrows = table.nrows  # 行数
    ncols = table.ncols  # 列数
    datamatrix = np.zeros((nrows, ncols))
    for x in range(ncols):
        cols = table.col_values(x)
        cols1 = np.matrix(cols)  # 把list转换为矩阵进行矩阵操作
        datamatrix[:, x] = cols1 # 把数据进行存储
    return datamatrix

def save(data,path):
    f = xlwt.Workbook()  # 创建工作簿
    sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)  # 创建sheet
    [h, l] = data.shape  # h为行数，l为列数
    for i in range(h):
        for j in range(l):
            sheet1.write(i, j, data[i, j])
    f.save(path)
    
#计算5点3阶SG算子的系数矩阵----------------------------------------------
A=np.mat([[1,-2,4,-8],[1,-1,1,-1],[1,0,0,0],[1,1,1,1],[1,2,4,8]])
B=(A*(A.T*A).I)*A.T


#读取并裁剪光谱数据----------------------------------------------------
data=excel2m('C:/Users/slw/Desktop/pyth test/data.xls')
linenum=data.shape[0]
rownum=data.shape[1]
wavenum=data[:,0]
wavenummin=0
wavenummax=0
for i in range(0,linenum):
    if wavenum[i]<800:
        wavenummin=wavenummin+1

for i in range(0,linenum):
    if wavenum[i]<=1800:
        wavenummax=wavenummax+1
newdata=data[wavenummin:wavenummax+1,0:rownum]        


#计算新光谱------------------------------------------------------------
new_linenum=newdata.shape[0]
new_rownum=newdata.shape[1]
new_spectra=np.zeros([1001,new_rownum])
for j in range(1,rownum):
    new_spectra[0,j]=B[0]*newdata[0:5,j:j+1]
    new_spectra[1,j]=B[1]*newdata[0:5,j:j+1]
    new_spectra[999,j]=B[3]*newdata[new_linenum-5:new_linenum,j:j+1]
    new_spectra[1000,j]=B[4]*newdata[new_linenum-5:new_linenum,j:j+1]
    for i in range(2,999):
        error=np.zeros([new_linenum,1])
        for p in range(0,new_linenum):
            error[p]=abs(newdata[p,0]-i-800)
        min_error=min(error)
        for q in range(0,new_linenum):
            if (error[q]==min_error):
                min_num=q
        factor=[newdata[min_num-2,j],newdata[min_num-1,j],newdata[min_num,j],newdata[min_num+1,j],newdata[min_num+2,j]]       
        factor=(np.matrix(factor)).T
        new_spectra[i,j]=B[2]*factor;

a=np.matrix(np.linspace(800,1800,1001)).T   
for i in range(0,a.shape[0]) :
    new_spectra[i,0]=a[i,0]
    
#归一化光谱--------------------------------------------
for j in range(1,rownum):    
    min_data=min(new_spectra[:,j])
    max_data=max(new_spectra[:,j])
    for i in range(0,1001):
        new_spectra[i,j]=(new_spectra[i,j]-min_data)/(max_data-min_data)
        
save(new_spectra,'C:/Users/slw/Desktop/pyth test/result.xls')