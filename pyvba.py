# !/usr/bin/env python
# -*- coding: utf-8 -*-
'''
@Author:      cz
@Tool:        Sublime Text3
@DateTime:    2018-12-20 12:17:39
'''
#subject:py-vba Excel-files merge dev.
import win32com
from win32com.client import Dispatch, constants
import os
import win32con,win32api
import urllib.request
import tkinter as tk
from tkinter import ttk
import tkinter.messagebox

#启动配置
def startnow():
	if not os.path.exists(os.getcwd()+'/vb.xlsm'):
		ur='http://datadt.oss-cn-beijing.aliyuncs.com/data/vb.xlsm'#oss资源包
		try:
			urllib.request.urlretrieve(ur,'vb.xlsm')
			win32api.SetFileAttributes('vb.xlsm', win32con.FILE_ATTRIBUTE_HIDDEN)
		except:
			tkinter.messagebox.showinfo("提示","请检查网络是否连接")

#调用vba宏模块
def useVBA(file_path,VBA):
    xlApp = win32com.client.DispatchEx("Excel.Application")#打开excel操作环境
    xlApp.Visible = False #True（1）进程可见，False（0）暗自进行
    xlApp.DisplayAlerts = 0#Excel窗口静默加载处理
    xlBook = xlApp.Workbooks.Open(file_path,False)#打开文件，有时候会有警告框说由外部链接什么的（与里面公式有关），要点是则True，否则False
    xlBook.Application.Run(VBA) #宏模块
    xlBook.Close(True)#关闭该文件，并保存，不保存就是False
    xlApp.quit()#关闭excel操作环境

#合并命令
def Mergefiles():
	startnow()
	global info
	info=tkinter.StringVar()
	if cb.get()=='多文件单工作表':
		useVBA(os.getcwd()+'/vb.xlsm','Filessigst')
		info.set('√[多文件单工作表]合并命令执行完成!')
	elif cb.get()=='单文件多工作表':
		useVBA(os.getcwd()+'/vb.xlsm','Filedobst')
		info.set('√[单文件多工作表]合并命令执行完成!')
	elif cb.get()=='多文件多工作表':
		useVBA(os.getcwd()+'/vb.xlsm','Filesdobst')
		info.set('√[多文件多工作表]合并命令执行完成!')
	else:
		useVBA(os.getcwd()+'/vb.xlsm','Filesspcst')
		info.set('√[多文件指定多表]合并命令执行完成!')
	l3=tk.Label(myapp,textvariable=info,font=('Microsoft YaHei UI',10),width=50,height=2,fg='red')
	l3.place(x=50,y=200)
#帮助
def tips():
	tkinter.messagebox.showinfo('帮助','1.初始化请保持网络畅通,需配置相关文件;\n2.表格文件的合并暂时支持xls/xlsx/csv格式;\n3.将要合并的文件放在与程序同一文件夹下;\n4.下拉选择合并模式后再点击开始合并按钮;\n5.本程序由datadt开发,仅供个人学习使用！\n--------------搭塔@2018--------------')

#菜单
def menus(myapp):
    menu=tk.Menu(myapp)
    menu.add_cascade(label='帮助',command=tips)
    menu.add_cascade(label='退出',command=myapp.quit)
    myapp.config(menu=menu)	

#主程序
myapp=tk.Tk()
myapp.title('表格合并小助手 搭塔@datadt')
myapp.resizable(0,0) #框体大小可调性，分别表示x,y方向的可变性
myapp.geometry('500x300')#主框体大小
menus(myapp)#启用菜单布局
frm=tk.Frame(myapp,width=500,height=222)#构建一个框架,放置主功能模块
frm.pack()
l1=tk.Label(frm,text='表格合并',font=('Arial',20),width=10,height=2,fg='blue').place(x=175,y=10)
l2=tk.Label(frm,text='选择模式',font=('Microsoft YaHei UI',10),width=10,height=1).place(x=150,y=100)
cbvalue=tk.StringVar()
cb=ttk.Combobox(frm,textvariable=cbvalue,font=('Microsoft YaHei UI',10),width=13)
cb["values"]=("多文件单工作表","单文件多工作表","多文件多工作表","多文件指定多表")
cb.current(0)#默认选择第一种模式
cb.place(x=220,y=100)
b=tk.Button(frm,text='开始合并',font=('Microsoft YaHei UI',15),width=18,height=1,fg='blue',bg='Coral',command=Mergefiles).place(x=140,y=130)
myapp.mainloop()