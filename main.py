#!/usr/bin/python
# -*-coding:utf-8 -*-
'''部分程序说明***
1.本程序可以调用excel软件，运行时请保证电脑已经安装pywin32(调用win32.client模块。
1*.由于编写的电脑很难使用xlutils模块，故采用win32com模块
2.本程序调用了很多外源模块，请确保电脑Python34（本程序采用）已装上相应的模块，且还有mysql、数据库存在。
3.本程序较好的实现了错误异常处理，综合考虑了许多种人为导致的报错情况，在比较好的用户体验的基础上，实现了错误调节机制。
4.本程序保存的时候加了导入信息修改后的保存按钮，之所以这样处理是为了避免修改后部分输入框没动而正常保存会弹出有信息没有写入的界面，
  同时也不能直接把读取的数值赋值给全局变量，因为用户或许不修改，继续录入信息，导致可能有的输入框没写入保存
  时不弹出相应的提示框而录入信息，导致信息错误写入，给日后的查询带来麻烦。
5.本程序在进行权重输入后学生总成绩计算保存时，因为可能打开多个excel文件进行读写操作，会有几秒的运行时间，还望耐心等待。

'''

''' '''
#使用前请仔细阅读本程序说明部分或者点击第一个窗口界面的程序说明按钮查看。(特别是信息修改后应该使用的保存。。。导入修改后保存再修改再保存问题提醒)
''' '''

###***课程管理系统的实现（面向计算机学院课程）***###

# 模块导入（包含外源模块是否存在异常处理问题）,说明类函数内容输入
import tkinter
import os
import subprocess  #该模块用于启动计算机程序，用于后面启动excel文件，方便信息的浏览
try:
    import win32com.client       #用于处理excel表格
    import xlwt                 #用于弥补win32com不能修改单元格宽度的不足
    import numpy as np                            ##模块比较多，程序运行的时候会有大概些许延迟
    import matplotlib.pyplot as plt
    from matplotlib.font_manager import FontProperties
    font=FontProperties(fname="C:\\windows\\fonts\\simsun.ttc",size=14)#指定默认字体与大小(用于显示中文)
except ImportError:
    print("你的电脑python未安装程序所需外源模块\n")
    print("请前往http://pypi.python.org/pypi/nltk或者\n")
    print("http://www.lfd.uci.edu/~gohlke/pythonlibs/或者\n")
    print("http://sourceforge.net/\n")
    print("或者使用电脑cmd命令pip install \"模块名\"或者easy_install.exe \"模块名\"等方式下载安装!")
    os.system("pause")      #暂停进程，使得整个程序停止。

def Use_description():     #说明函数定义，打开相应说明文件。
    try:
        subprocess.Popen("E:/课程管理系统使用说明.txt",shell=True)
    except Exception:
        pass

def BackError(arg):      #窗口跳转异常处理函数，防止多次跳转时某些窗口已经被关闭而报错
    try:
        arg.destroy()
    except Exception:
        pass

def Exit_aware(brg):     #温馨提示窗口的关闭(附带用于关闭成绩分析系统的操作错误提示窗口)
    brg.destroy()



######课程信息组件+学生信息组件+成绩分析组件
class Course_Widget:
    def __init__(self,course_name):
        self.course_name=course_name
    def All_command_one(self):              #某个课程页面的组件
        try:
            BackError(Error1)               #当有输入框内容为空保存时跳出Error1窗口，点击按钮返回app窗口，关闭Error1窗口（规避未进入执错窗口引发的报错）
        except Exception:
            pass
        try:
            BackError(Error21)               #当有输入框内容为空保存时跳出Error21窗口，点击按钮返回app窗口，关闭Error21窗口（规避未进入执错窗口引发的报错）
        except Exception:
            pass
        try:
            BackError(Error5)       #点击修改按钮时信息没导入或者没写入的错误提示窗口的关闭
        except Exception:
            pass
        try:
            BackError(Error3_2)     #查询某个课程信息的时候，整个课程实际上没有存入任何信息。窗口跳转异常处理
        except:
            pass
        BackError(root1)        #跳转关闭异常处理(函数调用)
    
        global app
        app=tkinter.Tk()
        app.title(self.course_name)
        app.geometry("500x400")

        global a1_model
        global b1_model
        global c1_model           #信息查询时用
        a1_model=tkinter.StringVar()
        b1_model=tkinter.StringVar()
        c1_model=tkinter.StringVar()
    
    
        tkinter.Label(app,text="欢迎进入《%s》管理系统"%(self.course_name),font=("楷体",12),
                      bg="#1E90FF",height=3,width=62).grid(row=0,column=0,columnspan=4,sticky="w")
        tkinter.Label(app,text="课程名称:",font=("楷体",12),bg="#00FFFF").grid(row=1,column=0,sticky="w")
        tkinter.Label(app,text=self.course_name,font=("楷体",12),bg="#D3D3D3",relief="sunken").grid(row=1,column=1,sticky="w")
        tkinter.Label(app,text="").grid(row=2,column=0,columnspan=4,sticky="w")
        tkinter.Label(app,text="任课教师:",font=("楷体",12),bg="#00FFFF").grid(row=3,column=0,sticky="w")

        global first_kind_info1
        first_kind_info1=tkinter.Entry(app,font=("楷体",12),bg="#D3D3D3")
        first_kind_info1.bind("<KeyRelease>",Entry_get.first_kind_get_data1)
        first_kind_info1.grid(row=3,column=1,sticky="w")
        tkinter.Label(app,text="").grid(row=4,column=0,columnspan=4,sticky="w")
        tkinter.Label(app,text="上课地点:",font=("楷体",12),bg="#00FFFF").grid(row=5,column=0,sticky="w")

        global first_kind_info2
        first_kind_info2=tkinter.Entry(app,font=("楷体",12),bg="#D3D3D3")
        first_kind_info2.bind("<KeyRelease>",Entry_get.first_kind_get_data2)
        first_kind_info2.grid(row=5,column=1,sticky="w")
        tkinter.Label(app,text="").grid(row=6,column=0,columnspan=4,sticky="w")
        tkinter.Label(app,text="上课时间:",font=("楷体",12),bg="#00FFFF").grid(row=7,column=0,sticky="w")

        global first_kind_info3
        first_kind_info3=tkinter.Entry(app,font=("楷体",12),bg="#D3D3D3")
        first_kind_info3.bind("<KeyRelease>",Entry_get.first_kind_get_data3)
        first_kind_info3.grid(row=7,column=1,sticky="w")
        tkinter.Label(app,text="").grid(row=8,column=0,columnspan=4,sticky="w")
        tkinter.Button(app,text="学生信息\n的管理界面",font=("楷体",18),bg="#EE82EE",
                       command=lambda:Course_Widget(self.course_name).All_command_two()).grid(row=1,column=2,rowspan=4,columnspan=2,sticky="w")
        tkinter.Button(app,text="正常写入保存",font=("楷体",15),bg="#48D1CC",
                       command=lambda:Save_data(self.course_name).Save_data_first_kind()).grid(row=9,column=0,rowspan=3,sticky="w")
        tkinter.Button(app,text="修改",font=("楷体",15),bg="#FFE4B5",
                       command=lambda:Revise(self.course_name).Revise_kind_1()).grid(row=9,column=2,rowspan=3,sticky="w")
        tkinter.Button(app,text="导入查询",font=("楷体",15),bg="#F0F8FF",
                       command=lambda:Info_look(self.course_name).Look_fk_info_Button()).grid(row=9,column=3,rowspan=3,sticky="w")
        tkinter.Button(app,text="录入下一个教师信息",font=("楷体",15),bg="#F0F8FF",
                       command=lambda:Save_continue(self.course_name).Save_next_first_kind()).grid(row=9,column=1,rowspan=3,sticky="w")
        tkinter.Button(app,text="<<--返回总课程界面",font=("楷体",15),bg="#D8BFD8",
                       command=lambda:Back_Turn(self.course_name).Back_first_kind()).grid(row=12,column=0,rowspan=2,columnspan=2,sticky="w")
        tkinter.Button(app,text="删除一个教师信息",font=("楷体",15),bg="#00FF7F",relief="raised",
                       command=lambda:Delete(self.course_name).Delete_fktype()).grid(row=12,column=2,rowspan=2,columnspan=2,sticky="w")
        tkinter.Button(app,text="信息导入\n修改的保存",font=("楷体",18),bg="#FF69B4",
                       command=lambda:Save_data(self.course_name).Insert_Revise_Save_fk()).grid(row=6,column=2,rowspan=3,columnspan=2,sticky="w")
        app.mainloop()


    def All_command_two(self):       #学生基本信息录入的页面组件
        try:
            BackError(Error2)      #当有输入框内容为空保存时跳出Error2窗口，点击按钮返回bpp窗口，关闭Error2窗口（规避未进入执错窗口引发的报错）
        except Exception:
            pass
        try:
            BackError(Error22)      #当有输入框内容为空保存时跳出Error22窗口，点击按钮返回bpp窗口，关闭Error22窗口（规避未进入执错窗口引发的报错）
        except Exception:
            pass
        try:
            BackError(Error6)       #点击修改按钮时信息没导入或者没写入的错误提示窗口的关闭
        except Exception:
            pass
        BackError(app)     #跳转关闭异常处理(函数调用)
    
        global bpp
        bpp=tkinter.Tk()
        bpp.title("《%s》学生信息管理系统"%(self.course_name))
        bpp.geometry("510x500")
        tkinter.Label(bpp,text="《%s》学生信息录入系统"%(self.course_name),font=("楷体",12),bg="#1E90FF",
                      height=3,width=62).grid(row=0,column=0,columnspan=4,sticky="w")
        tkinter.Label(bpp,text="学生姓名:",font=("楷体",12),bg="#00FFFF").grid(row=1,column=0,sticky="w")

        global a2_model
        a2_model=tkinter.StringVar()
        global b2_model
        b2_model=tkinter.StringVar()
        global c2_model
        c2_model=tkinter.StringVar()       ####学生信息查询的六个容器
        global d2_model
        d2_model=tkinter.StringVar()
        global e2_model
        e2_model=tkinter.StringVar()
        global f2_model
        f2_model=tkinter.StringVar()

        global second_kind_info1
        second_kind_info1=tkinter.Entry(bpp,font=("楷体",12),bg="#D3D3D3")
        second_kind_info1.bind("<KeyRelease>",Entry_get.second_kind_get_data1)
        second_kind_info1.grid(row=1,column=1,sticky="w")
        tkinter.Label(bpp,text="").grid(row=2,column=0,columnspan=4,sticky="w")
        tkinter.Label(bpp,text="学生学号:",font=("楷体",12),bg="#00FFFF").grid(row=3,column=0,sticky="w")

        global second_kind_info2
        second_kind_info2=tkinter.Entry(bpp,font=("楷体",12),bg="#D3D3D3")
        second_kind_info2.bind("<KeyRelease>",Entry_get.second_kind_get_data2)
        second_kind_info2.grid(row=3,column=1,sticky="w")
        tkinter.Label(bpp,text="").grid(row=4,column=0,columnspan=4,sticky="w")
        tkinter.Label(bpp,text="所在班级:",font=("楷体",12),bg="#00FFFF").grid(row=5,column=0,sticky="w")

        global second_kind_info3
        second_kind_info3=tkinter.Entry(bpp,font=("楷体",12),bg="#D3D3D3")
        second_kind_info3.bind("<KeyRelease>",Entry_get.second_kind_get_data3)
        second_kind_info3.grid(row=5,column=1,sticky="w")
        tkinter.Label(bpp,text="").grid(row=6,column=0,columnspan=4,sticky="w")
        tkinter.Label(bpp,text="平时成绩:",font=("楷体",12),bg="#00FFFF").grid(row=7,column=0,sticky="w")

        global second_kind_info4
        second_kind_info4=tkinter.Entry(bpp,font=("楷体",12),bg="#D3D3D3")
        second_kind_info4.bind("<KeyRelease>",Entry_get.second_kind_get_data4)
        second_kind_info4.grid(row=7,column=1,sticky="w")
        tkinter.Label(bpp,text="").grid(row=8,column=0,columnspan=4,sticky="w")
        tkinter.Label(bpp,text="期中成绩:",font=("楷体",12),bg="#00FFFF").grid(row=9,column=0,sticky="w")

        global second_kind_info5
        second_kind_info5=tkinter.Entry(bpp,font=("楷体",12),bg="#D3D3D3")
        second_kind_info5.bind("<KeyRelease>",Entry_get.second_kind_get_data5)
        second_kind_info5.grid(row=9,column=1,sticky="w")
        tkinter.Label(bpp,text="").grid(row=10,column=0,columnspan=4,sticky="w")
        tkinter.Label(bpp,text="期末成绩:",font=("楷体",12),bg="#00FFFF").grid(row=11,column=0,sticky="w")

        global second_kind_info6
        second_kind_info6=tkinter.Entry(bpp,font=("楷体",12),bg="#D3D3D3")
        second_kind_info6.bind("<KeyRelease>",Entry_get.second_kind_get_data6)
        second_kind_info6.grid(row=11,column=1,sticky="w")
        tkinter.Label(bpp,text="").grid(row=12,column=0,columnspan=4,sticky="w")
        tkinter.Button(bpp,text="正常写入保存",font=("楷体",15),bg="#48D1CC",
                       command=lambda:Save_data(self.course_name).Save_data_second_kind()).grid(row=13,column=0,rowspan=3,sticky="w")
        tkinter.Button(bpp,text="修改",font=("楷体",15),bg="#FFE4B5",
                       command=lambda:Revise(self.course_name).Revise_kind_2()).grid(row=13,column=2,rowspan=3,sticky="w")

        tkinter.Button(bpp,text="导入查询",font=("楷体",15),bg="#F0F8FF",
                       command=lambda:Info_look(self.course_name).Look_sk_info_Button()).grid(row=13,column=3,rowspan=3,sticky="w")
        tkinter.Button(bpp,text="录入下一个学生信息",font=("楷体",15),bg="#F0F8FF",
                       command=lambda:Save_continue(self.course_name).Save_next_second_kind()).grid(row=13,column=1,rowspan=3,sticky="w")
        tkinter.Button(bpp,text="学生成绩\n的分析系统",font=("楷体",18),bg="#EE82EE",height=4,
                       command=lambda:Course_Widget(self.course_name).All_command_three(),relief="raised").grid(row=1,column=2,rowspan=4,columnspan=2,sticky="w")
        tkinter.Button(bpp,text="<<----返回\n当前课程界面",font=("楷体",15),bg="#008B8B",
                       command=lambda:Back_Turn(self.course_name).Back_second_kind()).grid(row=17,column=0,rowspan=2,columnspan=2,sticky="w")
        tkinter.Button(bpp,text="<<----返回\n总课程界面",font=("楷体",15),bg="#008B8B",
                       command=lambda:Back_Turn(self.course_name).Back_third_kind()).grid(row=17,column=2,rowspan=2,columnspan=2,sticky="w")
        tkinter.Button(bpp,text="删除一个学生信息",font=("楷体",15),bg="#00FF7F",
                       command=lambda:Delete(self.course_name).Delete_sktype()).grid(row=11,column=2,columnspan=2,sticky="w")
        tkinter.Button(bpp,text="信息导入\n修改的保存",font=("楷体",18),bg="#FF69B4",
                       command=lambda:Save_data(self.course_name).Insert_Revise_Save_sk()).grid(row=6,column=2,rowspan=3,columnspan=2,sticky="w")
        bpp.mainloop()


    def All_command_three(self):          #学生成绩处理系统的组件
        BackError(bpp)
        global aware2
        aware2=tkinter.Tk()
        aware2.title("<<--温馨提醒-->>")
        tkinter.Label(aware2,text="欢迎进入学生成绩分析系统\n如果你要进行学生学期总成绩\n的分析且学生信息的相关excel\n表格没有总成绩一栏"
                      +"\n（本界面有打开相应excel所在文件夹的按钮,\n方便查看，看完后请务必关闭所有excel，\n否则会产生进程干扰，跳出警示窗口!）"
                      +"\n请务必点击权重菜单，选择相应的\n权重（小数），系统会自动计算每个学生的\n总成绩并保存(修改相同操作)，否则会提示操作错误!",
                      fg="blue",font=("楷体",12),bg="#E0FFFF").grid(row=0)
        tkinter.Button(aware2,text="点击退出",font=("楷体",12),bg="green",command=lambda:Exit_aware(aware2)).grid(row=1)

        global cpp
        cpp=tkinter.Tk()
        cpp.title("《%s》学生成绩分析系统"%(self.course_name))
        cpp.geometry("500x400")
        #**********************------->>>>>>菜单栏目的布置<<<<<<<-------********************#
        menubar=tkinter.Menu(cpp)
        GUI_menu1=tkinter.Menu(menubar,tearoff=0,activebackground="green",font=("楷体",12))
        GUI_menu1.add_command(label="总成绩的权重输入",command=lambda:Analyse_system(self.course_name).Quan_Input())

        GUI_menu1.add_separator()
        GUI_menu1.add_command(label="打开学生信息所在文件夹",command=lambda:Analyse_system(self.course_name).Open_Dir())

        GUI_menu1.add_separator()
        GUI_menu1.add_command(label="打开本系统的说明文件",command=Use_description)

        GUI_menu1.add_separator()
        GUI_menu1.add_command(label="返回总课程界面",command=lambda:Back_Turn(self.course_name).Back_fourth_kind())
        
        menubar.add_cascade(label="菜单1(important)", menu=GUI_menu1)

        GUI_menu2=tkinter.Menu(menubar,tearoff=0,activebackground="pink",font=("楷体",12))
        GUI_menu2.add_command(label="全体学生成绩排序",command=lambda:Analyse_system(self.course_name).Grades_sorted())
        GUI_menu2.add_separator()

        GUI_menu2.add_command(label="打开学生信息所在文件夹",command=lambda:Analyse_system(self.course_name).Open_Dir())
        GUI_menu2.add_separator()

        GUI_menu2.add_command(label="特定教师的学生成绩统计图分析",command=lambda:Analyse_system(self.course_name).Analyse_single_for_choose())
        GUI_menu2.add_separator()

        GUI_menu2.add_command(label="课程全体学生的成绩统计图分析",command=lambda:Analyse_system(self.course_name).Analyse_all())
        menubar.add_cascade(label="菜单2(significant)",menu=GUI_menu2)
        
        cpp.config(menu=menubar)
        cpp.mainloop()
        
        



#Entry组件输入内容的获取 
class Entry_get:
    def first_kind_get_data1(event):
        dict_globals=globals()
        if "fkoutcome1" not in dict_globals:
            global fkoutcome1
        fkoutcome1=first_kind_info1.get()
    
        
    def first_kind_get_data2(event):
        global fkoutcome2
        fkoutcome2=first_kind_info2.get()
    
  
    def first_kind_get_data3(event):
        global fkoutcome3
        fkoutcome3=first_kind_info3.get()


    def second_kind_get_data1(event):     
        global skoutcome1
        skoutcome1=second_kind_info1.get()
    
        
    def second_kind_get_data2(event):     
        global skoutcome2
        skoutcome2=second_kind_info2.get()


    def second_kind_get_data3(event):     
        global skoutcome3
        skoutcome3=second_kind_info3.get()    


    def second_kind_get_data4(event):     
        global skoutcome4
        skoutcome4=second_kind_info4.get()
    

    def second_kind_get_data5(event):     
        global skoutcome5
        skoutcome5=second_kind_info5.get()


    def second_kind_get_data6(event):     
        global skoutcome6
        skoutcome6=second_kind_info6.get()


    def look1_info_get(event):
        global look1_keywords
        look1_keywords=look1_info.get()


    def look2_info_get_1(event):
        global look2_keywords_1
        look2_keywords_1=look2_info_1.get()


    def look2_info_get_2(event):
        global look2_keywords_2
        look2_keywords_2=look2_info_2.get()


    def Delete1_teacher_get(event):
        global teacher_name
        teacher_name=teacher.get()


    def Delete2_tch_get(event):
        global tch_name
        tch_name=tch.get()
    

    def Delete2_stu_get(event):
        global stu_id
        stu_id=stu.get()


    
#信息的保存类
class Save_data:
    def __init__(self,course_name):
        self.course_name=course_name
    def Save_data_first_kind(self):
        #***************************************************
        '''#课程基本信息的保存                                  *
        #内加错误处理，防止有输入框为空内容而直接保存导致程序报错。   *
        #利用os模块判断文件夹是否存在，不存在的话利用os模块创建 '''  
        #**************************************************   *
        try:
            if os.path.exists("E:\课程管理系统信息保存\《%s》课程\课程教师信息汇总.xlsx"%(self.course_name))==False:
                os.makedirs(r"E:\课程管理系统信息保存\《%s》课程"%(self.course_name))
                w=xlwt.Workbook()                                     #判断表格或者文件夹存不存在，不存在创建，用xlwt设置单元格宽度
                sht=w.add_sheet("sheet1")
                sht.col(2).width=9999
                w.save("E:/课程管理系统信息保存/《%s》课程/课程教师信息汇总.xlsx"%(self.course_name))

            if (fkoutcome1=="" or fkoutcome1.isspace()==True) or (fkoutcome2=="" or
                    fkoutcome2.isspace()==True) or (fkoutcome3 =="" or fkoutcome3.isspace()==True):
                a=1/0                    ######强行执行报错语句（目的是为了排除在输入框里写入了空白字符串从而跳过后面的错误处理）
            else:
                first_kind_info1["state"]="readonly"
                first_kind_info2["state"]="readonly"
                first_kind_info3["state"]="readonly"

            xlApp=win32com.client.Dispatch("Excel.Application")
            xlBook1=xlApp.Workbooks.Open("E:/课程管理系统信息保存/《%s》课程/课程教师信息汇总.xlsx"%(self.course_name))
            xlSht1=xlBook1.Worksheets("sheet1")
            xlSht1.Cells(1,1).Value="任课教师"
            xlSht1.Cells(1,2).Value="上课地点"
            xlSht1.Cells(1,3).Value="上课时间"
            row1=2
            while xlSht1.Cells(row1,1).Value not in [None,""]:
                row1=row1+1
            for i in range(2,row1+1):
                if fkoutcome1==xlSht1.Cells(i,1).Value:
                    xlSht1.Cells(i,2).Value=fkoutcome2
                    xlSht1.Cells(i,3).Value=fkoutcome3
                    break
                else:
                    if i==row1:
                        xlSht1.Cells(i,1).Value=fkoutcome1
                        xlSht1.Cells(i,2).Value=fkoutcome2
                        xlSht1.Cells(i,3).Value=fkoutcome3
            xlBook1.Close(SaveChanges=1)
            del xlApp    
        except Exception:
            BackError(app)      #跳转关闭异常处理(函数调用)
            global Error1
            Error1=tkinter.Tk()
            tkinter.Button(Error1,text="对不起，有输入框\n未进行写入操作,请\n写入内容(点击进入\n课程信息录入界面)",font=("楷体",12),bg="#00FFFF",width=20,height=20,
                           command=lambda:Course_Widget(self.course_name).All_command_one()).grid()
            Error1.mainloop()

    def Save_data_second_kind(self):   
        #***************************************************
        '''#课程基本信息的保存                                  *
        #内加错误处理，防止有输入框为空内容而直接保存导致程序报错。   *
        #利用os模块判断文件夹是否存在，不存在的话利用os模块创建 '''  
        #**************************************************   *
        try:
            if os.path.exists("E:\课程管理系统信息保存\《%s》课程\教师%s的学生信息.xlsx"%(self.course_name,fkoutcome1))==False:
                w1=xlwt.Workbook()
                sht1=w1.add_sheet("sheet1")
                w1.save("E:/课程管理系统信息保存/《%s》课程/教师%s的学生信息.xlsx"%(self.course_name,fkoutcome1))
                                           #判断表格或者文件夹存不存在，不存在创建

            if (skoutcome1=="" or skoutcome1.isspace()==True) or (skoutcome2=="" or skoutcome2.isspace()==True) or(skoutcome3==""
                or skoutcome3.isspace()==True) or (skoutcome4=="" or skoutcome4.isspace()==True) or(skoutcome5==""
                or skoutcome5.isspace()==True) or (skoutcome6=="" or skoutcome6.isspace()==True):
                a=1/0                ######强行执行报错语句（目的是为了排除在输入框里写入了空白字符串从而跳过后面的错误处理）
            second_kind_info1["state"]="readonly"
            second_kind_info2["state"]="readonly"
            second_kind_info3["state"]="readonly"
            second_kind_info4["state"]="readonly"
            second_kind_info5["state"]="readonly"
            second_kind_info6["state"]="readonly"

            xlApp1=win32com.client.Dispatch("Excel.Application")
            xlBook2=xlApp1.Workbooks.Open("E:/课程管理系统信息保存/《%s》课程/教师%s的学生信息.xlsx"%(self.course_name,fkoutcome1))
            xlSht2=xlBook2.Worksheets("sheet1")
            xlSht2.Cells(1,1).Value="学生姓名"
            xlSht2.Cells(1,2).Value="学生学号"
            xlSht2.Cells(1,3).Value="所在班级"
            xlSht2.Cells(1,4).Value="平时成绩"
            xlSht2.Cells(1,5).Value="期中成绩"
            xlSht2.Cells(1,6).Value="期末成绩"
            row2=2
            while xlSht2.Cells(row2,1).Value not in [None,""]:
                row2=row2+1
            for i in range(2,row2+1):
                if skoutcome1==xlSht2.Cells(i,1).Value:
                    xlSht2.Cells(i,2).Value=skoutcome2
                    xlSht2.Cells(i,3).Value=skoutcome3
                    xlSht2.Cells(i,4).Value=skoutcome4
                    xlSht2.Cells(i,5).Value=skoutcome5
                    xlSht2.Cells(i,6).Value=skoutcome6
                    break
                else:
                    if i==row2:
                        xlSht2.Cells(i,1).Value=skoutcome1
                        xlSht2.Cells(i,2).Value=skoutcome2
                        xlSht2.Cells(i,3).Value=skoutcome3
                        xlSht2.Cells(i,4).Value=skoutcome4
                        xlSht2.Cells(i,5).Value=skoutcome5
                        xlSht2.Cells(i,6).Value=skoutcome6
            xlBook2.Close(SaveChanges=1)
            del xlApp1
        except Exception:
            BackError(bpp)       #跳转关闭异常处理(函数调用)
            global Error2
            Error2=tkinter.Tk()
            tkinter.Button(Error2,text="对不起，有输入框内\n容为空，或者前一界面\n教师信息未导入\n点击重新操作",
                           font=("楷体",12),bg="#00FFFF",width=20,height=20,
                           command=lambda:Course_Widget(self.course_name).All_command_two()).grid()
            Error2.mainloop()


    def Insert_Revise_Save_fk(self):
        try:
            dict_globals=globals()
            list_globals=dict_globals.keys()           #没有动的输入框内容赋值给相应的全局变量
            check_list=["fkoutcome2","fkoutcome3"]
            fkoutcome2=""
            fkoutcome3=""
            for i in check_list:
                if i not in list_globals:
                    if i=="fkoutcome2":
                        fkoutcome2=dict1["上课地点"]
                    elif i=="fkoutcome3":
                        fkoutcome3=dict1["上课时间"]
                else:
                    if i=="fkoutcome2":
                        fkoutcome2=dict_globals[i]
                    elif i=="fkoutcome3":
                        fkoutcome3=dict_globals[i]
                    
            try:
                if (fkoutcome1=="" or fkoutcome1.isspace()==True) or (fkoutcome2=="" or
                    fkoutcome2.isspace()==True) or (fkoutcome3 =="" or fkoutcome3.isspace()==True):
                    a=1/0                    ######强行执行报错语句（目的是为了排除内容修改为空白字符串从而跳过后面的错误处理）
                else:
                    first_kind_info1["state"]="readonly"
                    first_kind_info2["state"]="readonly"
                    first_kind_info3["state"]="readonly"

                xlApp=win32com.client.Dispatch("Excel.Application")
                xlBook1=xlApp.Workbooks.Open("E:/课程管理系统信息保存/《%s》课程/课程教师信息汇总.xlsx"%(self.course_name))
                xlSht1=xlBook1.Worksheets("sheet1")
                xlSht1.Cells(1,1).Value="任课教师"
                xlSht1.Cells(1,2).Value="上课地点"
                xlSht1.Cells(1,3).Value="上课时间"
                row1=2
                while xlSht1.Cells(row1,1).Value not in [None,""]:
                    row1=row1+1
                for i in range(2,row1+1):
                    if fkoutcome1==xlSht1.Cells(i,1).Value:
                        xlSht1.Cells(i,2).Value=fkoutcome2
                        xlSht1.Cells(i,3).Value=fkoutcome3
                        break
                xlBook1.Close(SaveChanges=1)
                del xlApp    
            except Exception:
                BackError(app)      #跳转关闭异常处理(函数调用)
                global Error21
                Error21=tkinter.Tk()
                tkinter.Button(Error21,text="对不起，有输入框\n未进行写入操作,请\n写入内容(点击进入\n课程信息录入界面)",font=("楷体",12),bg="#00FFFF",width=20,height=20,
                               command=lambda:Course_Widget(self.course_name).All_command_one()).grid()
                Error21.mainloop()
        except Exception:
            BackError(app)
            global Error5
            Error5=tkinter.Tk()
            tkinter.Button(Error5,text="信息未导入\n点击返回",font=("楷体",12),bg="#00FFFF",width=20,height=20,
                           command=lambda:Course_Widget(self.course_name).All_command_one()).grid()
            Error5.mainloop()


    def Insert_Revise_Save_sk(self):
        try:
            dict_globals=globals()
            list_globals=dict_globals.keys()           #没有动的输入框内容赋值给相应的全局变量
            check_list=["skoutcome1","skoutcome2","skoutcome3","skoutcome4","skoutcome5","skoutcome6"]
            skoutcome1=""
            skoutcome2=""
            skoutcome3=""
            skoutcome4=""
            skoutcome5=""
            skoutcome6=""
            for i in check_list:
                if i not in list_globals:
                    if i=="skoutcome1":
                        skoutcome1=dict2["学生姓名"]
                    elif i=="skoutcome2":
                        skoutcome2=dict2["学生学号"]
                    elif i=="skoutcome3":
                        skoutcome3=dict2["所在班级"]
                    elif i=="skoutcome4":
                        skoutcome4=dict2["平时成绩"]
                    elif i=="skoutcome5":
                        skoutcome5=dict2["期中成绩"]
                    elif i=="skoutcome6":
                        skoutcome6=dict2["期末成绩"]
                else:
                    if i=="skoutcome1":
                        skoutcome1=dict_globals[i]
                    elif i=="skoutcome2":
                        skoutcome2=dict_globals[i]
                    elif i=="skoutcome3":
                        skoutcome3=dict_globals[i]
                    elif i=="skoutcome4":
                        skoutcome4=dict_globals[i]
                    elif i=="skoutcome5":
                        skoutcome5=dict_globals[i]
                    elif i=="skoutcome6":
                        skoutcome6=dict_globals[i]
            try:
                if (skoutcome1=="" or str(skoutcome1).isspace()==True) or (skoutcome2=="" or str(skoutcome2).isspace()==True) or(skoutcome3==""
                    or str(skoutcome3).isspace()==True) or (skoutcome4=="" or str(skoutcome4).isspace()==True) or(skoutcome5==""
                    or str(skoutcome5).isspace()==True) or (skoutcome6=="" or str(skoutcome6).isspace()==True):
                    a=1/0                ######强行执行报错语句（目的是为了排除在输入框里写入了空白字符串从而跳过后面的错误处理

                second_kind_info1["state"]="readonly"
                second_kind_info2["state"]="readonly"
                second_kind_info3["state"]="readonly"
                second_kind_info4["state"]="readonly"
                second_kind_info5["state"]="readonly"
                second_kind_info6["state"]="readonly"

                xlApp1=win32com.client.Dispatch("Excel.Application")
                xlBook2=xlApp1.Workbooks.Open("E:/课程管理系统信息保存/《%s》课程/教师%s的学生信息.xlsx"%(self.course_name,fkoutcome1))
                xlSht2=xlBook2.Worksheets("sheet1")
                xlSht2.Cells(1,1).Value="学生姓名"
                xlSht2.Cells(1,2).Value="学生学号"
                xlSht2.Cells(1,3).Value="所在班级"
                xlSht2.Cells(1,4).Value="平时成绩"
                xlSht2.Cells(1,5).Value="期中成绩"
                xlSht2.Cells(1,6).Value="期末成绩"
                row2=2
                while xlSht2.Cells(row2,1).Value not in [None,""]:
                    row2=row2+1
                for i in range(2,row2+1):
                    if skoutcome1==xlSht2.Cells(i,1).Value:
                        xlSht2.Cells(i,2).Value=int(skoutcome2)
                        xlSht2.Cells(i,3).Value=skoutcome3
                        xlSht2.Cells(i,4).Value=skoutcome4
                        xlSht2.Cells(i,5).Value=skoutcome5
                        xlSht2.Cells(i,6).Value=skoutcome6
                        break
                xlBook2.Close(SaveChanges=1)
                del xlApp1
            except Exception:
                BackError(bpp)       #跳转关闭异常处理(函数调用)
                global Error22
                Error22=tkinter.Tk()
                tkinter.Button(Error22,text="对不起，有输入框内\n容为空，或者前一界面\n教师信息未导入\n点击重新操作",
                               font=("楷体",12),bg="#00FFFF",width=20,height=20,
                               command=lambda:Course_Widget(self.course_name).All_command_two()).grid()
                Error22.mainloop()

        except Exception:
            BackError(bpp)
            global Error6
            Error6=tkinter.Tk()
            tkinter.Button(Error6,text="信息未导入\n点击返回",font=("楷体",12),bg="#00FFFF",width=20,height=20,
                           command=lambda:Course_Widget(self.course_name).All_command_two()).grid()
            Error6.mainloop()


    def Quan_Save(self):
        try:
            if (Amodel.get()=="") or (Bmodel.get()=="") or(Cmodel.get()==""):
                a=1/0                #强行报错处理，用于判断权重是不是都选择了
            dir_list=os.listdir("E:\课程管理系统信息保存\《%s》课程"%(self.course_name))
            dir_list.remove("课程教师信息汇总.xlsx")
            stu_list=dir_list
            xlApp=win32com.client.Dispatch("Excel.Application")
            
            for i in stu_list:
                xlBook=xlApp.Workbooks.Open("E:/课程管理系统信息保存/《%s》课程/%s"%(self.course_name,i))
                xlSht=xlBook.Worksheets("sheet1")
                xlSht.Cells(1,7).Value="总成绩"
                rowbin=2
                while xlSht.Cells(rowbin,1).Value not in [None,""]:
                    rowbin=rowbin+1
                for j in range(2,rowbin):
                    xlSht.Cells(j,7).Value=round(eval(str(xlSht.Cells(j,4).Value)+"*"+str(Amodel.get())+"+"
                        +str(xlSht.Cells(j,5).Value)+"*"+str(Bmodel.get())+"+"
                        +str(xlSht.Cells(j,6).Value)+"*"+str(Cmodel.get())),3)
                xlBook.Close(SaveChanges=1)
            del xlApp
            Quan.destroy()
        except Exception:
            global Error7
            Error7=tkinter.Tk()
            tkinter.Button(Error7,text="操作错误\n点击返回",font=("楷体",12),bg="#00FFFF",width=20,height=20,
                           command=lambda:Exit_aware(Error7)).grid()
            Error7.mainloop()



#信息的继续录入
class Save_continue:
    def __init__(self,course_name):
        self.course_name=course_name
    def Save_next_first_kind(self):
        BackError(app)     #跳转关闭异常处理(函数调用)
        Course_Widget(self.course_name).All_command_one()
    def Save_next_second_kind(self):
        BackError(bpp)     #跳转关闭异常处理(函数调用)
        Course_Widget(self.course_name).All_command_two()

'''#****************************************************************************************************************************
#****************************************************************************************************************************
#^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
#^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^'''


#界面返回跳转
class Back_Turn:
    def __init__(self,course_name):
        self.course_name=course_name
    def Back_first_kind(self):
        a=self.course_name
        BackError(app)     #跳转关闭异常处理(函数调用)
        main()
    def Back_second_kind(self):
        BackError(bpp)     
        Course_Widget(self.course_name).All_command_one()
    def Back_third_kind(self):
        a=self.course_name
        BackError(bpp)     
        main()
    def Back_fourth_kind(self):
        a=self.course_name
        BackError(cpp)
        main()


#信息查询导入类
class Info_look:
    def __init__(self,course_name):
        self.course_name=course_name
    def Look_fk_info_Button(self):
        try:
            BackError(Error3_1)    ##跳转关闭异常处理(函数调用)
        except Exception:
            pass
        global look1
        look1=tkinter.Tk()
        look1.title("课程基本信息查询界面")
        tkinter.Label(look1,text="输入所需查询\n课程老师姓名：",font=("楷体",15)).grid(row=0,column=0)
        tkinter.Button(look1,text="点击查询",bg="#E0FFFF",font=("楷体",15),
                   command=lambda:Info_look(self.course_name).look1_fk_info_Search()).grid(row=1)
        global look1_info
        look1_info=tkinter.Entry(look1,bg="pink",font=("楷体",15),relief="sunken")
        look1_info.bind("<KeyRelease>",Entry_get.look1_info_get)
        look1_info.grid(row=0,column=1)
        look1.mainloop()
    def look1_fk_info_Search(self):
        BackError(look1)        #跳转关闭异常处理(函数调用)
        try:
            if os.path.exists("E:\课程管理系统信息保存\《%s》课程\课程教师信息汇总.xlsx"%(self.course_name))==True:
                BackError(look1)
                global dict1
                dict1={}       #创建一个空字典
                xlApp=win32com.client.Dispatch("Excel.Application")
                xlBook=xlApp.Workbooks.Open("E:/课程管理系统信息保存/《%s》课程/课程教师信息汇总.xlsx"%(self.course_name))
                xlSht=xlBook.Worksheets("sheet1")
                row=2
                while xlSht.Cells(row,1).Value not in [None,""]:
                    row=row+1
                for i in range(2,row):
                    if xlSht.Cells(i,1).Value==look1_keywords:
                        dict1["任课教师"]=xlSht.Cells(i,1).Value
                        dict1["上课地点"]=xlSht.Cells(i,2).Value
                        dict1["上课时间"]=xlSht.Cells(i,3).Value
                        break
                    else:
                        if i==row-1:
                            del xlApp
                            a=1/0           #强行报错处理

                del xlApp
                a1_model.set(dict1["任课教师"])
                b1_model.set(dict1["上课地点"])
                c1_model.set(dict1["上课时间"])

                first_kind_info1["textvariable"]=a1_model
                first_kind_info2["textvariable"]=b1_model
                first_kind_info3["textvariable"]=c1_model

                first_kind_info1["state"]="readonly"
                first_kind_info2["state"]="readonly"
                first_kind_info3["state"]="readonly"
                dict_globals=globals()
                if "fkoutcome1" not in dict_globals:
                    global fkoutcome1
                fkoutcome1=look1_keywords
            else:
                app.destroy()
                global Error3_2
                Error3_2=tkinter.Tk()    
                tkinter.Button(Error3_2,text="对不起，本课程\n未录入任何信息\n点击返回课程\n信息录入界面!",font=("楷体",12),bg="#00FFFF",width=20,height=20,
                               command=lambda:Course_Widget(self.course_name).All_command_one()).grid()
                Error3_2.mainloop()
                

        except Exception:                      #未输入教师名称报错（未进行写入操作）
            global Error3_1
            Error3_1=tkinter.Tk()    
            tkinter.Button(Error3_1,text="对不起，教师姓名\n不存在或者未\n输入教师姓名\n(点击重新查询)",font=("楷体",12),bg="#00FFFF",width=20,height=20,
                           command=lambda:Info_look(self.course_name).Look_fk_info_Button()).grid()
            Error3_1.mainloop()

    def Look_sk_info_Button(self):
        try:    
            BackError(Error4_1)    ##跳转关闭异常处理(函数调用)
        except Exception:
            pass
        try:
            BackError(Error4_2)    ##跳转关闭异常处理(函数调用)
        except Exception:
            pass
        global look2
        look2=tkinter.Tk()
        look2.title("学生信息查询页面")
        tkinter.Label(look2,text="请输入教师姓名：",font=("楷体",15)).grid(row=0,column=0)
        tkinter.Label(look2,text="请输入学生学号：",font=("楷体",15)).grid(row=1,column=0)
        tkinter.Button(look2,text="点击查询",bg="#E0FFFF",font=("楷体",15),
                       command=lambda:Info_look(self.course_name).look2_sk_info_Search()).grid(row=2)
        global look2_info_1
        global look2_info_2
        look2_info_1=tkinter.Entry(look2,bg="pink",font=("楷体",15),relief="sunken")
        look2_info_1.bind("<KeyRelease>",Entry_get.look2_info_get_1)
        look2_info_1.grid(row=0,column=1)
        look2_info_2=tkinter.Entry(look2,bg="pink",font=("楷体",15),relief="sunken")
        look2_info_2.bind("<KeyRelease>",Entry_get.look2_info_get_2)
        look2_info_2.grid(row=1,column=1)
        look2.mainloop()

    def look2_sk_info_Search(self):
        BackError(look2)        #跳转关闭异常处理(函数调用)
        try:
            if os.path.exists("E:\课程管理系统信息保存\《%s》课程\教师%s的学生信息.xlsx"%(self.course_name,look2_keywords_1))==True:
                BackError(look2)
                global dict2
                dict2={}       #创建一个空字典
                xlApp=win32com.client.Dispatch("Excel.Application")
                xlBook=xlApp.Workbooks.Open("E:/课程管理系统信息保存/《%s》课程/教师%s的学生信息.xlsx"%(self.course_name,look2_keywords_1))
                xlSht=xlBook.Worksheets("sheet1")
                row=2
                while xlSht.Cells(row,1).Value not in [None,""]:
                    row=row+1
                for i in range(2,row):
                    if xlSht.Cells(i,2).Value==int(look2_keywords_2):
                        dict2["学生姓名"]=xlSht.Cells(i,1).Value
                        dict2["学生学号"]=int(xlSht.Cells(i,2).Value)
                        dict2["所在班级"]=xlSht.Cells(i,3).Value
                        dict2["平时成绩"]=xlSht.Cells(i,4).Value
                        dict2["期中成绩"]=xlSht.Cells(i,5).Value
                        dict2["期末成绩"]=xlSht.Cells(i,6).Value
                        break
                    else:
                        if i==row-1:
                            del xlApp
                            a=1/0           #强行报错

                del xlApp
                a2_model.set(dict2["学生姓名"])
                b2_model.set(dict2["学生学号"])
                c2_model.set(dict2["所在班级"])
                d2_model.set(dict2["平时成绩"])
                e2_model.set(dict2["期中成绩"])
                f2_model.set(dict2["期末成绩"])

                second_kind_info1["textvariable"]=a2_model
                second_kind_info2["textvariable"]=b2_model
                second_kind_info3["textvariable"]=c2_model
                second_kind_info4["textvariable"]=d2_model
                second_kind_info5["textvariable"]=e2_model
                second_kind_info6["textvariable"]=f2_model

                second_kind_info1["state"]="readonly"
                second_kind_info2["state"]="readonly"
                second_kind_info3["state"]="readonly"
                second_kind_info4["state"]="readonly"
                second_kind_info5["state"]="readonly"
                second_kind_info6["state"]="readonly"
                dict_globals=globals()
                if "fkoutcome1" not in dict_globals:
                    global fkoutcome1
                fkoutcome1=look2_keywords_1
                
            else:
                global Error4_2                   #课程教师姓名不存在
                Error4_2=tkinter.Tk()
                tkinter.Button(Error4_2,text="对不起,教师不\n存在或你输入\n了空字符串\n,h\或者本课程未\n录入任何信息\n(点击重新查询)",
                            font=("楷体",12),bg="#00FFFF",width=20,height=20,
                            command=lambda:Info_look(self.course_name).Look_sk_info_Button()).grid()
                Error4_2.mainloop()

        except Exception:                      #未输入学号直接查询报错处理
            global Error4_1
            Error4_1=tkinter.Tk()
            tkinter.Button(Error4_1,text="未输入教师姓名\n或未输入学生学号\n或者该学生学号不存在\n(点击重新查询)",font=("楷体",12),bg="#00FFFF",width=20,height=20,
                          command=lambda:Info_look(self.course_name).Look_sk_info_Button()).grid()
            Error4_1.mainloop()



#信息的修改类    
class Revise:
    def __init__(self,course_name):
        self.course_name=course_name

    def Revise_kind_1(self):
        first_kind_info1["state"]="readonly"
        first_kind_info2["state"]="normal"
        first_kind_info3["state"]="normal"
        
    def Revise_kind_2(self):
        second_kind_info1["state"]="normal"
        second_kind_info2["state"]="normal"
        second_kind_info3["state"]="normal"
        second_kind_info4["state"]="normal"
        second_kind_info5["state"]="normal"
        second_kind_info6["state"]="normal"

'''#****************************************************************************************************************************
#****************************************************************************************************************************
#^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
#^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^'''

#信息的删除类
class Delete:
    def __init__(self,course_name):
        self.course_name=course_name
    
    def Delete_fktype(self):
        try:
            BackError(Error11)      ##跳转关闭异常处理(函数调用)
        except Exception:
            pass
        try:
            BackError(Error11_1)    ##跳转关闭异常处理(函数调用)
        except Exception:
            pass

        global delete1
        delete1=tkinter.Tk()
        tkinter.Label(delete1,text="输入所需删除\n课程老师姓名：",font=("楷体",15)).grid(row=0,column=0)
        tkinter.Button(delete1,text="点击删除",bg="#E0FFFF",font=("楷体",15),
                   command=lambda:Delete(self.course_name).Delete1_do()).grid(row=1)
        global teacher
        teacher=tkinter.Entry(delete1,bg="pink",font=("楷体",15),relief="sunken")
        teacher.bind("<KeyRelease>",Entry_get.Delete1_teacher_get)
        teacher.grid(row=0,column=1)
        delete1.mainloop()

    def Delete1_do(self):
        BackError(delete1)        #跳转关闭异常处理(函数调用)
        try:
            if os.path.exists("E:\课程管理系统信息保存\《%s》课程\课程教师信息汇总.xlsx"%(self.course_name))==True:
                xlApp=win32com.client.Dispatch("Excel.Application")
                xlBook=xlApp.Workbooks.Open("E:/课程管理系统信息保存/《%s》课程/课程教师信息汇总.xlsx"%(self.course_name))
                xlSht=xlBook.Worksheets("sheet1")
                row=2
                while xlSht.Cells(row,1).Value not in [None,""]:
                    row=row+1
                for i in range(2,row):
                    if xlSht.Cells(i,1).Value==teacher_name:
                        xlSht.Rows(i).Delete()
                        xlBook.Close(SaveChanges=1)
                        del xlApp
                        break
                    else:
                        if i==row-1:
                            del xlApp
                            global Error11                   #教师名称不存在或者输入的是空字符串
                            Erro11=tkinter.Tk()
                            tkinter.Button(Error11,text="对不起，该教师\n不存在或者\n你输入了空字符串\n(点击重新操作)",
                                        font=("楷体",12),bg="#00FFFF",width=20,height=20,
                                        command=lambda:Delete(self.course_name).Delete_fktype()).grid()
                            Error11.mainloop()
            if os.path.exists("E:\课程管理系统信息保存\《%s》课程\教师%s的学生信息.xlsx"%(self.course_name,teacher_name)):
                os.unlink("E:\课程管理系统信息保存\《%s》课程\教师%s的学生信息.xlsx"%(self.course_name,teacher_name))
            else:
                pass

        except Exception:                         #未输入教师名称报错（未进行写入操作）
            global Error11_1
            Error11_1=tkinter.Tk()    
            tkinter.Button(Error11_1,text="未输入教师名称\n(点击重新操作)",font=("楷体",12),bg="#00FFFF",width=20,height=20,
                           command=lambda:Delete(self.course_name).Delete_fktype()).grid()
            Error11_1.mainloop()


    def Delete_sktype(self):
        try:
            BackError(Error12)      ##跳转关闭异常处理(函数调用)
        except Exception:
            pass
        try:
            BackError(Error12_1)    ##跳转关闭异常处理(函数调用)
        except Exception:
            pass
        try:
            BackError(Error12_2)    ##跳转关闭异常处理(函数调用)
        except Exception:
            pass

        global delete2
        delete2=tkinter.Tk()
        tkinter.Label(delete2,text="输入所需删除\n学生的老师姓名：",font=("楷体",15)).grid(row=0,column=0)
        tkinter.Button(delete2,text="点击删除",bg="#E0FFFF",font=("楷体",15),
                       command=lambda:Delete(self.course_name).Delete2_do()).grid(row=2)
        tkinter.Label(delete2,text="输入所需删除\n学生学号：",font=("楷体",15)).grid(row=1,column=0)
    
        global tch
        tch=tkinter.Entry(delete2,bg="pink",font=("楷体",15),relief="sunken")
        tch.bind("<KeyRelease>",Entry_get.Delete2_tch_get)
        tch.grid(row=0,column=1)

        global stu
        stu=tkinter.Entry(delete2,bg="pink",font=("楷体",15),relief="sunken")
        stu.bind("<KeyRelease>",Entry_get.Delete2_stu_get)
        stu.grid(row=1,column=1)
        delete2.mainloop()
    

    def Delete2_do(self):
        BackError(delete2)        #跳转关闭异常处理(函数调用)
        try:
            if os.path.exists("E:\课程管理系统信息保存\《%s》课程\教师%s的学生信息.xlsx"%(self.course_name,tch_name))==True:
                xlApp=win32com.client.Dispatch("Excel.Application")
                xlBook=xlApp.Workbooks.Open("E:/课程管理系统信息保存/《%s》课程/教师%s的学生信息.xlsx"%(self.course_name,tch_name))
                xlSht=xlBook.Worksheets("sheet1")
                row=2
                while xlSht.Cells(row,1).Value not in [None,""]:
                    row=row+1
                for i in range(2,row):
                    if xlSht.Cells(i,2).Value==int(stu_id):
                        xlSht.Rows(i).Delete()
                        xlBook.Close(SaveChanges=1)
                        del xlApp
                        break
                    else:
                        if i==row-1:
                            del xlApp
                            global Error12                   #教师名称不存在或者学生学号不存在或者输入的是空字符串
                            Erro12=tkinter.Tk()
                            tkinter.Button(Error12,text="对不起,学生学号\n不存在或你输\n入了空字符串\n(点击重新操作)",
                                        font=("楷体",12),bg="#00FFFF",width=20,height=20,
                                        command=lambda:Delete(self.course_name).Delete_sktype()).grid()
                            Error12.mainloop()
            else:
                global Error12_2                   #课程教师姓名不存在
                Error12_2=tkinter.Tk()
                tkinter.Button(Error12_2,text="对不起,教师不\n存在或你输入\n了空字符串\n(点击重新操作)",
                            font=("楷体",12),bg="#00FFFF",width=20,height=20,
                            command=lambda:Info_look(self.course_name).Look_sk_info_Button()).grid()
                Error12_2.mainloop()

        except Exception:                         #未输入教师名称或者学生学号报错（未进行写入操作）
            global Error12_1
            Error12_1=tkinter.Tk()    
            tkinter.Button(Error12_1,text="未输入教师名或学号\n(点击重新操作)",font=("楷体",12),bg="#00FFFF",width=20,height=20,
                           command=lambda:Delete(self.course_name).Delete_sktype()).grid()
            Error12_1.mainloop()

'''#****************************************************************************************************************************
#****************************************************************************************************************************
#^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
#^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^'''

#学生成绩的分析处理
class Analyse_system:
    def __init__(self,course_name):
        self.course_name=course_name

    def Quan_Input(self):        #权重输入窗口（OptionMenu）
        global Quan
        Quan=tkinter.Tk()
        Quan.title("选择权重")
        Option_number=["0.05","0.10","0.15","0.20","0.25","0.30","0.35","0.40","0.45","0.50","0.55","0.60","0.65",
                   "0.70","0.75","0.80","0.85","0.90","0.95","1.00"]

        global Amodel
        Amodel=tkinter.StringVar()
        tkinter.Label(Quan,text="平时成绩权重:",font=("楷体",15),bg="#00FF7F").grid(row=0,column=0)
        tkinter.OptionMenu(Quan,Amodel,*Option_number).grid(row=0,column=1)

        global Bmodel
        Bmodel=tkinter.StringVar()
        tkinter.Label(Quan,text="期中成绩权重:",font=("楷体",15),bg="#00FF7F").grid(row=1,column=0)
        tkinter.OptionMenu(Quan,Bmodel,*Option_number).grid(row=1,column=1)

        global Cmodel
        Cmodel=tkinter.StringVar()
        tkinter.Label(Quan,text="期末成绩权重:",font=("楷体",15),bg="#00FF7F").grid(row=2,column=0)
        tkinter.OptionMenu(Quan,Cmodel,*Option_number).grid(row=2,column=1)

        tkinter.Button(Quan,text="确定并关闭",font=("楷体",15),bg="#EEE8AA",
                command=lambda:Save_data(self.course_name).Quan_Save()).grid(row=3)


    def Open_Dir(self):          #打开对应的文件夹，显示学生信息excel文件列表
        if os.path.exists("E:\课程管理系统信息保存\《%s》课程"%(self.course_name))==False:
            os.makedirs(r"E:\课程管理系统信息保存\《%s》课程"%(self.course_name))
        else:
            pass
        dir_name="E:\课程管理系统信息保存\《%s》课程"%(self.course_name)
        os.system("explorer.exe %s"%(dir_name))


    def Grades_sorted(self):     #实现全体学生的排序（按照总成绩的从高到低的排序）
        try:
            dir_list=os.listdir("E:\课程管理系统信息保存\《%s》课程"%(self.course_name))
            dir_list.remove("课程教师信息汇总.xlsx")
            stu_list=dir_list
            xlApp=win32com.client.Dispatch("Excel.Application")
            for i in stu_list:
                xlBook=xlApp.Workbooks.Open("E:/课程管理系统信息保存/《%s》课程/%s"%(self.course_name,i))
                xlSht=xlBook.Worksheets("sheet1")
                if str(xlSht.Cells(1,7).Value)=="":
                    a=1/0      #强行执行报错语句，处理总成绩不存在的情况
                row=2
                while xlSht.Cells(row,1).Value not in [None,""]:
                    row=row+1
                copylist=[]             #原来表格每行数据的复制存放在copylist里面
                data_dict={}            #创建这个字典为了获得  学号=>总成绩 的键值对，用于排序
                for i in range(2,row):    
                    copy=xlSht.Range(xlSht.Cells(i,1),xlSht.Cells(i,7)).Value   
                    copylist.append(copy)
                    data_dict[str(xlSht.Cells(i,2).Value)]=str(xlSht.Cells(i,7).Value)
                tuple_list=sorted(data_dict.items(),key=lambda d:d[1],reverse=True)
                for i in range(2,row):
                    for j in copylist:
                        if int(float(tuple_list[i-2][0]))==int(j[0][1]):
                            xlSht.Cells(i,1).Value=j[0][0]
                            xlSht.Cells(i,2).Value=int(j[0][1])
                            xlSht.Cells(i,3).Value=j[0][2]
                            xlSht.Cells(i,4).Value=j[0][3]
                            xlSht.Cells(i,5).Value=j[0][4]
                            xlSht.Cells(i,6).Value=j[0][5]
                            xlSht.Cells(i,7).Value=j[0][6]
                            break
                xlBook.Close(SaveChanges=1)
            del xlApp                
                
        except Exception:
            global Error8
            Error8=tkinter.Tk()    
            tkinter.Button(Error8,text="总成绩不存在\n点击返回",font=("楷体",12),bg="#00FFFF",width=20,height=20,
                           command=lambda:Exit_aware(Error8)).grid()
            Error8.mainloop()
            
        
    def Analyse_single_for_choose(self):   #特定教师姓名的获取
        global choose
        choose=tkinter.Tk()
        choose.title("选择教师")
        #获取路径名
        dir_list=os.listdir("E:\课程管理系统信息保存\《%s》课程"%(self.course_name))
        dir_list.remove("课程教师信息汇总.xlsx")
        Cteacher=[]
        for i in dir_list:    #路径名中教师姓名的获取
            d1=i.replace("教师","",1)
            d2=d1.replace("的学生信息.xlsx","",1)
            Cteacher.append(d2)
        global Dmodel
        Dmodel=tkinter.StringVar()
        tkinter.Label(choose,text="请选择教师",font=("楷体",15),bg="red").grid(row=0,column=0)
        tkinter.OptionMenu(choose,Dmodel,*Cteacher).grid(row=0,column=1)
        tkinter.Button(choose,text="确定并执行",font=("楷体",15),bg="blue",
                command=lambda:Analyse_system(self.course_name).Analyse_single()).grid(row=1)
        choose.mainloop()
        
    def Analyse_single(self):     #特定教师学生总成绩的统计图式分析
        try:
            if Dmodel.get()=="":
                a=1/0
            BackError(choose)
            xlApp=win32com.client.Dispatch("Excel.Application")
            xlBook=xlApp.Workbooks.Open("E:/课程管理系统信息保存/《%s》课程/教师%s的学生信息.xlsx"%(self.course_name,Dmodel.get()))
            xlSht=xlBook.Worksheets("sheet1")
            row=2
            while xlSht.Cells(row,7).Value not in [None,""]:
                row=row+1
            whitestr=""        #将用来把所有总成绩以及加号组合在一起，最后通过eval函数去运算
            acount,bcount,ccount,dcount,ecount=0,0,0,0,0    #用于分段统计（100-90，90-80，80-70，70-60，60以下）
            dict_grade={}            #收获所有成绩数据，然后排序找出最高分与最低分
            for i in range(2,row):
                if i!=row-1:
                    whitestr+=str(xlSht.Cells(i,7).Value)+"+"
                    dict_grade[i]=str(xlSht.Cells(i,7).Value)
                    if (abs(float(xlSht.Cells(i,7).Value)-90.0)<1e-7) or float(xlSht.Cells(i,7).Value)>90.0:
                        acount+=1
                        continue
                    elif (abs(float(xlSht.Cells(i,7).Value)-80.0)<1e-7) or float(xlSht.Cells(i,7).Value)>80.0:
                        bcount+=1
                        continue
                    elif (abs(float(xlSht.Cells(i,7).Value)-70.0)<1e-7) or float(xlSht.Cells(i,7).Value)>70.0:
                        ccount+=1
                        continue
                    elif (abs(float(xlSht.Cells(i,7).Value)-60.0)<1e-7) or float(xlSht.Cells(i,7).Value)>60.0:
                        dcount+=1
                        continue
                    else:
                        ecount+=1
                else:
                    whitestr+=str(xlSht.Cells(i,7).Value)
                    dict_grade[i]=str(xlSht.Cells(i,7).Value)
                    if (abs(float(xlSht.Cells(i,7).Value)-90.0)<1e-7) or float(xlSht.Cells(i,7).Value)>90.0:
                        acount+=1
                    elif (abs(float(xlSht.Cells(i,7).Value)-80.0)<1e-7) or float(xlSht.Cells(i,7).Value)>80.0:
                        bcount+=1
                    elif (abs(float(xlSht.Cells(i,7).Value)-70.0)<1e-7) or float(xlSht.Cells(i,7).Value)>70.0:
                        ccount+=1
                    elif (abs(float(xlSht.Cells(i,7).Value)-60.0)<1e-7) or float(xlSht.Cells(i,7).Value)>60.0:
                        dcount+=1
                    else:
                        ecount+=1
            average_single=round(eval("("+whitestr+")"+"/"+str(row-2)),3)      #得到平均分
            qualified_percent_single=str(round((acount+bcount+ccount+dcount)/(row-2)*100,3))+"%"   #合格率
            excellent_percent_single=str(round(acount/(row-2)*100,3))+"%"       #优秀率
            grade_sort_list_single=sorted(dict_grade.items(),key=lambda d:d[1],reverse=True)   #所有成绩从高到低排序，获得一个元组列表
            best_grade_single=str(grade_sort_list_single[0][1])        #获得最高分
            worst_grade_single=str(grade_sort_list_single[row-3][1])   #获得最低分
            xlBook.Close(SaveChanges=0)
            del xlApp
            
            #************cpp的Label的布局，用于显示平均分、最高分、最低分、合格率、优秀率。
            tkinter.Label(cpp,text="教师%s学生成绩平均分为%s"%(Dmodel.get(),str(average_single)),font=("楷体",15),bg="#FFB6C1",
                          width=50,height=3).grid(row=0,column=0,rowspan=2)
            tkinter.Label(cpp,text="教师%s学生成绩最高分为%s"%(Dmodel.get(),str(best_grade_single)),font=("楷体",15),bg="#BA55D3",
                          width=50,height=3).grid(row=2,column=0,rowspan=2)
            tkinter.Label(cpp,text="教师%s学生成绩最低分为%s"%(Dmodel.get(),str(worst_grade_single)),font=("楷体",15),bg="#0000CD",
                          width=50,height=3).grid(row=4,column=0,rowspan=2)
            tkinter.Label(cpp,text="教师%s学生成绩合格率为%s"%(Dmodel.get(),qualified_percent_single),font=("楷体",15),bg="#00FFFF",
                          width=50,height=3).grid(row=6,column=0,rowspan=2)
            tkinter.Label(cpp,text="教师%s学生成绩优秀率为%s"%(Dmodel.get(),excellent_percent_single),font=("楷体",15),bg="#F0E68C",
                          width=50,height=3).grid(row=8,column=0,rowspan=2)
            tkinter.Label(cpp,text="考的不错的同学再接再厉，不要骄傲!\n考的不好的同学要加油哦，永不言弃!",font=("楷体",15),bg="#E0FFFF",
                          width=50,height=3).grid(row=10,column=0,rowspan=2)
            
            #************统计图开始编写(用于分数段人数的统计直观显示)*******
            a,b=plt.subplots()     #调用pyplot模块里的subplots()函数，并将返回值付予变量a和b
            array1=np.arange(1,5+1)   #生成(x轴)array数组一
            array2=[acount,bcount,ccount,dcount,ecount]  ##生成(y轴)array数组二(值为各分数段的人数)
            c=b.bar(array1,array2,0.5,color="#228B22",yerr=1)
            #  bar是柱的显示， 0.5设置了柱形的宽度，color设置了颜色，yerr为后缀值
            for i in c: #迭代序列对象
                height=i.get_height() #获取柱形高度值
                width=i.get_width()  #获取柱形宽度值
                x,y=i.get_x(),i.get_y()  #x,y轴的获取
                b.text(x+0.25,y+height,height,ha="center",va="bottom")
                #第一个参数是在x轴上的(柱形)上标识(柱形)的宽度,这里加0.25的意思是我们可以把标识数值往(柱形的右侧移一些位置,这样美观一点)
                #第二个参数指的是标识写在(柱形)的什么高度的位置
                #第三个参数指的是(柱形)的高度值
                #ha='center'意思是在(柱形)的顶部标识
                #va='bottom'意思是在(柱形)的底部标识
            b.legend((c[0],),(u'.C',))#两个参数皆为元组(注:这里有中文,所以前面加u转换为U码,但仅仅这样还不能显示中文,所以有一开始的指定默认字体语句)
            b.set_xticklabels([u'90~100分',u'80~90分','70~80分',u'60~70分',u'60分以下'], fontproperties=font)#设置x轴的自定义标签
            b.set_xticks([1,2,3,4,5])#X轴作标记,用于微调X轴的标签位置
            plt.title(u'教师%s的学生成绩的分段统计'%(Dmodel.get()), fontproperties=font)   #标题
            plt.xlabel(u'分数段', fontproperties=font) #X轴名称
            plt.ylabel(u'人数', fontproperties=font)  #Y轴名称
            plt.show()      #显示统计图

        except Exception:
            global Error9
            Error9=tkinter.Tk()    
            tkinter.Button(Error9,text="操作错误\n点击关闭",font=("楷体",12),bg="#00FFFF",width=20,height=20,
                           command=lambda:Exit_aware(Error9)).grid()
            Error9.mainloop() 
    def Analyse_all(self):       #全体学生总成绩的统计图式分析
        try:
            dir_list=os.listdir("E:\课程管理系统信息保存\《%s》课程"%(self.course_name))
            dir_list.remove("课程教师信息汇总.xlsx")
            stu_list=dir_list
            xlApp=win32com.client.Dispatch("Excel.Application")
            whitestr_set=""        #将用来把所有总成绩以及加号组合在一起（所有excel）
            acount_set,bcount_set,ccount_set,dcount_set,ecount_set=0,0,0,0,0    #用于分段统计（100-90，90-80，80-70，70-60，60以下）（所有excel）
            dict_grade_set={}            #收获所有成绩数据(所有excel)
            index=0          #字典所有项目的键名
            #****获取所有excel的信息
            for j in stu_list:
                xlBook=xlApp.Workbooks.Open("E:/课程管理系统信息保存/《%s》课程/%s"%(self.course_name,j))
                xlSht=xlBook.Worksheets("sheet1")
                row=2
                while xlSht.Cells(row,7).Value not in [None,""]:
                    row=row+1
                whitestr=""        #将用来把所有总成绩以及加号组合在一起（一个excel）
                acount,bcount,ccount,dcount,ecount=0,0,0,0,0    #用于分段统计（100-90，90-80，80-70，70-60，60以下）（一个excel）
                dict_grade={}            #收获所有成绩数据(一个excel)
                for i in range(2,row):
                    if i!=row-1:
                        whitestr+=str(xlSht.Cells(i,7).Value)+"+"
                        dict_grade[index+i-1]=str(xlSht.Cells(i,7).Value)
                        if (abs(float(xlSht.Cells(i,7).Value)-90.0)<1e-7) or float(xlSht.Cells(i,7).Value)>90.0:
                            acount+=1
                            continue
                        elif (abs(float(xlSht.Cells(i,7).Value)-80.0)<1e-7) or float(xlSht.Cells(i,7).Value)>80.0:
                            bcount+=1
                            continue
                        elif (abs(float(xlSht.Cells(i,7).Value)-70.0)<1e-7) or float(xlSht.Cells(i,7).Value)>70.0:
                            ccount+=1
                            continue
                        elif (abs(float(xlSht.Cells(i,7).Value)-60.0)<1e-7) or float(xlSht.Cells(i,7).Value)>60.0:
                            dcount+=1
                            continue
                        else:
                            ecount+=1
                    else:
                        whitestr+=str(xlSht.Cells(i,7).Value)
                        dict_grade[index+i-1]=str(xlSht.Cells(i,7).Value)
                        if (abs(float(xlSht.Cells(i,7).Value)-90.0)<1e-7) or float(xlSht.Cells(i,7).Value)>90.0:
                            acount+=1
                        elif (abs(float(xlSht.Cells(i,7).Value)-80.0)<1e-7) or float(xlSht.Cells(i,7).Value)>80.0:
                            bcount+=1
                        elif (abs(float(xlSht.Cells(i,7).Value)-70.0)<1e-7) or float(xlSht.Cells(i,7).Value)>70.0:
                            ccount+=1
                        elif (abs(float(xlSht.Cells(i,7).Value)-60.0)<1e-7) or float(xlSht.Cells(i,7).Value)>60.0:
                            dcount+=1
                        else:
                            ecount+=1
                xlBook.Close(SaveChanges=0)            
                whitestr_set=whitestr_set+whitestr+"+"
                index=index+row-2
                acount_set+=acount
                bcount_set+=bcount
                ccount_set+=ccount
                dcount_set+=dcount
                ecount_set+=ecount
                dict_grade_set.update(dict_grade.items())      #字典相加的一种方法
                
            whitestr_set+="0"    #处理所有数据综合在一起多了一个"+"的问题
            average_all=round(eval("("+whitestr_set+")"+"/"+str(index)),3)      #得到平均分
            qualified_percent_all=str(round((acount_set+bcount_set+ccount_set+dcount_set)/(index)*100,3))+"%"   #合格率
            excellent_percent_all=str(round(acount_set/(index)*100,3))+"%"    #优秀率
            grade_sort_list_all=sorted(dict_grade_set.items(),key=lambda d:d[1],reverse=True)   #所有成绩从高到低排序，获得一个元组列表
            best_grade_all=str(grade_sort_list_all[0][1])        #获得最高分
            worst_grade_all=str(grade_sort_list_all[index-1][1])   #获得最低分
            del xlApp
            
            #************cpp的Label的布局，用于显示平均分、最高分、最低分、合格率、优秀率。
            tkinter.Label(cpp,text="全体学生成绩平均分为%s"%(str(average_all)),font=("楷体",15),bg="#FFB6C1",
                          width=50,height=3).grid(row=0,column=0,rowspan=2)
            tkinter.Label(cpp,text="全体学生成绩最高分为%s"%(str(best_grade_all)),font=("楷体",15),bg="#BA55D3",
                          width=50,height=3).grid(row=2,column=0,rowspan=2)
            tkinter.Label(cpp,text="全体学生成绩最低分为%s"%(str(worst_grade_all)),font=("楷体",15),bg="#0000CD",
                          width=50,height=3).grid(row=4,column=0,rowspan=2)
            tkinter.Label(cpp,text="全体学生成绩合格率为%s"%(qualified_percent_all),font=("楷体",15),bg="#00FFFF",
                          width=50,height=3).grid(row=6,column=0,rowspan=2)
            tkinter.Label(cpp,text="全体学生成绩优秀率为%s"%(excellent_percent_all),font=("楷体",15),bg="#F0E68C",
                          width=50,height=3).grid(row=8,column=0,rowspan=2)
            tkinter.Label(cpp,text="考的不错的同学再接再厉，不要骄傲!\n考的不好的同学要加油哦，永不言弃!",font=("楷体",15),bg="#E0FFFF",
                          width=50,height=3).grid(row=10,column=0,rowspan=2)
            
            #************统计图开始编写(用于分数段人数的统计直观显示)*******
            a,b=plt.subplots()     #调用pyplot模块里的subplots()函数，并将返回值付予变量a和b
            array1=np.arange(1,5+1)   #生成(x轴)array数组一
            array2=[acount_set,bcount_set,ccount_set,dcount_set,ecount_set]  ##生成(y轴)array数组二(值为各分数段的人数)
            c=b.bar(array1,array2,0.5,color="#228B22",yerr=1)
            #  bar是柱的显示， 0.5设置了柱形的宽度，color设置了颜色，yerr为后缀值
            for i in c: #迭代序列对象
                height=i.get_height() #获取柱形高度值
                width=i.get_width()  #获取柱形宽度值
                x,y=i.get_x(),i.get_y()  #x,y轴的获取
                b.text(x+0.25,y+height,height,ha="center",va="bottom")
                #第一个参数是在x轴上的(柱形)上标识(柱形)的宽度,这里加0.25的意思是我们可以把标识数值往(柱形的右侧移一些位置,这样美观一点)
                #第二个参数指的是标识写在(柱形)的什么高度的位置
                #第三个参数指的是(柱形)的高度值
                #ha='center'意思是在(柱形)的顶部标识
                #va='bottom'意思是在(柱形)的底部标识
            b.legend((c[0],),(u'.C',))#两个参数皆为元组(注:这里有中文,所以前面加u转换为U码,但仅仅这样还不能显示中文,所以有一开始的指定默认字体语句)
            b.set_xticklabels([u'90~100分',u'80~90分','70~80分',u'60~70分',u'60分以下'], fontproperties=font)#设置x轴的自定义标签
            b.set_xticks([1,2,3,4,5])#X轴作标记,用于微调X轴的标签位置
            plt.title(u'本课程所有学生成绩的分段统计图',fontproperties=font)   #标题
            plt.xlabel(u'分数段', fontproperties=font) #X轴名称
            plt.ylabel(u'人数', fontproperties=font)  #Y轴名称
            plt.show()      #显示统计图

        except Exception:
            global Error10
            Error10=tkinter.Tk()
            tkinter.Button(Error10,text="操作错误\n点击关闭",font=("楷体",12),bg="#00FFFF",width=20,height=20,
                           command=lambda:Exit_aware(Error10)).grid()
            Error10.mainloop()
            
            
'''#****************************************************************************************************************************
#****************************************************************************************************************************
#^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
#^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^'''


#采用按钮grid布局方式，每一个按钮对应一门课程系统。
def main():
    global root1
    root1=tkinter.Tk()
    root1.title("课程管理系统")
    root1.geometry("524x490")
    tkinter.Label(root1,text="欢迎使用课程管理系统!",font=("楷体",12),bg="#FFC0CB",height=3,width=65).grid(row=0,column=0,columnspan=4,sticky="w")
    tkinter.Button(root1,text="点击查看本课程管理系统的使用说明",font=("楷体",12),relief="raised",borderwidth=5,bg="#E0FFFF",
                   height=3,width=64,command=Use_description).grid(row=1,column=0,columnspan=4,sticky="w")
    tkinter.Button(root1,text="工科数学分析",font=("楷体",12),bg="#87CEEB",height=4,width=15,
                   command=lambda:Course_Widget("工科数学分析").All_command_one()).grid(row=2,column=0,sticky="w")

    tkinter.Button(root1,text="工科高等代数",font=("楷体",12),bg="#87CEEB",height=4,width=15,
                   command=lambda:Course_Widget("工科高等代数").All_command_one()).grid(row=2,column=1,sticky="w")
    tkinter.Button(root1,text="离散数学",font=("楷体",12),bg="#87CEEB",height=4,width=15,
                   command=lambda:Course_Widget("离散数学").All_command_one()).grid(row=2,column=2,sticky="w")
    tkinter.Button(root1,text="博雅课堂",font=("楷体",12),bg="#87CEEB",height=4,width=15,
                   command=lambda:Course_Widget("博雅课堂").All_command_one()).grid(row=2,column=3,sticky="w")

    tkinter.Button(root1,text="计算机基础操作",font=("楷体",12),bg="#7FFFD4",height=4,width=15,
                   command=lambda:Course_Widget("计算机基础操作").All_command_one()).grid(row=3,column=0,sticky="w")
    tkinter.Button(root1,text="走进计算机\n科学系列讲座",font=("楷体",12),bg="#7FFFD4",height=4,width=15,
                   command=lambda:Course_Widget("走进计算机科学系列讲座").All_command_one()).grid(row=3,column=1,sticky="w")
    tkinter.Button(root1,text="计算机导论\n与伦理学",font=("楷体",12),bg="#7FFFD4",height=4,width=15,
                   command=lambda:Course_Widget("计算机导论与伦理学").All_command_one()).grid(row=3,column=2,sticky="w")
    tkinter.Button(root1,text="航空航天概论",font=("楷体",12),bg="#7FFFD4",height=4,width=15,
                   command=lambda:Course_Widget("航空航天概论").All_command_one()).grid(row=3,column=3,sticky="w")

    tkinter.Button(root1,text="大学生体育",font=("楷体",12),bg="#FFFFF0",height=4,width=15,
                   command=lambda:Course_Widget("大学生体育").All_command_one()).grid(row=4,column=0,sticky="w")
    tkinter.Button(root1,text="思想道德修养\n与法律基础",font=("楷体",12),bg="#FFFFF0",height=4,width=15,
                   command=lambda:Course_Widget("思想道德修养与法律基础").All_command_one()).grid(row=4,column=1,sticky="w")
    tkinter.Button(root1,text="英语听说写B",font=("楷体",12),bg="#FFFFF0",height=4,width=15,
                   command=lambda:Course_Widget("英语听说写B").All_command_one()).grid(row=4,column=2,sticky="w")
    tkinter.Button(root1,text="学业英语阅读\n与写作B",font=("楷体",12),bg="#FFFFF0",height=4,width=15,
                   command=lambda:Course_Widget("学业英语阅读与写作B").All_command_one()).grid(row=4,column=3,sticky="w")

    tkinter.Button(root1,text="高级英语听说写A",font=("楷体",12),bg="#F08080",height=4,width=15,
                   command=lambda:Course_Widget("高级英语听说写A").All_command_one()).grid(row=5,column=0,sticky="w")
    tkinter.Button(root1,text="批判阅读\n与写作A",font=("楷体",12),bg="#F08080",height=4,width=15,
                   command=lambda:Course_Widget("批判阅读与写作A").All_command_one()).grid(row=5,column=1,sticky="w")
    tkinter.Button(root1,text="科技实践课堂\n线上课程",font=("楷体",12),bg="#F08080",height=4,width=15,
                   command=lambda:Course_Widget("科技实践课堂线上课程").All_command_one()).grid(row=5,column=2,sticky="w")
    tkinter.Button(root1,text="科技实践课堂\n线下课程",font=("楷体",12),bg="#F08080",height=4,width=15,
                   command=lambda:Course_Widget("科技实践课堂线下课程").All_command_one()).grid(row=5,column=3,sticky="w")
    tkinter.Label(root1,text="温馨提示：本程序是面向计算机学院2015级大一新生的\n课程管理系统，暂不服务其他系同学的课程管理，请见谅!",font=("楷体",12),
                  bg="#FFFAFA",fg="#B8860B",height=4,width=65).grid(row=6,column=0,columnspan=4,sticky="w")
    root1.mainloop()
     
####################################
#程序运行初始点               
if __name__=="__main__":
    global aware1
    aware1=tkinter.Tk()
    aware1.title("<<--温馨提醒-->>")
    tkinter.Label(aware1,text="您好，欢迎使用课程管理系统!\n使用系统前请将名为《课程管理系统\n使用说明》的txt文件直接放入\nE盘即可,不要放入文件夹内,谢谢!",
                  fg="green",font=("楷体",12),bg="#E0FFFF").grid(row=0)
    tkinter.Button(aware1,text="点击退出",font=("楷体",12),bg="#FFFFF0",command=lambda:Exit_aware(aware1)).grid(row=1)
    main()
    
'''
this program
'''
    
    
    










