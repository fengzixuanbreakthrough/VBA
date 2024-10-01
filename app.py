# import win32com.client
# print('转换开始')
# #####调用vba程序。需要安装win32com库

# xls = win32com.client.Dispatch("Excel.Application")
# xls.Workbooks.Open(r"C:\Users\fengzixuan\Desktop\工作表 在 FDR-FR-SRD文档及追溯关系转换工具（VBA）使用说明.xlsm")  ##存储vba代码的文件
# try:
#     xls.Application.Run('FRDTranslate.FRD_Translate')  ##开始调用vba宏
# except Exception as e:
#     print(e)
# xls.Application.Quit()

# print('转换完成！')
# import xlwings as xw

# wb = xw.Workbook.active()

# your_macro = wb.macro('YourMacro')

# your_macro(1, 2)

import os
import time
from typing import final
import pythoncom
import marco
from flask import Flask, jsonify, render_template, request, redirect, url_for, send_file, send_from_directory
errorlistFDR = []
pathFDR = ''
pathFDR2 = ''
FDRFirstName = ''
pathFDR2 = ''
file_path_abs2 = ''
app = Flask(__name__)


@app.route("/")  # 使用Flask的 app.route 装饰器将URL路由映射到该函数：
def home():

    return render_template('vba.html')


@app.route("/vabfirst", methods=['GET', 'POST'])
def vabfirst():
    global FDRFirstName
    global pathFDR
    errorlistFDR = []
    pathFDR = ''
    FDRFirstName = ''
    pathsource = ''
    if request.method == 'POST':

        file = request.files['firstFDR']
        print("用户上传的文件")
        print(file)
        print(file.filename)
        print(type(file.filename))

        if file.name == '':
            error2 = '文件未上传，请上传'
            return render_template('show.html', error2=error2)

        else:
            print('满足文件类型')
            filename1 = file.filename
            print("文件名：")
            print(filename1)
            pathFDR = "static/FDR/"
            if os.path.exists(pathFDR):
                print('exist')
            else:
                os.mkdir(pathFDR)
            # 创建项目名称文件夹下版本编号下用户下时间的文件夹
            time22 = time.localtime()
            t = time.strftime('%Y%m%d%H%M%S', time22)

            pathFDR = "static/FDR/" + t+'/'
            if os.path.exists(pathFDR):
                print('exist')
            else:
                os.mkdir(pathFDR)

            ##############################################创建文件夹结束#####################################

                ###
            file_path = pathFDR + filename1
            print("此处的path：（1）")
            print(file_path)

            file.save(file_path)
            #User.query.filter_by(username=username).update({'filepath': file_path})
            pythoncom.CoInitialize()
            # dictionary,allecu,errorlist=dictToDBC.main2(fileexcelpath=file_path,OneStepTotalNumber=version)
            # reresult=doct6.zhuhanshu(file_path,path)#path为文件夹的地址 filepath为文件夹下文件的详细地址
            print("返回结aaaaaaa果：")
            file_path_abs = os.path.abspath(file_path)
            print(file_path_abs)
            FDRFirstName = request.form.get('FDRFirstName')
            if FDRFirstName == None:
                errorlistFDR.append["文件名字未输入"]
            else:
                print(FDRFirstName)
                pathsource = r"D:\3-PythonProjects\VBA\tool\FDR转换.xlsm"
                try:
                    marco.macroFRD_Translate(
                        pathsource, file_path_abs, FDRFirstName)
                except Exception as e:
                    errorlistFDR = ["生成失败"]
            pythoncom.CoUninitialize()
            # marco.macroFRD_Translate(pathsource,file_path_abs,"fengzixuan")
            return jsonify(errorlistFDR)


@app.route('/downloadfirst', methods=['GET', 'POST'])
def downloadfirst():
    global FDRFirstName
    global pathFDR

    firstFinalPath = os.path.abspath(pathFDR+FDRFirstName+'.xlsx')
    print(pathFDR)
    print(firstFinalPath)
    return send_file(firstFinalPath.encode('utf-8').decode('utf-8'), as_attachment=True)


@app.route("/vabsecond", methods=['GET', 'POST'])
def vabsecond():
    global pathFDR2
    global file_path_abs2
    pathFDR2 = ''
    file_path_abs2 = ''
    pathsource = ''
    errorlistFDR2=[]
    if request.method == 'POST':

        file = request.files['secondFDR']
        print("用户上传的文件")
        print(file)
        print(file.filename)
        print(type(file.filename))

        if file.name == '':
            error2 = '文件未上传，请上传'
            return render_template('show.html', error2=error2)

        else:
            print('满足文件类型')
            filename1 = file.filename
            print("文件名：")
            print(filename1)
            pathFDR2 = "static/FDR2"
            if os.path.exists(pathFDR2):
                print('exist')
            else:
                os.mkdir(pathFDR2)
            # 创建项目名称文件夹下版本编号下用户下时间的文件夹
            time22 = time.localtime()
            t = time.strftime('%Y%m%d%H%M%S', time22)

            pathFDR2 = "static/FDR/" + t+'/'
            if os.path.exists(pathFDR2):
                print('exist')
            else:
                os.mkdir(pathFDR2)

            ##############################################创建文件夹结束#####################################

                ###
            file_path2 = pathFDR2 + filename1
            print("此处的path：（1）")
            print(file_path2)

            file.save(file_path2)
            #User.query.filter_by(username=username).update({'filepath': file_path})
            pythoncom.CoInitialize()
            # dictionary,allecu,errorlist=dictToDBC.main2(fileexcelpath=file_path,OneStepTotalNumber=version)
            # reresult=doct6.zhuhanshu(file_path,path)#path为文件夹的地址 filepath为文件夹下文件的详细地址
            print("返回结aaaaaaa果：")
            file_path_abs2 = os.path.abspath(file_path2)

            pathsource = r"D:\3-PythonProjects\VBA\tool\FDR追溯关系_V0.2.xlsm"
            try:
                marco.macroFindFDRID(pathsource, file_path_abs2)
            except Exception as e:
                errorlistFDR2 = ["生成ID失败"]
            pythoncom.CoUninitialize()
            # marco.macroFRD_Translate(pathsource,file_path_abs,"fengzixuan")
            return jsonify(errorlistFDR2)


@app.route("/relativeLaws", methods=['GET', 'POST'])
def relativeLaws():
    global pathFDR2
    global file_path_abs2
  
    errorlistFDR2=[]
    pathsource = ''
    pathsource = r"D:\3-PythonProjects\VBA\tool\FDR追溯关系_V0.21.xlsm"
    if request.method == 'POST':
        # try:
        pythoncom.CoInitialize()
            # dictionary,allecu,errorlist=dictToDBC.main2(fileexcelpath=file_path,OneStepTotalNumber=version)
            # reresult=doct6.zhuhanshu(file_path,path)#path为文件夹的地址 filepath为文件夹下文件的详细地址
        print("返回结aaaaaaa果：")
        marco.macroFindlaw(pathsource, file_path_abs2)
        # except Exception as e:
        #     errorlistFDR2 = ["关联法律失败"]
        pythoncom.CoUninitialize()
    return jsonify(errorlistFDR2)


@app.route("/relativeUC", methods=['GET', 'POST'])
def relativeUC():
    global pathFDR2
    global file_path_abs2
    errorlistFDR2=[]
    pathsource = ''
    pathsource = r"D:\3-PythonProjects\VBA\tool\FDR追溯关系_V0.22.xlsm"
    if request.method == 'POST':
        try:
            pythoncom.CoInitialize()
            # dictionary,allecu,errorlist=dictToDBC.main2(fileexcelpath=file_path,OneStepTotalNumber=version)
            # reresult=doct6.zhuhanshu(file_path,path)#path为文件夹的地址 filepath为文件夹下文件的详细地址
            print("返回结aaaaaaa果：")
            marco.macrofindUC(pathsource, file_path_abs2)
        except Exception as e:
            errorlistFDR2 = ["关联UC失败"]
    return jsonify(errorlistFDR2)

@app.route('/downloadsecond', methods=['GET', 'POST'])
def downloadsecond():
    global pathFDR2
    global file_path_abs2

    
    return send_file(file_path_abs2.encode('utf-8').decode('utf-8'), as_attachment=True)
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
# my_sum(1,2)
# print(my_sum(1,2))
# import xlwings as xw

# app = xw.App(visible=True,add_book=False)

# # 设置测试文件的路径

# wb = app.books.open(r'C:/Users/zhoux/Desktop/test.xlsm')

# # 调用VBA脚本

# my_sum = wb.macro('MySum')

# my_sum(1, 2)
# print(my_sum(1, 2))
# import win32com.client
# import os
# xl=win32com.client.Dispatch("Excel.Application")
# xl.Visible = True
# Path = os.getcwd()+"/test.xlsm"
# # wb为了单独处理excel，xl是处理窗口
# wb=xl.Workbooks.Open(Filename=Path)
# wb.Save()
# wb.Close()
# xl.Quit()
