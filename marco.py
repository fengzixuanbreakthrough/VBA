

import xlwings as xw
import pythoncom
def macroFRD_Translate(pathsource,pathdest,name):

    app = xw.App(visible=False,add_book=False)
        # path="D:\2-以太网研讨\FUN_BD_010_转向信号灯控制\第二代商用车_EE功能需求描述_转向信号灯控制_V0.6.doc"
 # 设置测试⽂件的路径
    wb = app.books.open(pathsource)
            # 调⽤VBA脚本
    your_macro = wb.macro('FRDTranslate.FRD_Translate')
    your_macro(pathdest,name)
    print('转换完成！')
    wb.close()
    app.quit()
def macroFindFDRID(pathsource,pathdest):

    app = xw.App(visible=False,add_book=False)
        # path="D:\2-以太网研讨\FUN_BD_010_转向信号灯控制\第二代商用车_EE功能需求描述_转向信号灯控制_V0.6.doc"
 # 设置测试⽂件的路径
    wb = app.books.open(pathsource)
            # 调⽤VBA脚本
    your_macro = wb.macro('ascend1.FindFDRID')
    your_macro(pathdest)
    print('转换完成！')
    wb.close()
    app.quit()
def macroFindlaw(pathsource,pathdest):

    app = xw.App(visible=False,add_book=False)
        # path="D:\2-以太网研讨\FUN_BD_010_转向信号灯控制\第二代商用车_EE功能需求描述_转向信号灯控制_V0.6.doc"
 # 设置测试⽂件的路径
    pythoncom.CoInitialize()
    wb = app.books.open(pathsource)
            # 调⽤VBA脚本
    your_macro = wb.macro('ascend1.Findlaw')
    your_macro(pathdest)
    print('转换完成！')
    wb.close()
    app.quit()
def macrofindUC(pathsource,pathdest):

    app = xw.App(visible=False,add_book=False)
        # path="D:\2-以太网研讨\FUN_BD_010_转向信号灯控制\第二代商用车_EE功能需求描述_转向信号灯控制_V0.6.doc"
 # 设置测试⽂件的路径
    wb = app.books.open(pathsource)
            # 调⽤VBA脚本
    your_macro = wb.macro('ascend1.findUC')
    your_macro(pathdest)
    print('转换完成！')
    wb.close()
    app.quit()
if __name__ == '__main__':
    pathdest=r"D:\3-PythonProjects\VBA\static\FDR\20220902105326\第二代商用车_EE功能需求描述_转向信号灯控制_V0.6.doc"
    pathsource=r"D:\3-PythonProjects\VBA\tool\FDR转换.xlsm"
    macroFRD_Translate(pathsource,pathdest,"fengzixuan")
