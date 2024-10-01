import pdfkit

'''将html文件生成pdf文件'''
def html_to_pdf(html, to_file):
    # 将wkhtmltopdf.exe程序绝对路径传入config对象
    path_wkthmltopdf = r'D:\wkhtmltopdf\bin\wkhtmltopdf.exe'
    config = pdfkit.configuration(wkhtmltopdf=path_wkthmltopdf)
    # 生成pdf文件，to_file为文件路径
    pdfkit.from_file(html, to_file, configuration=config)
    print('html转pdf文件完成')

#html_to_pdf(r'C:\Users\libo5\Desktop\work\GWNEWROUTE\CANoeConfig\TestModules\NewFile1new_report.html','out_2.pdf')

test= 'abc.html'
result=test.split('.')
print(result)