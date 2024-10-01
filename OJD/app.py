from flask import Flask,render_template,request
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import create_engine
from sqlalchemy import and_, or_
from flask import Flask,flash,send_file,make_response
import datetime
import time
from flask import abort, redirect, url_for, render_template
import doct6
import os
from sqlalchemy import create_engine, MetaData, Table, Column, Integer, String
import flask_excel
import pymysql ,re,json
import zipfile
from io import BytesIO
import pymysql
import zipfile
import shutil
import os
import pdfkit
app = Flask(__name__)
app.secret_key='123'
#————————————————————————配置数据库————————————————————————————
HOSTNAME='127.0.0.1'
PORT='3306'
DATABASE='flask_sql'
USERNAME='root'
PASSWORD='123456'
DB_URI='mysql+pymysql://{}:{}@{}:{}/{}'.format(USERNAME,PASSWORD,HOSTNAME,PORT,DATABASE)
app.config['SQLALCHEMY_DATABASE_URI']=DB_URI
app.config['SQLALCHEMY_TRACK_MODIFICATIONS']=False#跟踪数据库修改
#设置下方这行code后，在每次请求结束后会自动提交数据库中的变动
app.config['SQLALCHEMY_COMMIT_ON_TEARDOWN'] = True
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = True
app.config['UPLOAD_FOLDER']='upload'
db=SQLAlchemy(app)#实例化数据库对象，它提供访问Flask-SQLAlchemy的所有功能
engine=db.get_engine()#创建数据库引擎

#db2 = pymysql.connect(host="localhost", user="root", passwd=PASSWORD, database=DATABASE, port=3306)
#———————————————————————————数据库配置结束————————————————————————————
#定义全局变量
username=''
regtime=''
image=''
file_path=''


class User(db.Model):
    __name__='users'
    id = db.Column(db.Integer, primary_key=True,autoincrement=True)
    username = db.Column(db.String(80), unique=True)
    password = db.Column(db.String(80))
    regtime = db.Column(db.DateTime, default=datetime.datetime.now)  # 注册时间
    image = db.Column(db.String(80))
    progamname = db.Column(db.String(80))
    version = db.Column(db.String(80))
    time = db.Column(db.String(80))
    filepath = db.Column(db.String(500))
    htmlpath=db.Column(db.String(500))

db.create_all()

def create_table(table_name):
    meta = MetaData()
    table_name = Table(table_name, meta,Column('id', Integer, primary_key=True,autoincrement=True),Column('progamname', String(80)),Column('version', String(80)),Column('filetime', String(80)),Column('filepath', String(255)),Column('htmlpath', String(255)))
    meta.create_all(engine)
    return table_name


# 登录检验（用户名、密码验证）
def valid_login(username, password):
    user = User.query.filter(and_(User.username == username, User.password == password)).first()
    if user:
        print('用户名与用户密码核对正确')
        return True
    else:
        print('用户名与用户密码核对不匹配')
        return False
# 注册检验（用户名验证）
def valid_regist(username):
    user = User.query.filter(User.username == username).first()
    if user:
        return False
    else:
        return True

def changedate(date):
    year=date[0:4]+'年'
    month=date[4:6]+'月'
    ri = date[6:8] + '日'
    hour=date[8:10]+'时'
    minute=date[10:12]+'分'
    second=date[12:14]+'秒'
    date=year+month+ri+hour+minute+second
    return date

def allowed_file(filename):
    ALLOWED_EXTENSIONS = {'png', 'jpg'}
    return '.' in filename and filename.rsplit('.',1)[1] in ALLOWED_EXTENSIONS

'''
@app.route('/article', methods=['GET', 'POST'])
def article_view():
    userperson=User(username='1',password='2')
    db.session.add(userperson)
    db.session.commit()
    return 'success'
'''
def allowed_file2(filename):
    ALLOWED_EXTENSIONS = {'xls'}
    return '.' in filename and filename.rsplit('.',1)[1] in ALLOWED_EXTENSIONS


'''将html文件生成pdf文件'''
def html_to_pdf(html, to_file):
    # 将wkhtmltopdf.exe程序绝对路径传入config对象
    path_wkthmltopdf = r'D:\wkhtmltopdf\bin\wkhtmltopdf.exe'
    config = pdfkit.configuration(wkhtmltopdf=path_wkthmltopdf)
    # 生成pdf文件，to_file为文件路径
    pdfkit.from_file(html, to_file, configuration=config)
    print('html转pdf文件完成')

@app.route('/index', methods=['GET', 'POST'])
def index():
    global username
    if request.method == 'GET':
        yonghu = User.query.filter(User.username == username).one()
        regtime=yonghu.regtime
        imagepath=yonghu.image
        return render_template('index.html',username=username,regtime=regtime,imagepath=imagepath)
    if request.method == 'POST':
        return 'bbb'



@app.route('/gateway', methods=['GET', 'POST'])
def gateway():
    global username
    global file_path

    if request.method == 'GET':
        return render_template("gateway.html",username=username)
    if request.method == 'POST':


        progamname=request.form.get('progamname')
        print("用户输入的项目名")
        print(progamname)
        User.query.filter_by(username=username).update({'progamname': progamname})

        version = request.form.get('version')
        print("用户输入的版本号")
        print(version)
        User.query.filter_by(username=username).update({'version': version})

        discribe=request.form.get('discribe')
        print("用户输入的测试描述")
        print(discribe)
        print(type(discribe))

        file = request.files['gatewaylist']
        print("用户上传的文件")
        print(file)
        print(file.filename)
        print(type(file.filename))



        if file.name=='':
            error2 = '文件未上传，请上传'
            return render_template('gateway.html', error2=error2)
        elif progamname=='' or version=='' or discribe=='' or file.name=='':
            error1='有选项未填写，请填写完全'
            return render_template('gateway.html',error1=error1)
        elif file and allowed_file2(file.filename):
            print('满足文件类型')
            filename = file.filename
            print("文件名：")
            print(filename)

            ##############################################下述为创建文件夹#####################################
            #创建项目名称、版本编号对应的文件夹
            path = "static/xls/" + str(progamname)+'/'

            if os.path.exists(path):
                print('exist')
            else:
                os.mkdir(path)
            # 创建项目名称文件夹下版本编号对应的文件夹
            path = "static/xls/" + str(progamname)+'/'+str(version)+'/'

            if os.path.exists(path):
                print('exist')
            else:
                os.mkdir(path)
            # 创建项目名称文件夹下版本编号下用户的文件夹
            path = "static/xls/" + str(progamname) +'/'+str(version)+'/' + str(username)+'/'
            if os.path.exists(path):
                print('exist')
            else:
                os.mkdir(path)
            # 创建项目名称文件夹下版本编号下用户下时间的文件夹
            t = time.strftime('%Y%m%d%H%M%S')
            User.query.filter_by(username=username).update({'time': t})
            path = "static/xls/" + str(progamname) + '/' + str(version) + '/' + str(username) + '/'+t+'/'
            if os.path.exists(path):
                print('exist')
            else:
                os.mkdir(path)

            ##############################################创建文件夹结束#####################################

            ###
            file_path =path+filename
            print("此处的path：（1）")
            print(file_path)

            file.save(file_path)
            path_numcin = 'D:/OJD/' + path+'number.cin'
            path_nummsg = 'D:/OJD/' + path + 'message.cin'
            path_dictxt = 'D:/OJD/' + path + '测试描述.txt'

            User.query.filter_by(username=username).update({'filepath': file_path})
            db.session.commit()
            print('来此1')
            reresult=doct6.zhuhanshu(file_path,path)#path为文件夹的地址 filepath为文件夹下文件的详细地址
            print("返回结aaaaaaa果：")
            print(reresult)

            #####测试文件写入
            path_txt = 'D:/OJD/' + path
            print("测试描述写入地址")
            print(path_txt)
            filename_discribe= path+'测试描述.txt'
            with open(filename_discribe, 'w') as file_object:
                file_object.write("  本测试服务的项目：{}; \n".format(progamname))
                file_object.write("  本测试服务的测试人员：{}; \n".format(username))
                file_object.write("  本测试服务网关版本：{}; \n".format(version))
                file_object.write("  本次测试目的：{}; \n".format(discribe))
                file_object.write("  本次测试日期：{}; \n".format(t))
                shutil.copyfile(filename_discribe, r'D:\CANoeConfig\测试描述.txt')

                file_object.close()
            ##########
            if reresult == '已成功执行完毕':  # 该文件还未生成过那两个.cin文件
                filename_orign = r'D:\OJD\static\CANoeConfig.zip'
                shutil.copyfile(filename_orign, r'D:\test\CANoeConfig.zip')  # 目标文件无需存在


                memory_file = zipfile.ZipFile( r'D:\test\CANoeConfig.zip', 'a')


                addfilename1_orign ='D:/OJD/'+path+'number.cin'
                print('此时的addfilename1——orign')
                print(addfilename1_orign)
                shutil.copyfile(addfilename1_orign, r'D:\CANoeConfig\TestModules\number.cin')  # 目标文件无需存在

                addfilename2_orign ='D:/OJD/'+path+'message.cin'
                print('此时的addfilename2——orign')
                print(addfilename2_orign)
                shutil.copyfile(addfilename2_orign, r'D:\CANoeConfig\TestModules\message.cin')  # 目标文件无需存在
                print('到此123')
                addfilename3_orign = 'D:/OJD/'+path+'测试描述.txt'
                shutil.copyfile(addfilename3_orign, r'D:\CANoeConfig\测试描述.txt')  # 目标文件无需存在



                memory_file.write(r'D:\CANoeConfig\TestModules\number.cin')
                memory_file.write(r'D:\CANoeConfig\TestModules\message.cin')
                memory_file.write(r'D:\CANoeConfig\测试描述.txt')

                os.remove(r'D:\CANoeConfig\TestModules\number.cin')
                os.remove(r'D:\CANoeConfig\TestModules\message.cin')
                os.remove(r'D:\CANoeConfig\测试描述.txt')

                memory_file.close()


                success=1



                yonghu = User.query.filter(User.username == username).one()

                tableName=create_table(username)
                print('tablename类型')
                print(type(tableName))
                # 创建 insert 对象
                ins = tableName.insert()
                # 绑定要插入的数据
                ins = ins.values(progamname=yonghu.progamname,version=yonghu.version,filetime=yonghu.time,filepath=yonghu.filepath,htmlpath=yonghu.filepath)
                # 连接引擎
                conn = engine.connect()
                # 执行语句
                result = conn.execute(ins)







                return render_template('gateway.html',success=success,username=username,success_upload=1)
            else:
                return render_template('gateway2.html', text2=reresult)

        else:
            return 'error'

@app.route('/help', methods=['GET', 'POST'])
def help():
    if request.method == 'GET':
        return render_template("help.html")
    if request.method == 'POST':
        return 'a'

@app.route('/', methods=['GET', 'POST'])
def login():
    global username
    if request.method == 'GET':
        print("网页获取登陆界面")
        return render_template('login.html')
    if request.method == 'POST':

        username = request.form.get("email")
        password = request.form.get("password")
        print("在网页上输入的用户名：")
        print(username)
        print("在网页上输入的密码：")
        print(password)

        if valid_login(username, password):
            yonghu = User.query.filter(User.username == username).one()
            #regtime=yonghu.regtime
            #imagepath=yonghu.image


            return redirect(url_for('index'))
            #return render_template('index.html',username=username,regtime=regtime,imagepath=imagepath)
        elif User.query.filter(User.username == username).first():
            error = '密码错误'
            return render_template('login.html', error=error)

        else:
            error2 = '该用户名不存在'
            return render_template('login.html', error2=error2)


@app.route('/changephoto', methods=['get','post'])
def changephoto():
    global username
    yonghu = User.query.filter(User.username == username).one()
    regtime = yonghu.regtime
    imagepath = yonghu.image
    if request.method == 'GET':
        return render_template('photo.html',username=username,regtime=regtime,imagepath=imagepath)
    if request.method == 'POST':
        print('上传图片文件')
        file = request.files['photo']
        if file and allowed_file(file.filename):
            print('满足文件类型')
            filename = file.filename
            print("文件名：")
            print(filename)

            t = time.strftime('%Y%m%d%H%M%S')
            path =  "static/photo/"
            file_path = path +t+ filename
            file.save(file_path)
            imagepath = file_path

            User.query.filter_by(username=username).update({'image': imagepath})
            db.session.commit()
            return render_template('index.html',username=username,imagepath=imagepath,regtime=regtime)
        else:
            tip='未上传图片，请检查！'
            return render_template('photo.html',tip=tip,username=username,imagepath=imagepath,regtime=regtime)

@app.route('/changepwd', methods=['get','post'])
def changepwd():
    global username
    yonghu = User.query.filter(User.username == username).one()
    regtime = yonghu.regtime
    imagepath = yonghu.image
    if request.method == 'GET':
        return render_template('changepwd.html',username=username,imagepath=imagepath,regtime=regtime)

    if request.method == 'POST':
        oldpwd = request.form.get('oldpwd')
        password1 = request.form.get('upwd1')
        password2 = request.form.get('upwd2')
        print("老密码")
        print(oldpwd)
        print(password1)
        print(password2)
        user = User.query.filter(User.username == username).first()
        password=user.password
        if password1=='' or password2=='' or oldpwd=='':
            nulltip='有数据项未填写，请检查！'
            return render_template('changepwd.html', nulltip=nulltip,username=username,imagepath=imagepath,regtime=regtime)
        elif password1==password2:
            if oldpwd==password:
                #更新数据库
                User.query.filter_by(username=username).update({'password': password1})
                db.session.commit()
                success='修改完成'
                return render_template('login.html', success=success)
            else:
                error2='原始密码输入错误，请检查！'
                return render_template('changepwd.html', error2=error2,username=username,imagepath=imagepath,regtime=regtime)

        else:
            error1='两次输入新密码不一致，请检查！'
            return render_template('changepwd.html',error1=error1,username=username,imagepath=imagepath,regtime=regtime)


@app.route('/upload', methods=['GET', 'POST'])
def upload():
    print("进入upload路由")

    return send_file(r'D:\test\CANoeConfig.zip', attachment_filename='pack.zip', as_attachment=True)
@app.route('/save', methods=['get','post'])
def save():
    global username
    yonghu = User.query.filter(User.username == username).one()
    if request.method == 'GET':
        return 'a'
    if request.method == 'POST':
        asc = request.files['asc']
        asc_name=asc.filename
        print('asc::')
        print(asc)
        print(asc_name)

        html2 = request.files['html2']
        html2_name=html2.name
        print('html::')
        print(html2)
        print('1111111')
        print(html2.name)
        print(html2_name)
        if html2_name=='':
            return render_template('gateway.html',noresult=1,success=1)
        elif asc_name=='':
            return render_template('gateway.html',noasc=1,success=1)
        else:

            path = "static/xls/" + str(yonghu.progamname) + '/' + str(yonghu.version) + '/' + str(username) + '/' + yonghu.time + '/'
            asc_path = path + asc_name
            asc.save(asc_path)

            html_path= path+ 'result.html'
            html2.save(html_path)
            print('留档完成')

            return render_template('gateway.html',success_save=1,success=1)



@app.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'GET':
        return render_template('register.html')
    if request.method == 'POST':
        global username
        username = request.form.get('email1')
        password1 = request.form.get('password1')
        password2 = request.form.get('password2')
        if  password1!= password2:
            error1="两次输入密码不一致，请重新输入"


            return render_template('register.html',error1=error1)
        elif valid_regist(username):
            user = User(username=username, password=password1)
            db.session.add(user)
            db.session.commit()

            yonghu = User.query.filter(User.username == username).one()
            regtime = yonghu.regtime
            imagepath = yonghu.image
            return render_template('index.html',username=username,imagepath=imagepath,regtime=regtime)
        elif username=='':
            error2 = '用户名不能为空'
            return render_template('register.html', error2=error2)
        else:
            error3 = '该用户名已被注册'
            return render_template('register.html', error3=error3)

@app.route('/changefile', methods=['GET', 'POST'])
def changefile():
    global username
    yonghu = User.query.filter(User.username == username).one()
    if request.method == 'GET':
        return render_template('changefile.html',username=username)
    if request.method == 'POST':
        print('hello')
        f = request.files['htmltopdf']
        filename_html = f.filename
        print("文件名：")
        print(filename_html)

        path = "static/xls/pdf/"
        if os.path.exists(path):
            print('exist')
        else:
            os.mkdir(path)


        path = "static/xls/pdf/"+username+'/'
        if os.path.exists(path):
            print('exist')
        else:
            os.mkdir(path)

        path = "static/xls/pdf/" + username + '/'+yonghu.time+'/'
        if os.path.exists(path):
            print('exist')
        else:
            os.mkdir(path)

        file_path=path+f.filename
        print("file_path：:::")
        print(file_path)

        f.save(file_path)

        fnamenohtml=f.filename.split('.')
        filehhtml_path=path+fnamenohtml[0]+'.pdf'
        User.query.filter_by(username=username).update({'htmlpath': file_path})
        db.session.commit()
        html_to_pdf(file_path,filehhtml_path)
        return send_file(filehhtml_path, attachment_filename=str(f.filename.split('.')[0]+'.pdf'), as_attachment=True)

@app.route('/history', methods=['GET'])
def history():
    global username
    db.reflect()

    all_table = {table_obj.name: table_obj for table_obj in db.get_tables_for_bind()}
    print(all_table)
    a = db.session.query(all_table[username]).all()
    print('每行：：')
    allhisory=[]
    for row in a:
        row=list(row)

        row[3]=changedate(row[3])
        aa=row[4].split('/')
        row[4]=aa[-1]
        allhisory.append(row)
    print('让我check一下')
    print(allhisory)


    if request.method == 'GET':
        return render_template('history.html',username=username,allhisory=allhisory)

@app.route('/loadzip', methods=['GET'])
def loadzip():
    return 'aaa'
if __name__ == '__main__':
    flask_excel.init_excel(app)
    app.run(host='0.0.0.0',port=5000,debug=True)
