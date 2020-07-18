import pypyodbc,sqlite3
import sys,os,datetime,shutil,hashlib,json,re
from tkinter import Tk
import tkinter.filedialog
from PIL import Image
from PIL import ImageDraw
from PIL import ImageFont
#链接mdb
def connect_db(filename):
    mdb = 'Driver={Microsoft Access Driver (*.mdb,*.accdb)};DBQ='+filename
    conn = pypyodbc.win_connect_mdb(mdb)
    return conn
def sort_key(s):
    #sort_strings_with_embedded_numbers
    re_digits = re.compile(r'(\d+)')
    pieces = re_digits.split(s)  # 切成数字与非数字
    pieces[1::2] = map(int, pieces[1::2])  # 将数字部分转成整数
    return pieces
#解包数据库中的pak文件
def unpack_db_pak(filename,src,dest):
    c=connect_db(filename)
    sql="Select ID,ScreenShot From ScreenShot"
    s=c.cursor()
    res=s.execute(sql)
    row=res.fetchone()
    while(row!=None):
        if(row[1]!=None):
            print("正在提取："+str(row[0])+".pak")
            fiobj=open(os.path.join(src,str(row[0])+".pak"),"wb+")
            fiobj.write(row[1])
            fiobj.close()
            unpak_pak(os.path.join(src,str(row[0])+".pak"),os.path.join(dest,str(row[0])))
        row=res.fetchone()
    row=None
    res=None
    s.close()
    c.close()
#解包PAK
def unpak_pak(pakname,unpakpath):
    if os.path.getsize(pakname)>0: 
        print("正在解包pak:"+pakname)
        print(os.path.join(os.path.dirname(os.path.realpath(sys.argv[0])),"PAKLoader.exe")+" "+pakname+" --extract "+unpakpath)
        os.system(os.path.join(os.path.dirname(os.path.realpath(sys.argv[0])),"PAKLoader.exe")+" "+pakname+" --extract "+unpakpath)
        print("解包完成pak:"+pakname)
#转换基础数据库结构
def get_codename(srcfile,buildname):
    c_mdb=connect_db(srcfile)
    s_mdb=c_mdb.cursor()
    print("正在查找CodeName："+str(buildname))
    res1=s_mdb.execute("SELECT DISTINCT ProductName,CodeName FROM OSInformation Where ProductName='"+buildname+"';")
    row1=res1.fetchall()
    temp_code=""
    print(row1)
    for i in range(0,len(row1)):
        if(row1[i][1]!="N/A"):
            temp_code=temp_code+row1[i][1]+","
    temp_code=temp_code[0:-1]
    if temp_code=="":
        temp_code=="null"
    print("找到CodeName："+str(temp_code))
    s_mdb.close()
    c_mdb.close()
    return temp_code
def get_ProductID(destfile,buildname):
    c_sl3=sqlite3.connect(destfile)
    s_sl3 = c_sl3.cursor()
    cussor=s_sl3.execute("select ID from Product where Name='"+buildname+"'")
    for row in cussor:
        return row[0]
    return -1
def get_ProductName(destfile,buildID):
    c_sl3=sqlite3.connect(destfile)
    s_sl3 = c_sl3.cursor()
    cussor=s_sl3.execute("select Name from Product where ID='"+str(buildID)+"'")
    for row in cussor:
        return row[0]
    return -1
def BuildID_to_ProductID(destfile,buildID):
    c_sl3=sqlite3.connect(destfile)
    s_sl3 = c_sl3.cursor()
    cussor=s_sl3.execute("select ProductID from Build where ID='"+str(buildID)+"'")
    for row in cussor:
        return row[0]
    return -1
def OldBuildID_to_NewBuildID(srcfile,destfile,oldproid,oldbuildid):
    c_mdb=connect_db(srcfile)
    s_mdb=c_mdb.cursor()
    c_sl3=sqlite3.connect(destfile)
    s_sl3 = c_sl3.cursor()
    res1=s_mdb.execute("SELECT Version,Stage,BuildTag,ProductName FROM OSInformation Where ProductID="+str(oldproid)+" and BuildID="+str(oldbuildid)+";")
    row1=res1.fetchone()
    if row1==None:
        return -1
    cussor=s_sl3.execute("select ID from Build where Version='"+row1[0]+"' and Stage='"+row1[1]+"' and BuildTag='"+row1[2]+"' and ProductID="+str(get_ProductID(destfile,row1[3])))
    for row in cussor:
        return row[0]
    return -1
def NewBuildID_to_ScreenshotID(srcfile,destfile,newbuildid):
    c_mdb=connect_db(srcfile)
    s_mdb=c_mdb.cursor()
    c_sl3=sqlite3.connect(destfile)
    s_sl3 = c_sl3.cursor()
    res1=s_sl3.execute("SELECT Version,Stage,BuildTag,Date,ProductID FROM Build Where ID="+str(newbuildid)+";")
    for row1 in res1:
        cussor=s_mdb.execute("select ScreenShotID from OSInformation where ScreenShotID>0 and Version='"+row1[0]+"' and Stage='"+row1[1]+"' and BuildTag='"+row1[2]+"' and BIOSDate=#"+row1[3]+"# and ProductName='"+get_ProductName(destfile,row1[4])+"'")
        print("select ScreenShotID from OSInformation where ScreenShotID>0 and Version='"+row1[0]+"' and Stage='"+row1[1]+"' and BuildTag='"+row1[2]+"' and Date='"+row1[3]+"'")
    row=cussor.fetchall()
    if len(row)>0:
        return row[0][0]
    return -1
def ScreenShotID_to_NewBuildID(srcfile,destfile,screenid):
    c_mdb=connect_db(srcfile)
    s_mdb=c_mdb.cursor()
    res1=s_mdb.execute("SELECT ProductID,BuildID FROM OSInformation Where ScreenShotID="+str(screenid)+";")
    row1=res1.fetchone()
    if(row1==None):
        return -1
    return OldBuildID_to_NewBuildID(srcfile,destfile,row1[0],row1[1])
def convert_db(srcfile,destfile):
    c_mdb=connect_db(srcfile)
    c_sl3=sqlite3.connect(destfile)
    s_sl3 = c_sl3.cursor()
    s_mdb=c_mdb.cursor()
    #创建基础数据库结构
    s_sl3.execute("create table 'Contributor' ('ID' integer primary key autoincrement, 'Name' text not null unique);")
    s_sl3.execute("create table 'Version' ('Version' text not null, 'Date' date not null, 'UpdateURL' text not null);")
    s_sl3.execute("create table 'Product' ('ID' integer primary key autoincrement, 'Name' text not null unique, 'Codename' text);")
    s_sl3.execute("create table 'Build' ('ID' integer primary key autoincrement, 'ProductID' integer not null, 'Version' text not null, 'Stage' text not null, 'Buildtag' text not null, 'Architecture' text not null, 'Edition' text, 'Language' text not null, 'Date' date not null, 'Serial' text, 'Notes' text, 'NotesEN' text, 'CodeName' text, foreign key ('ProductID') references 'Product'('ID') on update cascade on delete cascade);")
    s_sl3.execute("create table 'ChangeLog' ('Version' text not null, 'BuildID' integer not null, foreign key ('BuildID') references 'Build'('ID') on update cascade on delete cascade);")
    #转换Contributor
    res=s_mdb.execute("select ID,ContributorName from ContributorList")
    row=res.fetchone()
    while(row!=None):
        print("正在转换Contributor："+str(row[0])+"/"+str(row[1]))
        #print("insert into 'Contributor' values (null,'"+row[1]+"');")
        s_sl3.execute("insert into 'Contributor' values (null,'"+row[1]+"');")
        print("转换成功："+str(row[1]))
        row=res.fetchone()
    c_sl3.commit()
    #转换ProductList
    res=s_mdb.execute("select ProductName from ProductList Order by ID")
    row=res.fetchone()
    while(row!=None):
        print("正在转换ProductList："+str(row[0]))
        #print("insert into 'Contributor' values (null,'"+row[1]+"');")
        temp_code=get_codename(srcfile,row[0])
        s_sl3.execute("insert into 'Product' values (null,'"+row[0]+"','"+temp_code+"');")
        print("转换成功："+str(row[0]))
        row=res.fetchone()
    c_sl3.commit()
    #转换BuildList
    res=s_mdb.execute("SELECT ProductName,Stage,Version,BuildTag,Architecture,Edition,Language,BIOSDate,SerialNumber,Fixes,FixesEN,CodeName FROM OSInformation order by ProductID,BuildID")
    row=res.fetchone()
    while(row!=None):
        print("正在转换BuildList："+str(row))
        proid=get_ProductID(destfile,row[0])
        print("insert into 'Build' values (null, "+str(proid)+", '"+row[2]+"', '"+row[1]+"', '"+row[3]+"', '"+row[4]+"', '"+row[5]+"', '"+row[6]+"', '"+datetime.datetime.strftime(row[7], "%Y/%m/%d")+"', '"+row[8]+"', '"+str(row[9]).replace("'","''")+"', '"+str(row[10]).replace("'","''")+"');")
        s_sl3.execute("insert into 'Build' values (null, "+str(proid)+", '"+row[2]+"', '"+row[1]+"', '"+row[3]+"', '"+row[4]+"', '"+row[5]+"', '"+row[6]+"', '"+datetime.datetime.strftime(row[7], "%Y/%m/%d")+"', '"+row[8]+"', '"+str(row[9]).replace("'","''")+"', '"+str(row[10]).replace("'","''")+"','"+str(row[11]).replace("'","''")+"');")
        print("转换成功："+str(proid))
        row=res.fetchone()
    c_sl3.commit()
    #转换ChangeLog
    res=s_mdb.execute("SELECT ProductID,BuildID FROM ChangeLog")
    row=res.fetchone()
    while(row!=None):
        print("正在转换ChangeLOG："+str(row))
        new_build=OldBuildID_to_NewBuildID(srcfile,destfile,row[0],row[1])
        s_sl3.execute("insert into 'ChangeLog' values ('todo', "+str(new_build)+");")
        print("转换成功："+str(new_build))
        row=res.fetchone()
    c_sl3.commit()
    #转换Version
    print("处理数据库版本信息...")
    res=s_mdb.execute("SELECT * FROM VersionInformation")
    row=res.fetchall()
    db_ver=row[0][0]
    db_date=row[0][1]
    update_address=row[3][0]
    sql="insert into 'Version' values ('"+db_ver+"', '"+datetime.datetime.strftime(db_date, "%Y-%m-%d")+"', '"+update_address+"');"
    s_sl3.execute(sql)
    c_sl3.commit()
    sql="update 'ChangeLog' SET Version='"+db_ver+"'"
    s_sl3.execute(sql)
    c_sl3.commit()
    print("处理数据库版本信息完成，基础数据转换完成...")
#给图片打水印
def watermark_img(src,dest):
    watermark = Image.open("logo.png")
    imageFile = Image.open(src)
    layer = Image.new('RGBA', imageFile.size, (0,0,0,0))
    #(image_x - watermark_x, image_y - watermark_y)
    layer.paste(watermark, (imageFile.size[0]-watermark.size[0], imageFile.size[1]-watermark.size[1]))
    out=Image.composite(layer,imageFile,layer)
    out.save(dest)
def do_picture(src,dest,mdb,sl3):
    for root, dirs, files in os.walk(src):
        for d in dirs:
            print("正在复制解包的图片文件夹:"+str(d))
            screenid=ScreenShotID_to_NewBuildID(mdb,sl3,int(d))
            print("对应的截图ID:"+str(screenid))
            if screenid!=-1:
                shutil.copytree(os.path.join(src,str(d)),os.path.join(dest,str(screenid)))
                fiobj=open(os.path.join(dest,str(screenid)+"\\"+str(d)),"w+")
                fiobj.close()
            print("复制完成:"+str(screenid))
def put_watermark_on_pic(dest):
    for root, dirs, files in os.walk(dest):
        for f in files:
            path1=os.path.join(root, f)
            print("正在处理图片:"+str(path1))
            f1=os.path.splitext(f)[0]
            e=os.path.splitext(f)[1]
            if not ("_w" in f1):
                if e=='.png':
                    fn=f1+"_w"+e
                    path2=os.path.join(root, fn)
                    print(path1+"->"+path2)
                    watermark_img(path1,path2)
                    os.remove(path1)
            print("处理图片完成:"+str(path1))
def GetFileMd5(filename):
    if not os.path.isfile(filename):
        return
    myHash = hashlib.md5()
    f = open(filename,'rb')
    while True:
        b = f.read(8096)
        if not b :
            break
        myHash.update(b)
    f.close()
    return myHash.hexdigest()
def make_pic_json(src,buildid,dbfile):
    filelist=os.listdir(src)
    filelist.sort(key=sort_key) #重排序
    i=0
    mainpic=-1
    a={"productid": BuildID_to_ProductID(dbfile,buildid),"buildid": buildid,"main_pic":0,"screenshot":[]}
    for f in filelist:
        f1=os.path.splitext(f)[0]
        e=os.path.splitext(f)[1]
        if e==".png":
            dict_d={'screenshotid':i,"screenshothash":GetFileMd5(os.path.join(src,f)),"screenshottitle": f1[0:-2],"parentid": buildid}
            a['screenshot'].append(dict_d)
            if f1[0:-2].lower()=="version":
                mainpic=i
            if (f1[0:-2].lower()=="interface 1") and (mainpic==-1):
                mainpic=i
            i=i+1
    if (mainpic==-1):
        mainpic=0
    a['main_pic']=mainpic
    #print(a)
    json_a=json.dumps(a)
    print(json_a)
    fiobj=open(os.path.join(src,"pic.json"),"w")
    fiobj.write(json_a)
    fiobj.close()
def make_and_change(src,dbfile):
    for root, dirs, files in os.walk(src):
        for d in dirs:
            print("正在写入描述文件："+str(d))
            make_pic_json(os.path.join(src,d),int(d),dbfile)
            print("写入描述文件完成："+str(d))
            filelist=os.listdir(os.path.join(src,d))
            print("重命名文件："+os.path.join(src,d))
            for f in filelist:
                f1=os.path.splitext(f)[0]
                e=os.path.splitext(f)[1]
                if e==".png":
                    pic_hash=GetFileMd5(os.path.join(os.path.join(src,d),f))
                    shutil.move(os.path.join(os.path.join(src,d),f),os.path.join(os.path.join(src,d),pic_hash+e))
            print("重命名文件完成："+os.path.join(src,d))
def make_info_json(srcfile,destfolder):
    print("正在写入图片储存库描述文件....")
    c_mdb=connect_db(srcfile)
    s_mdb=c_mdb.cursor()
    res=s_mdb.execute("SELECT * FROM VersionInformation")
    row=res.fetchall()
    db_ver=row[0][0]
    db_date=datetime.datetime.strftime(row[0][1], "%Y-%m-%d")
    a={"GalleryVersion":db_ver+"P","GalleryUpdateTime":db_date}
    json_a=json.dumps(a)
    print(json_a)
    fiobj=open(os.path.join(destfolder,"info.json"),"w")
    fiobj.write(json_a)
    fiobj.close()
    print("写入图片储存库描述文件完成....")
def zip_file(src):
    for root, dirs, files in os.walk(src):
        for d in dirs:
            print("正在压缩版本包:"+d)
            sz_filename=os.path.join(src,d+".zip")
            sz_filepath=os.path.join(src,d+"\\*.*")
            zip_command=" a "+sz_filename+" "+sz_filepath+" -mx=9 -mmt -r"
            os.system(os.path.join(os.path.dirname(os.path.realpath(sys.argv[0])),"7z.exe")+zip_command)
            shutil.rmtree(os.path.join(src,d))
            print("压缩版本包完成:"+d)
def openfilediaglog(title,filefilter,initdir):
    Tk().withdraw()
    filename = tkinter.filedialog.askopenfilename(title=title, filetypes=filefilter,initialdir=initdir)
    return filename
def openfolderdiaglog(title):
    Tk().withdraw()
    filename = tkinter.filedialog.askdirectory(title=title)
    return filename
def create_path(path):
    if os.path.exists(path):
        shutil.rmtree(path)
    os.mkdir(path)
def rm_path(path):
    if os.path.exists(path):
        shutil.rmtree(path)
print("当前运行路径："+os.path.dirname(os.path.realpath(sys.argv[0])))
srcfile=openfilediaglog("请选择已经解密的BWDB的文件数据库：",[("DB文件","*.db"),("MDB文件","*.mdb"),("所有文件","*.*")],os.path.dirname(os.path.realpath(sys.argv[0])))
if srcfile=="":
    exit()
print("你选择的BWDB的文件数据库路径："+srcfile)
destfolder=openfolderdiaglog("请选择新版本数据库的输出路径：")
if destfolder=="":
    exit()
print("你选择的输出路径："+destfolder)
input("确认后按任意键开始转换数据...")
print("正在创建临时环境...")
create_path(os.path.join(destfolder,"pak"))
create_path(os.path.join(destfolder,"unpak"))
create_path(os.path.join(destfolder,"gallery"))
if os.path.exists(os.path.join(destfolder,"DataBase.db")):
    os.remove(os.path.join(destfolder,"DataBase.db"))
print("开始转化基础数据，这需要一定的时间...")
convert_db(srcfile,os.path.join(destfolder,"DataBase.db"))
print("正在提取PAK文件,这可能需要很长的一段时间（视数量决定）...")
unpack_db_pak(srcfile,os.path.join(destfolder,"pak"),os.path.join(destfolder,"unpak"))
print("正在处理图片文件,这可能需要很长的一段时间（视数量决定）...")
do_picture(os.path.join(destfolder,"unpak"),os.path.join(destfolder,"gallery"),srcfile,os.path.join(destfolder,"DataBase.db"))
put_watermark_on_pic(os.path.join(destfolder,"gallery"))
#make_pic_json(os.path.join(os.path.dirname(os.path.realpath(sys.argv[0])),"gallery/1"),1)
make_and_change(os.path.join(destfolder,"gallery"),os.path.join(destfolder,"DataBase.db"))
make_info_json(srcfile,os.path.join(destfolder,"gallery"))
#print(NewBuildID_to_ScreenshotID("D:\\BW_wiki\\BetaWorld Library 11.43.2\\OSInformationA.mdb",os.path.join(os.path.dirname(os.path.realpath(sys.argv[0])),"DataBase.db"),193))
print("正在打包图片储存库,这可能需要很长的一段时间（视数量决定）...")
zip_file(os.path.join(destfolder,"gallery"))
print("正在清理临时文件...")
rm_path(os.path.join(destfolder,"pak"))
rm_path(os.path.join(destfolder,"unpak"))
print("处理完成！")
print("信息数据库位置："+os.path.join(destfolder,"DataBase.db"))
print("图片数据库位置："+os.path.join(destfolder,"gallery"))
input("按任意键退出...")
#unpak_pak(os.path.join(os.path.dirname(os.path.realpath(sys.argv[0])),"pak\\1.pak"),os.path.join(os.path.dirname(os.path.realpath(sys.argv[0])),"unpak\\1"))
#待做
#print(OldBuildID_to_NewBuildID("D:\\BW_wiki\\BetaWorld Library 11.43.2\\OSInformationA.mdb",os.path.join(os.path.dirname(os.path.realpath(sys.argv[0])),"DataBase.db"),40,0))