import os
import shutil
from reportlab.pdfgen import canvas
import xlwt
from appJar import gui
import datetime

global total_file_numbers
global dealed_file_numbers

def count_file_numbers(wating_deal_folder):

    global total_file_numbers

    os.chdir(wating_deal_folder)

    file_list = os.listdir(wating_deal_folder)

    for file_name in file_list:
        if os.path.isfile(file_name):
            if os.path.splitext(file_name)[1] in {'.zip','.ZIP'}:
                total_file_numbers+=1
                to_deal_file_list.append(os.path.abspath(file_name))

        if os.path.isdir(file_name):
            trench_name = os.path.abspath(file_name)
            count_file_numbers(trench_name)
            os.chdir(wating_deal_folder)     

def deal_file(file_absDir): #only deal file in dir
    
    global dealed_file_numbers
    sheet_data_element = []

    file_dir = os.path.dirname(file_absDir)
    file_name = os.path.basename(file_absDir)

    os.chdir(file_dir)
    
    tif_exit_flag = False
    
    dealFileName = os.path.splitext(file_name)[0]
    dealed_file_numbers+=1
        
    print('{0}/{1} {2}'.format(dealed_file_numbers, total_file_numbers, file_name))
        
    os.makedirs(dealFileName)                            #make unpackDir
        
    shutil.unpack_archive(file_name, dealFileName)  #unpack file

    sheet_data_element.append(dealFileName)

    pdf_file_name = dealFileName+".pdf"        
    pdf_file = canvas.Canvas(pdf_file_name) #creat pdf        
        
    mdir_list = os.listdir(dealFileName) #write to excel_data
    for mdir in mdir_list:
        if os.path.isdir(os.path.join(dealFileName, mdir)):
            number_list = mdir.split('_')
            if len(number_list)>1:
                sheet_data_element.append(number_list[0])
                sheet_data_element.append(number_list[1])                
            
    for root, dirs, files in os.walk(dealFileName):       
        for file in files:
            final_file = os.path.join(root, file)
            if os.path.splitext(final_file)[1] in {'.tif', '.TIF'}:
                tif_exit_flag = True;
                pdf_file.drawImage(final_file,0,0,590,892,None,True,'c')
                pdf_file.showPage()       #save current pdf page

            if os.path.splitext(final_file)[1] in {'.pdf','.PDF'}:
                target_file = dealFileName+'.pdf'
                shutil.move(final_file,target_file) #move pdf to root
                   
    if tif_exit_flag:
        pdf_file.save() #save pdf file
        changed_list.append(dealFileName+".zip")

    os.remove(file_name) #delete orginal file      

    shutil.rmtree(dealFileName) #remove the unpackFile

    sheet_data.append(sheet_data_element)


def deal_folder():
    for file in to_deal_file_list:
        try:
            deal_file(file)
        except:
            print('FOUND ERROR AT '+file)
                        

def press_select(button):
    if button=="button1":
        temp = app.directoryBox("Select a path")
        if temp:
            app.setEntry("文件夹路径", temp)
    if button=="button2":
        temp = app.openBox("Select a file",fileTypes=[("packfiles",".zip"),(("packfiles",".ZIP"))])
        if temp:
            app.setEntry("文件路径", temp)

def press_action(button):
    if button=="开始":
        flag=0
        if app.getEntry("文件夹路径"):
            process_folder()
            flag=1
        if app.getEntry("文件路径"):
            process_file()
            flag=1
        if flag:
            display_dealed_files()
            write_to_excel()       
    if button=="清空":
        app.clearEntry("文件夹路径")
        app.clearEntry("文件路径")

def process_file():
    global total_file_numbers
    global dealed_file_numbers
    total_file_numbers+=1
    file = app.getEntry("文件路径")
    try:
        deal_file(file)
        total_file_numbers = 0
        dealed_file_numbers = 0
    except:
        print('FOUND ERROR AT '+file)    

def process_folder():
    global total_file_numbers
    deal_dir= app.getEntry("文件夹路径")
    file_list = os.listdir(deal_dir)
    os.chdir(deal_dir)
    print("正在统计需整理文件数目......")
    count_file_numbers(deal_dir)
    print('待整理文件数：{0}\n'.format(total_file_numbers))

    print('已处理文件:')

    deal_folder()

    os.chdir(deal_dir)

def display_dealed_files():
    print("\n以下专利被整理：")
    for i in changed_list:
        print(i)
    changed_list.clear()    

def write_to_excel():
    nowTime=datetime.datetime.now().strftime('%Y-%m-%d %H-%M-%S')
    print("\n已创建Excel:汇总"+nowTime+".xls\n")

    wbk = xlwt.Workbook()
    sheet = wbk.add_sheet('sheet 1')
    sheet.write(0,0,'专利名称')#第0行第一列写入内容
    sheet.write(0,1,'申请号')
    sheet.write(0,2,'申请日')
    sheet.write(0,3,'公开号')
    sheet.write(0,4,'公开日')
    row = 1

    try:
        for data in sheet_data:
            sheet.write(row,0,data[0])
            sheet.write(row,1,data[1])
            sheet.write(row,3,data[2])
            row = row+1
    except:
        pass
    finally:
        wbk.save('汇总'+nowTime+'.xls')
        print('任务完成!!!!!!!!')
    total_file_numbers=0
    dealed_file_numbers=0
    to_deal_file_list.clear()    
    sheet_data.clear()

total_file_numbers=0
dealed_file_numbers=0
to_deal_file_list = []
changed_list = []
sheet_data = []
app = gui("to PDF")
app.setResizable(False)

app.addLabelEntry("文件夹路径",0,0)
app.addNamedButton("选择","button1",press_select,0,1)
app.addLabelEntry("文件路径",1,0)
app.addNamedButton("选择","button2",press_select,1,1)

app.addButtons(["开始", "清空"], press_action)
app.go()
