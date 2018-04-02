import os
import shutil
from reportlab.pdfgen import canvas
import xlwt
from appJar import gui

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

        if os.path.isdir(file_name):
            trench_name = os.path.abspath(file_name)
            os.chdir(trench_name)
            count_file_numbers(trench_name)
            os.chdir(wating_deal_folder)     

def deal_file(file_name, changed_list, sheet_data): #only deal file in dir
    
    global dealed_file_numbers
    sheet_data_element = []
    
    tif_exit_flag = False
    
    if os.path.splitext(file_name)[1] in {'.zip','.ZIP'}:
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


def deal_folder(wating_deal_folder, changed_list, sheet_data):

    os.chdir(wating_deal_folder)

    file_list = os.listdir(wating_deal_folder)

    for file_name in file_list:
        if os.path.isfile(file_name):        
            try:
                deal_file(file_name, changed_list, sheet_data)
            except:
                print("FOUND ERROR AT "+file_name)
                continue

        if os.path.isdir(file_name):
            trench_name = os.path.abspath(file_name)
            os.chdir(trench_name)
            deal_folder(trench_name, changed_list, sheet_data)
            os.chdir(wating_deal_folder)          

def press_select(button):
    if button=="button1":
        temp = app.directoryBox("Select a path")
        if temp:
            root_dir = temp
            app.setEntry("源文件夹路径", root_dir)

def press_action(button):
    if button=="开始":
        action()
        changed_list.clear()
        sheet_data.clear()
    if button=="清空":
        app.clearEntry("源文件夹路径")

def action():
    global total_file_numbers
    deal_dir= app.getEntry("源文件夹路径")
    file_list = os.listdir(deal_dir)
    os.chdir(deal_dir)
    print("正在统计需整理文件数目......")
    count_file_numbers(deal_dir)
    print('待整理文件数：{0}\n'.format(total_file_numbers))

    print('已处理文件:')

    deal_folder(os.path.abspath(deal_dir),changed_list, sheet_data)

    print("\n以下专利被整理：")

    for i in changed_list:
        print(i)

    print("\n已创建Excel:汇总.xls\n")

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
        wbk.save('汇总.xls')
        print('任务完成!!!!!!!!')
    total_file_numbers=0
    dealed_file_numbers=0

total_file_numbers=0
dealed_file_numbers=0
root_dir = os.getcwd()
changed_list = []
sheet_data = []
app = gui("to PDF","450x70")
app.setResizable(False)

app.addLabelEntry("源文件夹路径",0,0)
app.setEntry("源文件夹路径",root_dir)
app.addNamedButton("选择","button1",press_select,0,1)

app.addButtons(["开始", "清空"], press_action)
app.go()
