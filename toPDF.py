import os
import shutil
from reportlab.pdfgen import canvas
import xlwt

def my_deal_file(file_name, changed_list, sheet_data): #only deal file in dir
    sheet_data_element = []
    tif_exit_flag = False
    
    if os.path.splitext(file_name)[1] in {'.zip','.ZIP'}:
        dealFileName = os.path.splitext(file_name)[0]
        print(dealFileName+'.zip')
        
        try:
            os.makedirs(dealFileName)                            #make unpackDir
        except:
            return
        
        try:
            shutil.unpack_archive(file_name, dealFileName)  #unpack file
        except:
            return

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
            #shutil.move(pdf_file_name,pdf_file_to_position)       #move pdf file                                     
            #shutil.make_archive(dealFileName,'zip', os.path.abspath(dealFileName)) #pack file
            changed_list.append(dealFileName+".zip")

        os.remove(file_name) #delete orginal file      

        shutil.rmtree(dealFileName) #remove the unpackFile

        sheet_data.append(sheet_data_element)


def final_deal(wating_deal_file, changed_list, sheet_data):

    root_dir = os.getcwd()

    file_list = os.listdir(root_dir)

    for file_name in file_list:
        if os.path.isfile(file_name):        
            my_deal_file(file_name, changed_list, sheet_data) 

        if os.path.isdir(file_name):
            os.chdir(file_name)
            sub_file_list = os.listdir(r".")
            final_deal(file_name, changed_list, sheet_data)
            os.chdir(root_dir)
        


file_list = os.listdir(r".")

print('待整理文件数：{0}\n'.format(len(file_list)))

print('已处理文件:')

changed_list = []
sheet_data = []
root_dir = os.getcwd()

final_deal(root_dir,changed_list, sheet_data)

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

for data in sheet_data:
    sheet.write(row,0,data[0])
    sheet.write(row,1,data[1])
    sheet.write(row,3,data[2])
    row = row+1

wbk.save('汇总.xls')
os.system("Pause")
