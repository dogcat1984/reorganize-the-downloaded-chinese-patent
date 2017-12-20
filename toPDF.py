import os
import shutil
from reportlab.pdfgen import canvas
import xlwt

file_list = os.listdir(r".")

wbk = xlwt.Workbook()
sheet = wbk.add_sheet('sheet 1')
sheet.write(0,0,'专利名称')#第0行第一列写入内容
sheet.write(0,1,'申请号')
sheet.write(0,2,'申请日')
sheet.write(0,3,'公开号')
sheet.write(0,4,'公开日')
row = 1

print('已处理文件')

changed_List = []

for file_name in file_list:

    tif_exit_flag = False
    
    if os.path.splitext(file_name)[1] == '.zip':
        dealFileName = os.path.splitext(file_name)[0]
        print(dealFileName+'.zip')
        os.makedirs(dealFileName)                            #make unpackDir
        shutil.unpack_archive(file_name, os.path.abspath(dealFileName))  #unpack file
        sheet.write(row,0,dealFileName) #add title to xls

        pdf_file_name = os.path.abspath(dealFileName)+".pdf"        
        pdf_file = canvas.Canvas(pdf_file_name) #creat pdf        
        
        mdir_list = os.listdir(os.path.abspath(dealFileName)) #write to excel
        for mdir in mdir_list:
            if os.path.isdir((os.path.join(os.path.abspath(dealFileName), mdir))):
                number_list = mdir.split('_')
                if len(number_list)>1:
                    sheet.write(row,1,number_list[0])
                    sheet.write(row,3,number_list[1])
                row=row+1
            
        for root, dirs, files in os.walk(os.path.abspath(dealFileName)):       
            for file in files:
                final_file = os.path.join(root, file)
                if os.path.splitext(final_file)[1] == '.tif'or os.path.splitext(os.path.join(root, file))[1] == '.TIF' or os.path.splitext(os.path.join(root, file))[1] == '.Tif':
                    tif_exit_flag = True;
                    pdf_file.drawImage(final_file,0,0,590,892,None,True,'c')
                    pdf_file.showPage()       #save current pdf page
                    pdf_file_to_position = root

                if os.path.splitext(final_file)[1] == '.pdf':
                    target_file = dealFileName+'.pdf'
                    shutil.move(final_file,target_file) #remove pdf to root
                   
        if tif_exit_flag:
            pdf_file.save() #save pdf file
            #shutil.move(pdf_file_name,pdf_file_to_position)       #move pdf file                                     
            #shutil.make_archive(dealFileName,'zip', os.path.abspath(dealFileName)) #pack file
            changed_List.append(dealFileName+".zip")

        os.remove(file_name) #delete orginal file      

        shutil.rmtree(dealFileName) #remove the unpackFile

print("\n以下专利被整理：")

for i in changed_List:
    print(i)

print("\n已创建Excel:汇总.xls\n")

wbk.save('汇总.xls')
os.system("Pause")     
