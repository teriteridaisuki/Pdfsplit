import PyPDF2
import os
import openpyxl


def pdfsplit(filename,data_use=False):#将名为filenname的pdf的全部页面拆分成单页pdf
    if data_use == True:
        ws_data = openpyxl.load_workbook('拆分顺序.xlsx')["Sheet1"]
    reader=PyPDF2.PdfFileReader(r"文件源/"+filename)
    nums=reader.numPages
    for i in range(nums):
        writer=PyPDF2.PdfFileWriter()
        page=reader.getPage(i)
        writer.addPage(page)
        if data_use==True:
            with open(r"输出\%s.pdf"%(ws_data.cell(i+1,1).value),"wb") as f:
                writer.write(f)
        else:
            with open(r"输出\%s%s.pdf"%(str(filename)[0:4],i+1),"wb") as f:
                writer.write(f)


filenames=os.listdir(r"文件源")
for filename in filenames:#将文件夹内的所有文件都拆分
    pdfsplit(filename,True)
    print(filename)
