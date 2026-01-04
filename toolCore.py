import os

from docx import Document
from win32com import client as wc
def one_err_fun(res,e):
    print(res)
    print(e)

def get_files_by_extension(directory, extension):
    files_with_extension = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith(extension):
                if file.startswith("~$"):
                    print("过滤缓存文件")
                    continue
                file_path = os.path.join(root, file)
                files_with_extension.append(file_path)

    return files_with_extension

def word_tables_to_list(file_path):
    res=[]
    doc=Document(file_path)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                t=cell.text.replace("\n",'').replace(" ",'').replace("\t","").strip()
                res.append(t)
    return res

def TransDocToDocx(oldDocName,newDocxName):
    word = wc.Dispatch('Word.Application')
    doc = word.Documents.Open(oldDocName)
    doc.SaveAs(newDocxName, 12)
    doc.Close()
    word.Quit()
def TransPdfToDocx(oldPdfmae,newDocNmae):
    from pdf2docx import Converter
    a = Converter(oldPdfmae)
    a.convert(newDocNmae)
    a.close()

#从cells提取指定信息
def get_info_from_cells(cells_list,key_list):
    info={}
    for i in range(len(cells_list)):
        for key in key_list:
            if key==cells_list[i]:
                if i+1>=len(cells_list):
                    index=i
                else:
                    index=i+1
                info[key]=cells_list[index]
    return info






