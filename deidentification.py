import os
import re
import win32com.client
import win32gui
from docx2pdf import convert
import docx2txt
from fpdf import FPDF
import openpyxl as op
import pandas as pd

from pdfminer.high_level import extract_text

import torch
from transformers import BertTokenizer

from models.bert import BERT_NER

import subprocess

def getData():
    powerpoint = win32com.client.Dispatch('Powerpoint.Application')
    excel = win32com.client.Dispatch('Excel.Application')
    hwp = win32com.client.gencache.EnsureDispatch("HWPFrame.HwpObject")
    pdf = FPDF('L')

    docs_path = os.path.join(os.getcwd(), 'docs')
    pdfs_path = os.path.join(os.getcwd(), 'pdfs')
    txts_path = os.path.join(os.getcwd(), 'txts')

    if not os.path.isdir(docs_path):
        os.makedirs(docs_path)

    if not os.path.isdir(pdfs_path):
        os.makedirs(pdfs_path)
        
    if not os.path.isdir(txts_path):
        os.makedirs(txts_path)

    doc_files = [f for f in os.listdir(docs_path)]
    ppt_files = []
    word_files = []
    excel_files = []
    hwp_files = []
    img_files = []


    for file in doc_files:
        print(file)
        if re.match('.*([.]ppt|[.]pptx)', file):
            ppt_files.append(file)
        elif re.match('.*([.]doc|[.]docx)', file):
            word_files.append(file)
        elif re.match('.*([.]xls|[.]xlsx)', file):
            excel_files.append(file)
        elif re.match('.*[.]hwp', file):
            hwp_files.append(file)
        elif re.match('.*([.]jpg|[.]png|[.]gif|[.]jpeg)', file):
            img_files.append(file)


    #PPT to PDF
    for file in ppt_files:
        pre, ext = os.path.splitext(file)
        deck = powerpoint.Presentations.Open(os.path.join(docs_path, file))
        deck.SaveAs(os.path.join(pdfs_path, pre + '.pdf'), 32)  # formatType = 32 for ppt to pdf
        deck.Close()
    powerpoint.Quit()


    #Word to PDF
    for file in word_files:
        pre, ext = os.path.splitext(file)
        convert(os.path.join(docs_path, file), os.path.join(pdfs_path, pre + '.pdf'))


    #Excel to PDF
    for file in excel_files:
        pre, ext = os.path.splitext(file)

        file_with_sheet = []
        wb = op.load_workbook(os.path.join(docs_path, file)) 
        ws_list = wb.sheetnames
        wb.close()

        wb = excel.Workbooks.Open(os.path.join(docs_path, file)) 
        for sheet in ws_list:
            ws = wb.Worksheets(sheet)
            ws.Select()
            ## 임시방편
#             wb.ActiveSheet.ExportAsFixedFormat(0, os.path.join(pdfs_path, pre + '_' + sheet + '_' + '.pdf'))
            wb.ActiveSheet.ExportAsFixedFormat(0, os.path.join(pdfs_path, pre + '.pdf'))
        wb.Close() 
    excel.Quit()


    #HWP to PDF
    for file in hwp_files:
        pre, ext = os.path.splitext(file)
        hwp.Open(os.path.join(docs_path, file))
        hwp.SaveAs(os.path.join(pdfs_path, pre + ".pdf"), "PDF")
    hwp.Quit()


    #Image to PDF
    for file in img_files:
        pre, ext = os.path.splitext(file)
        pdf.add_page()
        pdf.image(os.path.join(docs_path,file), 0, 0, 330)
        pdf.output(os.path.join(pdfs_path, pre + '.pdf'), 'F')
    
    
    # docs to txt
    # Word to txt
    for file in word_files:
        word2txt(docs_path, txts_path, file)
        
    # Excel to txt
    for file in excel_files:
        xlsx2txt(docs_path, txts_path, file)
    
    # hwp to txt
    ## hwp를 pdf로 변환한 파일을 사용
    for file in hwp_files:
        pre, ext = os.path.splitext(file)
        hwp2txt(pdfs_path, txts_path, pre)
        

def word2txt(input_path, output_path, file_name):
    _file = docx2txt.process(input_path + "/" + file_name).split("\n")
    f = open(output_path + "/" + file_name.split(".")[0] + ".txt", "w")
    
    for line in _file:
        if line == "":
            continue
        else:
            for text in line.split():
                f.write(text+"\n")
    
    f.close()
    
def xlsx2txt(input_path, output_path, file_name):
    with open(output_path+"/"+"tmp.txt", "w") as file:
        pd.read_excel(input_path+"/"+file_name).to_string(file, index=False)
        
    _file = open(output_path+"/"+"tmp.txt")
    file_name = file_name.split(".")
    file_name = file_name[0:len(file_name)-1]
    file_name = "".join(file_name)
    f = open(output_path + "/" + file_name + ".txt", "w")
    for line in _file:
        for text in line.split():
            if(text == "Unnamed:"):
                break
            elif(text == "NaN"):
                continue
            elif("\\n\\n" in text):
                for n_text in text.split("\\n\\n"):
                    f.write(n_text+"\n")
            elif(text != "\n"):
                text = text.lstrip("\\n")
                text = text.rstrip("\\n")
                f.write(text+"\n")
    _file.close()
    f.close()
    
    os.remove(output_path+"/"+"tmp.txt")
    
def hwp2txt(input_path, output_path, file_name):
    _file = extract_text(input_path + "/" + file_name + ".pdf").split("\n")
    f = open(output_path + "/" + file_name + ".txt", "w")
    
    for line in _file:
        if line == "":
            continue
        else:
            for text in line.split():
                try:
                    f.write(text + "\n")
                except UnicodeEncodeError:
                    continue

                    
def evaluate():
    txts_path = os.path.join(os.getcwd(), 'txts')
    results_path = os.path.join(os.getcwd(), 'results')
    
    if not os.path.isdir(results_path):
        os.makedirs(results_path)
        
    tokenizer = BertTokenizer.from_pretrained("klue/bert-base")
    model = BERT_NER()
    model.load_state_dict(torch.load("BERT_NER_epoch_10.pt", map_location="cpu"))
    
    txts_files = [f for f in os.listdir(txts_path)]
    
    for _file in txts_files:
        tmp = []
        for line in open(os.path.join(txts_path, _file)).readlines():
            tmp.append(line.rstrip("\n"))
            
        sentence = " ".join(tmp)
        input_ids = tokenizer(sentence, return_tensors="pt").input_ids
        
        model.eval()
        with torch.no_grad():
            pred = model.predict(input_ids)

        tokens = tokenizer.convert_ids_to_tokens(input_ids[0])
        pred_tag = [model.tag_list[tag_id.item()] for tag_id in pred[0]]
        
        f = open(os.path.join(results_path, _file),"w")
        
        model.eval()
        with torch.no_grad():
            pred = model.predict(input_ids)

        tokens = tokenizer.convert_ids_to_tokens(input_ids[0])
        pred_tag = [model.tag_list[tag_id.item()] for tag_id in pred[0]]

        for token, tag in zip(tokens, pred_tag):
            f.write(str(token)+"\t"+str(tag)+"\n")
            
        f.close()
        
def deidentification():
    pdfs_path = os.path.join(os.getcwd(), "pdfs")
    results_path = os.path.join(os.getcwd(), 'results')
        
    pdfs_files = [f for f in os.listdir(pdfs_path)]
    
    for pdf_file in pdfs_files:
        result_file = open(os.path.join(results_path, pdf_file.split(".")[0])+".txt")
        tmp = ""
        check_point = False
        for result in result_file.readlines():
            result = result.rstrip("\n").split("\t")
            if result[1] == "NUM-B" or result[1] == "NUM-I" or result[1] == "LOC-I" or result[1] == "LOC-B" or result[1] == "ORG-I" or result[1] == "ORG-B":
                if tmp == "":
                    tmp += result[0]
                elif "##" in result[0]:
                    tmp += result[0].lstrip("##")
                elif "##" not in result[0]:
                    subprocess.run(['python', 'pdf_highlighter.py', '-i', os.path.join(pdfs_path, pdf_file), "-a", "Highlight", "-s", tmp])
                    tmp = result[0]    
            elif tmp != "":
                subprocess.run(['python', 'pdf_highlighter.py', '-i', os.path.join(pdfs_path, pdf_file), "-a", "Highlight", "-s", tmp])
                tmp = ""                
                
            
            
#             if "##" in result[0]:
#                 result[0] = result[0].replace("##", "")
#             if result[1] != "O":
#                 subprocess.run(['python', 'pdf_highlighter.py', '-i', os.path.join(pdfs_path, pdf_file), "-a", "Highlight", "-s", result[0]])
        
if __name__ == "__main__":
    getData()
    evaluate()
    deidentification()