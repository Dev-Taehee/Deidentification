import os
from pdf2image import convert_from_path
from img2pdf import convert

pdfs_path = os.path.join(os.getcwd(), "pdfs")
pdfs_files = [f for f in os.listdir(pdfs_path)]

for pdf_file in pdfs_files:
    pages = convert_from_path(os.path.join(pdfs_path, pdf_file))
    for i, page in enumerate(pages):
        page.save(os.path.join(pdfs_path, pdf_file.split(".")[0])+".jpg", "JPEG")
        
    with open(os.path.join(pdfs_path, pdf_file.split(".")[0])+"_JPG"+".pdf", "wb") as f:
        image_list = []
        image_list.append(os.path.join(pdfs_path, pdf_file.split(".")[0])+".jpg")
        pdf = convert(image_list)
        f.write(pdf)