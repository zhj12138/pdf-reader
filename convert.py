import win32com.client as win32
import os
import fitz
import PySimpleGUI as sg
import time
import pdfkit
import markdown


# 测试成功
def docxToPdf(file_path: str, to_name: str):
    word = win32.gencache.EnsureDispatch('Word.Application')
    pdfForm = 17
    if os.path.exists(file_path):
        if file_path.endswith('.docx'):
            doc = word.Documents.Open(file_path)
            doc.SaveAs(to_name, FileFormat=pdfForm)
            doc.Close()


def picsToPdf(filenames: list, to_name: str):
    doc = fitz.open()
    total = len(filenames)
    count = 0
    for filename in filenames:
        count += 1
        if os.path.exists(filename) and filename.endswith(('.png', '.jpg', 'jpeg')):
            img = fitz.open(filename)
            rect = img[0].rect
            pdfbytes = img.convertToPDF()
            img.close()
            imgPdf = fitz.open("pdf", pdfbytes)
            page = doc.newPage(width=rect.width, height=rect.height)
            page.showPDFpage(rect, imgPdf, 0)
            sg.OneLineProgressMeter("convert images", count, total)
    doc.save(to_name)


def htmlToPdf(url, to_name):
    try:
        pdfkit.from_url(url, to_name)
    except:
        return False
    else:
        return True


def mdToPdf(file_path, to_name):
    extensions = [
        'markdown.extensions.extra',
        'markdown.extensions.codehilite',
        'markdown.extensions.legacy_em',
        'markdown.extensions.toc',
        'markdown.extensions.wikilinks',
        'markdown.extensions.admonition',
        'markdown.extensions.legacy_attrs',
        'markdown.extensions.meta',
        'markdown.extensions.nl2br',
        'markdown.extensions.sane_lists',
        'markdown.extensions.smarty',
    ]
    with open(file_path, 'r', encoding='utf-8') as f:
        text = f.read()
    html = markdown.markdown(text, output_format='html', extensions=extensions)
    pdfkit.from_string(html, to_name, options={'encoding': 'utf-8'})





