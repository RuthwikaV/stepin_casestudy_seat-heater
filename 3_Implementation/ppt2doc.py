import comtypes.client
from pdf2docx import Converter
import os


def PPTtoPDF():
    givenFileName = "E:\pyconv\ppt_doc\MOUNTAINS.pptx"
    gotFileName = "E:\pyconv\pdf_doc\MOUNTAINS.pdf"
    formatType = 32
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = False

    if gotFileName[-3:] != 'pdf':
       gotFileName = gotFileName + ".pdf"
    deck = powerpoint.Presentations.Open(givenFileName)
    deck.SaveAs(gotFileName, formatType) # formatType = 32 for ppt to pdf
    deck.Close()
    powerpoint.Quit()
    return 0


def conv():
    pdffile = "E:\pyconv\pdf_doc\MOUNTAINS.pdf"
    wordfile = "E:\pyconv\ppt_doc\MOUNTAINS.doc"
    cnv = Converter(pdf_file)
    cnv.convert(word_file, start=0, end=None)
    cnv.close()
    os.remove("E:\pyconv\pdf_doc\MOUNTAINS.pdf")
    return 0

# ppt2doc.PPTtoPDF()
# ppt2doc.conv()
