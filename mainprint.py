from win32com.client import Dispatch
import os

class Pinrt():

    def __init__(self):
        
        self.wordname = "Word.Application"
        self.excelname = "Excel.Application"

    def toPdf(self,filepath):
        
        filename = os.path.split(filepath)[-1]

        pdfname = os.path.splitext(filename)[0]+"pdf"

        return pdfname
        

    def wordPrint(self,path,outpath):

        word = Dispatch(self.wordname)

        pdfname = self.toPdf(path)

        pdffile = os.path.join(outpath,pdfname)

        word.Visible = 0
        word.DisplayAlerts =0

        doc = word.Documents.Open(path)

        doc.SaveAs(pdffile,17)

        doc.Close()

        word.Quit()

    def excelPrint(self,path,outpath):

        excel = Dispatch(self.excelname)

        pdfname = self.toPdf(path)

        pdffile = os.path.join(outpath,pdfname)

        excel.Visible = 0
        excel.DisplayAlerts =0

        doc = excel.Workbooks.Open(path)

        doc.SaveAs(pdffile,17)

        doc.Close()

        excel.Quit()


