
"""
@author: vitopiepoli
"""

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QHeaderView, QTableWidgetItem, QProgressDialog
from PyQt5.QtWidgets import QMessageBox, QProgressBar, QHeaderView, QApplication, QPushButton
from PyQt5.QtGui import QCloseEvent
import os
import os.path
from os.path import join
import glob
import os
from PyPDF2 import PdfFileReader, PdfFileWriter
import pdftotext
from whoosh import index
from whoosh.fields import Schema, TEXT, ID, STORED
from whoosh.analysis import RegexTokenizer
from whoosh.analysis import StopFilter
from whoosh import scoring
from whoosh.index import open_dir
from whoosh import qparser
from whoosh.qparser import QueryParser, AndGroup
from whoosh import highlight
import pandas as pd
from PIL import Image
import pytesseract
import sys
from pdf2image import convert_from_path
import win32com.client as client
import cv2
import pandas as pd
import numpy as np
from fpdf import FPDF
import xlsxwriter
import concurrent.futures
pytesseract.pytesseract.tesseract_cmd = r'C:/Program Files/Tesseract-OCR/tesseract.exe'

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1200, 1000)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(40, 30, 100, 30))
        self.pushButton.setObjectName("pushButton")
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(180, 30, 80, 30))
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_3 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_3.setGeometry(QtCore.QRect(720, 60, 80, 30))
        self.pushButton_3.setObjectName("pushButton_3")
        self.pushButton_4 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_4.setGeometry(QtCore.QRect(180, 60, 80, 30))
        self.pushButton_4.setObjectName("pushButton_4")
        self.pushButton_5 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_5.setGeometry(QtCore.QRect(340, 60, 80, 30))
        self.pushButton_5.setObjectName("pushButton_5")
        self.pushButton_6 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_6.setGeometry(QtCore.QRect(260, 60, 80, 30))
        self.pushButton_6.setObjectName("pushButton_6")
        self.lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit.setGeometry(QtCore.QRect(480, 60, 200, 30))
        self.lineEdit.setObjectName("lineEdit")
        self.lineEdit_2 = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_2.setGeometry(QtCore.QRect(40, 90, 50, 21))
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(480, 30, 50, 35))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.label2 = QtWidgets.QLabel(self.centralwidget)
        self.label2.setGeometry(QtCore.QRect(40, 70, 150, 16))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label2.setFont(font)
        self.label2.setObjectName("label")
        self.tableWidget = QtWidgets.QTableWidget(self.centralwidget)
        self.tableWidget.setGeometry(QtCore.QRect(0, 120, 1121, 800))
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setHorizontalScrollBarPolicy(
            QtCore.Qt.ScrollBarAsNeeded)
        self.tableWidget.setSizeAdjustPolicy(
            QtWidgets.QAbstractScrollArea.AdjustToContents)
        self.tableWidget.resizeColumnsToContents()
        self.tableWidget.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeToContents)
        self.tableWidget.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1126, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.exitbtn = QPushButton("Exit", self.centralwidget)
        self.exitbtn.resize(self.exitbtn.sizeHint())
        self.exitbtn.move(1027, 80)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

        self.pushButton.clicked.connect(self.open_directory)
        self.pushButton_2.clicked.connect(self.createindex)
        self.pushButton_3.clicked.connect(self.export)
        self.pushButton_4.clicked.connect(self.OCR)
        self.pushButton_5.clicked.connect(self.TXT)
        self.pushButton_6.clicked.connect(self.DOC)
        self.lineEdit.returnPressed.connect(self.search)
        self.lineEdit.returnPressed.connect(self.datatable)
        self.exitbtn.clicked.connect(self.aquit)

    def aquit(self):
        MainWindow.close()

    def datatable(self):
        try:
            numrows = len(self.data)
            numcols = len(self.data[0])
            self.tableWidget.setColumnCount(numcols)
            self.tableWidget.setRowCount(numrows)
            self.tableWidget.setHorizontalHeaderLabels(
                (list(self.data[0].keys())))
            for row in range(numrows):
                for column in range(numcols):
                    item = (list(self.data[row].values())[column])
                    self.tableWidget.setItem(
                        row, column, QTableWidgetItem(item))
        except:
            self.clickMethod2()

    def open_directory(self):
        self.dialog = QtWidgets.QFileDialog()
        self.folder_path = self.dialog.getExistingDirectory(
            None, "Select Folder")
        return self.folder_path

    def TXT(self):
        os.chdir(self.folder_path)
        files = glob.glob("*.txt")
        for file in files:
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=11)
            fd = open(file, "r")
            for i in fd:
                pdf.cell(0, 10, txt=i, ln=1, align="L")
            pdf.output(os.path.splitext(os.path.basename(file))[0] + ".pdf")

        self.createindex()

    def IndexBuilding(self):
        os.chdir(self.folder_path)
        self.files = glob.glob('./Splitted/Txt/*.txt')
        MYDIR = ("indexdir")
        CHECK_FOLDER = os.path.isdir(MYDIR)
        if not CHECK_FOLDER:
            os.makedirs(MYDIR)
        if index.exists_in("indexdir"):
            self.my_analyzer = RegexTokenizer() | StopFilter(lang="en")
            self.schema = Schema(title=TEXT(stored=True), path=ID(stored=True),
                                 content=TEXT(analyzer=self.my_analyzer),
                                 textdata=TEXT(stored=True, phrase = True))
        # set an index writer to add document as per schema
            self.ix = index.open_dir("indexdir")
            self.writer = self.ix.writer()

            self.filepaths = [i for i in self.files]
            progressbar = ProgressBar(
                len(self.filepaths), title="Building index...")
            for i in range(len(self.filepaths)):
                progressbar.setValue(i)
                if progressbar.wasCanceled():
                    break
            for path in self.filepaths:
                self.fp = open(path, "r", encoding='utf-8', errors='ignore')
                self.text = self.fp.read()
                self.writer.add_document(title=os.path.splitext(os.path.basename(path))[
                                         0], path=path, content=self.text, textdata=self.text)
                self.fp.close()
            self.writer.commit()

            self.clickMethod()
        else:
            os.chdir(self.folder_path)
            self.files = glob.glob('./Splitted/Txt/*.txt')
            self.my_analyzer = RegexTokenizer() | StopFilter(lang="en")
            self.schema = Schema(title=TEXT(stored=True), path=ID(stored=True),
                                 content=TEXT(analyzer=self.my_analyzer),
                                 textdata=TEXT(stored=True))
            self.ix = index.create_in("indexdir", self.schema)
            self.writer = self.ix.writer()

            self.filepaths = [i for i in self.files]
            progressbar = ProgressBar(
                len(self.filepaths), title="Building index...")
            for i in range(len(self.filepaths)):
                progressbar.setValue(i)
                if progressbar.wasCanceled():
                    break
            for path in self.filepaths:
                self.fp = open(path, "r", encoding='utf-8', errors='ignore')
                self.text = self.fp.read()
                self.writer.add_document(title=os.path.splitext(os.path.basename(path))[
                                         0], path=path, content=self.text, textdata=self.text)
                self.fp.close()
            self.writer.commit()
            self.clickMethod()

    def DOC(self):
        os.chdir(self.folder_path)
        docfiles = []
        for ext in ('*.doc', '*.rtf', "*.docx"):
            docfiles.extend(glob.glob(join("", ext)))

        try:
            word = client.DispatchEx("Word.Application")
            for file in docfiles:
                if file.endswith('.doc') or file.endswith('.docx') or file.endswith('.rtf'):
                    filepath = (os.getcwd() + "\\" + file)
                    target_path = filepath.replace(
                        os.path.splitext(file)[1], ".pdf")
                    word_doc = word.Documents.Open(filepath)
                    word_doc.SaveAs(target_path, FileFormat=17)
                    word_doc.Close()
                    word.Quit()
        except Exception as e:
            raise e
        
            
        docfiles2 = []
        for ext in ('*.doc', '*.rtf', "*.docx"):
            docfiles2.extend(glob.glob(join("", ext)))
        for f in docfiles2:
            os.remove(f)
        self.createindex()

    def OCR(self):

        # Setting Directories
        if not os.path.isdir(self.folder_path + "/Splitted/Txt"):
            os.chdir(self.folder_path)
            MYDIR = ("Splitted")
            CHECK_FOLDER = os.path.isdir(MYDIR)
            if not CHECK_FOLDER:
                os.makedirs(MYDIR)
            os.chdir(self.folder_path + "/Splitted")
            MYDIR = ("Txt")
            CHECK_FOLDER = os.path.isdir(MYDIR)
            if not CHECK_FOLDER:
                os.makedirs(MYDIR)

            self.toJpg()

        else:
            os.chdir(self.folder_path)
            files = glob.glob('./Splitted/*.jpg')
            for jpg in files:
                os.remove(jpg)

            os.chdir(self.folder_path)
            files = glob.glob('./Splitted/*.pdf')
            for pdf in files:
                os.remove(pdf)

            files = glob.glob('./Splitted/Txt/*.txt')
            for f in files:
                os.remove(f)

            os.chdir(self.folder_path)
            MYDIR = ("Splitted")
            CHECK_FOLDER = os.path.isdir(MYDIR)
            if not CHECK_FOLDER:
                os.makedirs(MYDIR)
            os.chdir(self.folder_path + "/Splitted")
            MYDIR = ("Txt")
            CHECK_FOLDER = os.path.isdir(MYDIR)
            if not CHECK_FOLDER:
                os.makedirs(MYDIR)

            self.toJpg()

    def toJpg(self):
        Image.MAX_IMAGE_PIXELS = None
        os.chdir(self.folder_path)
        self.myPDfs = glob.glob("*.pdf")
        allpages = []
        for PDF_file in self.myPDfs:
            self.pages = convert_from_path(PDF_file, dpi=200, fmt='jpeg')
            allpages.append({"File": PDF_file, "Page": self.pages})
            aaa = pd.DataFrame.from_dict(allpages)
            aab = pd.concat(
                [aaa["File"], aaa["Page"].apply(pd.Series)], axis=1)
            aab = aab.fillna(0)
            apg = aab.melt(id_vars="File")
            apg = apg[apg.value != 0]
            for p, t, n in zip(apg["File"], apg["value"], apg["variable"]):
                filename = os.path.splitext(os.path.basename(p))[
                    0] + " page " + str(n+1)+".jpg"
                t.save(os.path.join("./Splitted", filename), 'JPEG')

        self.tess()

    def tess(self):
        myfiles = glob.glob("./Splitted/*.jpg")
        for file in myfiles:
            outfile = os.path.join(
                "./Splitted/Txt", os.path.splitext(os.path.basename(file))[0] + ".txt")
            f = open(outfile, "w")
            text = str(((pytesseract.image_to_string(Image.open(file)))))
            text = text.replace('-\n', '')
            f.write(text)
            f.close()

        self.IndexBuilding()

    def createindex(self):
        if not os.path.isdir(self.folder_path + "/Splitted/Txt"):
            os.chdir(self.folder_path)
            self.mypdfiles = glob.glob("*.pdf")
            MYDIR = ("Splitted")
            CHECK_FOLDER = os.path.isdir(MYDIR)
            if not CHECK_FOLDER:
                os.makedirs(MYDIR)

        # save split downloaded file and save into new folder
            for self.file in self.mypdfiles:
                progressbar = ProgressBar(
                    len(self.mypdfiles), title="Splitting files...")
                for i in range(len(self.mypdfiles)):
                    progressbar.setValue(i)
                    if progressbar.wasCanceled():
                        break
                self.fname = os.path.splitext(os.path.basename(self.file))[0]
                self.pdf = PdfFileReader(self.file)
                for self.page in range(self.pdf.getNumPages()):
                    self.pdfwrite = PdfFileWriter()
                    self.pdfwrite.addPage(self.pdf.getPage(self.page))
                    self.outputfilename = '{}_page_{}.pdf'.format(
                        self.fname, self.page+1)
                    with open(os.path.join("./Splitted", self.outputfilename), 'wb') as out:
                        self.pdfwrite.write(out)
                        print('Created: {}'.format(self.outputfilename))

            os.chdir(self.folder_path + "/Splitted")
            self.spltittedfiles = glob.glob("*.pdf")
            MYDIR = ("Txt")
            CHECK_FOLDER = os.path.isdir(MYDIR)
            if not CHECK_FOLDER:
                os.makedirs(MYDIR)

            for self.file in self.spltittedfiles:
                progressbar = ProgressBar(
                    len(self.spltittedfiles), title="Writing files...")
                for i in range(len(self.spltittedfiles)):
                    progressbar.setValue(i)
                    if progressbar.wasCanceled():
                        break
                with open(self.file, "rb") as f:
                    self.pdf = pdftotext.PDF(f)
                    with open(os.path.join("./TXT", os.path.splitext(os.path.basename(self.file))[0] + ".txt"), 'w', encoding='utf-8') as f:
                        f.write("\n\n".join(self.pdf))
                    f.close()
        else:
            os.chdir(self.folder_path)
            files = glob.glob('./Splitted/*.jpg')
            for jpg in files:
                os.remove(jpg)

            os.chdir(self.folder_path)
            files = glob.glob('./Splitted/*.pdf')
            for pdf in files:
                os.remove(pdf)

            files = glob.glob('./Splitted/Txt/*.txt')
            for f in files:
                os.remove(f)

            os.chdir(self.folder_path)
            self.mypdfiles = glob.glob("*.pdf")
            progressbar = ProgressBar(
                len(self.mypdfiles), title="Splitting files...")
            for i in range(len(self.mypdfiles)):
                progressbar.setValue(i)
                if progressbar.wasCanceled():
                    break
        # save split downloaded file and save into new folder
            for self.file in self.mypdfiles:
                self.fname = os.path.splitext(os.path.basename(self.file))[0]
                self.pdf = PdfFileReader(self.file)
                for self.page in range(self.pdf.getNumPages()):
                    self.pdfwrite = PdfFileWriter()
                    self.pdfwrite.addPage(self.pdf.getPage(self.page))
                    self.outputfilename = '{}_page_{}.pdf'.format(
                        self.fname, self.page+1)
                    with open(os.path.join("./Splitted", self.outputfilename), 'wb') as out:
                        self.pdfwrite.write(out)
                        print('Created: {}'.format(self.outputfilename))

            os.chdir(self.folder_path + "/Splitted")
            self.spltittedfiles = glob.glob("*.pdf")
            progressbar = ProgressBar(
                len(self.spltittedfiles), title="Writing files...")
            for i in range(len(self.spltittedfiles)):
                progressbar.setValue(i)
                if progressbar.wasCanceled():
                    break
            for self.file in self.spltittedfiles:
                with open(self.file, "rb") as f:
                    self.pdf = pdftotext.PDF(f)
                    with open(os.path.join("./TXT", os.path.splitext(os.path.basename(self.file))[0] + ".txt"), 'w', encoding='utf-8') as f:
                        f.write("\n\n".join(self.pdf))
                    f.close()
        self.IndexBuilding()

    def clickMethod(self):
        QMessageBox.information(
            None, "Index Completed", "Now you start searching.", QtWidgets.QMessageBox.Ok)

    def clickMethod2(self):
        QMessageBox.information(
            None, "Search Alert", "No such information found", QtWidgets.QMessageBox.Ok)

    def search(self):

        os.chdir(self.folder_path)
        self.ix = open_dir("indexdir")
        MYDIR = ("Results")
        CHECK_FOLDER = os.path.isdir(MYDIR)
        if not CHECK_FOLDER:
            os.makedirs(MYDIR)
        self.text = self.lineEdit.text()
        self.query_str = self.text
        self.query = qparser.QueryParser(
            "content", schema=self.ix.schema)
        self.q = self.query.parse(self.query_str)
        self.topN = self.lineEdit_2.text()
        if self.lineEdit_2.text() == "":
            self.topN = 1000
        else:
            self.topN = int(self.lineEdit_2.text())

        self.data = []
        with self.ix.searcher() as searcher:
            self.results = searcher.search(self.q, terms=True, limit=self.topN)
            Upper = highlight.UppercaseFormatter()
            self.results.formatter = Upper
            my_cf = highlight.ContextFragmenter(maxchars=500, surround=300)
            self.results.fragmenter = my_cf
            for self.i in self.results:
                self.data.append({"Title": self.i['title'], "Text": self.i.highlights("content", text = self.i["textdata"]), "Score": str(round((self.i.score), 3))})
        pd.DataFrame(self.data).to_excel(
            self.text.strip("\"") + ".xlsx", engine="xlsxwriter")

    def export(self):
        with self.ix.searcher() as searcher:
            self.results = searcher.search(self.q, terms=True, limit=None)
            Upper = highlight.UppercaseFormatter()
            self.results.formatter = Upper
            my_cf = highlight.ContextFragmenter(maxchars=500, surround=300)
            self.results.fragmenter = my_cf
            self.countrow = len(self.results)
            for self.i in self.results:
                with open(os.path.join(self.folder_path, self.text.strip("\"") + ".txt"), 'a', encoding="utf-8") as f:
                    print("Title {}".format(self.i['title']), "Text {}".format(
                        self.i['textdata']), file=f)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Search Text"))
        self.pushButton.setText(_translate("MainWindow", "Select Folder"))
        self.pushButton_2.setText(_translate("MainWindow", "PDF"))
        self.pushButton_3.setText(_translate("MainWindow", "Export"))
        self.pushButton_4.setText(_translate("MainWindow", "OCR"))
        self.pushButton_5.setText(_translate("MainWindow", "TXT"))
        self.pushButton_6.setText(_translate("MainWindow", "DOC"))
        self.label.setText(_translate("MainWindow", "Search"))
        self.label2.setText(_translate("MainWindow", "Top Results"))


class ProgressBar(QProgressDialog):
    def __init__(self, max, title):
        super().__init__()
        # Sets how long the loop should last before progress bar is shown (in miliseconds)
        self.setMinimumDuration(1000)
        self.setWindowTitle(title)
        self.setModal(True)
        self.setValue(0)
        self.setMinimum(1000)
        self.setMaximum(max)
        self.resize(280, 100)

        self.show()


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())


input()
