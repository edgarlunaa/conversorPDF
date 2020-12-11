from PyPDF2 import PdfFileWriter, PdfFileReader
import menu
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication
import sys
import os

class MainUiClass (QtWidgets.QMainWindow, menu.Ui_MainWindow):
    def __init__(self, parent=None):
        super (MainUiClass, self).__init__(parent)
        self.setupUi(self)

def extract_page(doc_name, page_num):
    pdf_reader = PdfFileReader(open(doc_name, 'rb'))
    pdf_writer = PdfFileWriter()
    pdf_writer.addPage(pdf_reader.getPage(page_num))
    with open(f'document-page{page_num + 1}.pdf', 'wb') as doc_file:
        pdf_writer.write(doc_file)

class Aplicacion (object):
    def __init__(self, ui):
        self.ui = ui
    
    def btn_pdf_click(self):
        self.ui.lbl_pdf.setText("")
        filename = QtWidgets.QFileDialog.getOpenFileName(filter = "PDF (*.pdf)")
        if filename[0]:
            self.ui.lbl_pdf.setText(filename[0])
            self.ui.btn_iniciar.setEnabled(True)
            pdf_reader = PdfFileReader(open(filename[0], 'rb'))
            cant_paginas = pdf_reader.getNumPages()
            self.ui.lbl_cant_pag.setText(str(cant_paginas))
        else:
            self.ui.lbl_pdf.setText("Error al cargar pdf")
            self.ui.btn_iniciar.setEnabled(False)
    
    def btn_iniciar_click(self):
        if self.ui.cant_a_partir.value() - 1 + self.ui.cant_pag_a_recortar.value() >= int(self.ui.lbl_cant_pag.text()) + 1:
            return
        for i in range(self.ui.cant_pag_a_recortar.value()):
            pagina = self.ui.cant_a_partir.value() - 1 + i
            extract_page(self.ui.lbl_pdf.text(), pagina)

def main():
    app = QApplication.instance()
    if app is None:
        directorio = os.path.dirname(os.path.realpath(__file__))
        app = QApplication(sys.argv)
        ui = MainUiClass()
        apli = Aplicacion(ui)
        ui.show()
        ui.btn_pdf.clicked.connect(apli.btn_pdf_click)
        ui.btn_iniciar.clicked.connect(apli.btn_iniciar_click)
        ui.btn_iniciar.setEnabled(False)
        app.exec_()



if __name__ == "__main__":
    main()