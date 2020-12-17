from PyPDF2 import PdfFileWriter, PdfFileReader
import Ui_main
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QMessageBox
import sys
import os

class MainUiClass (QtWidgets.QMainWindow, Ui_main.Ui_MainWindow):
    def __init__(self, parent=None):
        super (MainUiClass, self).__init__(parent)
        self.setupUi(self)

def extract_page(doc_name, page_num_desde, page_num_hasta, intervalo, nombre_carpeta):
    os.system("mkdir " + nombre_carpeta)
    pdf_reader = PdfFileReader(open(doc_name, 'rb'))
    for i in range(page_num_desde, page_num_hasta, intervalo):
        pdf_writer = PdfFileWriter()

        for j in range(intervalo):
            pdf_writer.addPage(pdf_reader.getPage(page_num_desde + i + j))

        with open(nombre_carpeta + f'\Página de {page_num_desde + i + 1} hasta {page_num_desde + i + intervalo}.pdf', 'wb') as doc_file:
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
        if self.ui.cant_a_partir.value() == 0:
            msg_a_partir_cero = QMessageBox()
            msg_a_partir_cero.setIcon(QMessageBox.Warning)
            msg_a_partir_cero.setText("Error, no se puede a partir de 0 (cero).")
            msg_a_partir_cero.setWindowTitle("Error")
            msg_a_partir_cero.exec_()
            return
        if self.ui.cant_pag_a_recortar.value() == 0:
            msg_cant_pag_a_recortar_cero = QMessageBox()
            msg_cant_pag_a_recortar_cero.setIcon(QMessageBox.Warning)
            msg_cant_pag_a_recortar_cero.setText("Error, ingrese una cantidad de páginas a recortar.")
            msg_cant_pag_a_recortar_cero.setWindowTitle("Error")
            msg_cant_pag_a_recortar_cero.exec_()
            return
        if self.ui.cant_intervalo.value() == 0:
            msg_cant_intervalo_cero = QMessageBox()
            msg_cant_intervalo_cero.setIcon(QMessageBox.Warning)
            msg_cant_intervalo_cero.setText("Error, ingrese un intervalo.")
            msg_cant_intervalo_cero.setWindowTitle("Error")
            msg_cant_intervalo_cero.exec_()
            return
        if self.ui.cant_a_partir.value() > int(self.ui.lbl_cant_pag.text()):
            msg_a_partir_ultima_pag = QMessageBox()
            msg_a_partir_ultima_pag.setIcon(QMessageBox.Warning)
            msg_a_partir_ultima_pag.setText("Error, no se puede a partir de más de la última página")
            msg_a_partir_ultima_pag.setWindowTitle("Error")
            msg_a_partir_ultima.exec_()
            return
        if self.ui.cant_a_partir.value() - 1 + self.ui.cant_pag_a_recortar.value() >= int(self.ui.lbl_cant_pag.text()) + 1:
            msg_a_partir_exedido = QMessageBox()
            msg_a_partir_exedido.setIcon(QMessageBox.Warning)
            msg_a_partir_exedido.setText("Error, se exedió el numero de páginas del documento.")
            msg_a_partir_exedido.setWindowTitle("Error")
            msg_a_partir_exedido.exec_()
            return
        if self.ui.txt_carpeta.text() == "":
            msg_txt_carpeta_vacio = QMessageBox()
            msg_txt_carpeta_vacio.setIcon(QMessageBox.Warning)
            msg_txt_carpeta_vacio.setText("Error, el nombre de la carpeta no debe estar vacío.")
            msg_txt_carpeta_vacio.setWindowTitle("Error")
            msg_txt_carpeta_vacio.exec_()
            return
        if self.ui.cant_intervalo.value() > int(self.ui.lbl_cant_pag.text()):
            msg_intervalo_exedido = QMessageBox()
            msg_intervalo_exedido.setIcon(QMessageBox.Warning)
            msg_intervalo_exedido.setText("Error, el intervalo de páginas exede al del documento.")
            msg_intervalo_exedido.setWindowTitle("Error")
            msg_intervalo_exedido.exec_()
            return

        pagina = self.ui.cant_a_partir.value() - 1
        pagina_hasta = self.ui.cant_pag_a_recortar.value()
        intervalo = self.ui.cant_intervalo.value()
        nombre_carpeta = self.ui.txt_carpeta.text()

        extract_page(self.ui.lbl_pdf.text(), pagina, pagina_hasta, intervalo, nombre_carpeta)

        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText("Proceso terminado con éxito.")
        msg.setWindowTitle("Información")
        msg.exec_()

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