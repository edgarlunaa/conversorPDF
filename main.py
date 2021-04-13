from PyPDF2 import PdfFileWriter, PdfFileReader
import Ui_main
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QMessageBox
import sys
import os
from openpyxl import load_workbook

class MainUiClass (QtWidgets.QMainWindow, Ui_main.Ui_MainWindow):
    def __init__(self, parent=None):
        super (MainUiClass, self).__init__(parent)
        self.setupUi(self)

def extract_page(doc_name, page_num_desde, page_num_hasta, intervalo, nombre_carpeta, nombres_desde_excel):
    os.system('mkdir \"' + nombre_carpeta + '\"')
    pdf_reader = PdfFileReader(open(doc_name, 'rb'))
    index_nombre = -1
    pagina_actual = page_num_desde
    count = 0
    for i in range(page_num_desde, page_num_hasta, intervalo):
        pdf_writer = PdfFileWriter()

        for j in range(intervalo):
            try:
                pdf_writer.addPage(pdf_reader.getPage(pagina_actual))
                pagina_actual += 1
            except:
                return count
        index_nombre += 1
        if nombres_desde_excel:
            try:
                with open(nombre_carpeta + f'\\{lista_nombres[index_nombre]}.pdf', 'wb') as doc_file:
                    pdf_writer.write(doc_file)
                    count += 1
            except:
                msg = QMessageBox()
                msg.setText("Error, la cantidad de archivos no coincide con la cantidad de nombres dados")
                msg.setIcon(QMessageBox.Warning)
                msg.setWindowTitle("Error")
                msg.exec_()
        else:
            if intervalo == 1:
                with open(nombre_carpeta + f'\\Página {pagina_actual}.pdf', 'wb') as doc_file:
                    pdf_writer.write(doc_file)
                    count += 1
            else:
                with open(nombre_carpeta + f'\\Página desde {pagina_actual - intervalo + 1} hasta {pagina_actual}.pdf', 'wb') as doc_file:
                    pdf_writer.write(doc_file)
                    count += 1

    return count
class Aplicacion (object):
    def __init__(self, ui):
        self.ui = ui

    def btn_arriba_click(self):
        seleccionado = self.ui.listWidget.currentRow()
        if seleccionado != 0:
            self.lista_pdfs[seleccionado], self.lista_pdfs[seleccionado - 1] = self.lista_pdfs[seleccionado - 1], self.lista_pdfs[seleccionado]
            self.actualizar_lista()
            self.ui.listWidget.setCurrentRow(seleccionado - 1)

    def btn_abajo_click(self):
        seleccionado = self.ui.listWidget.currentRow()
        if seleccionado != self.ui.listWidget.count() - 1:
            self.lista_pdfs[seleccionado], self.lista_pdfs[seleccionado + 1] = self.lista_pdfs[seleccionado + 1], self.lista_pdfs[seleccionado]
            self.actualizar_lista()
            self.ui.listWidget.setCurrentRow(seleccionado + 1)

    def btn_excel_click(self):
        self.ui.lbl_excel.setText("")
        filename = QtWidgets.QFileDialog.getOpenFileName(filter = "Excel (*.xlsx)")
        if filename[0]:
            self.ui.lbl_excel.setText(filename[0])
            wb = load_workbook(filename[0])
            ws = wb.active

            global lista_nombres
            lista_nombres = []

            for row in ws.values:
                if row[0]:
                    lista_nombres.append(str(row[0]))
            
            self.ui.lbl_cantidad_nombres.setText(str(len(lista_nombres)))

            if int(self.ui.lbl_cantidad_nombres.text()) < int(self.ui.lbl_cantidad_archivos.text()):
                self.ui.btn_iniciar.setEnabled(False)
            elif int(self.ui.lbl_cantidad_nombres.text()) > int(self.ui.lbl_cantidad_archivos.text()):
                self.ui.lbl_warnings.setText("⚠¡Atención la cantidad de nombres\ningresados por excel excede a la \ncantidad de archivos a generar!⚠")
                self.ui.btn_iniciar.setEnabled(True)

    def btn_pdf_click(self):
        self.ui.lbl_pdf.setText("")
        filename = QtWidgets.QFileDialog.getOpenFileName(filter = "PDF (*.pdf)")
        if filename[0]:
            self.ui.lbl_pdf.setText(filename[0])
            self.ui.btn_iniciar.setEnabled(True)
            pdf_reader = PdfFileReader(open(filename[0], 'rb'))
            cant_paginas = pdf_reader.getNumPages()
            self.ui.lbl_cant_pag.setText(str(cant_paginas))
            self.ui.cant_pag_a_recortar.setValue(cant_paginas)
            self.ui.cant_pag_a_recortar.setEnabled(True)
            self.ui.cant_a_partir.setEnabled(True)
            self.ui.cant_intervalo.setEnabled(True)
            self.ui.txt_carpeta.setEnabled(True)
            self.ui.chkBox_excel.setEnabled(True)
        else:
            self.ui.lbl_pdf.setText("Error al cargar pdf")
            self.ui.btn_iniciar.setEnabled(False)
    
    def btn_iniciar_click(self):
        if self.ui.lbl_pdf.text() == "Ubicación del pdf":
            msg_cargar_pdf = QMessageBox()
            msg_cargar_pdf.setIcon(QMessageBox.Warning)
            msg_cargar_pdf.setText("¡Cargue un pdf primero!.")
            msg_cargar_pdf.setWindowTitle("Error")
            msg_cargar_pdf.exec_()
            return
        if self.ui.chkBox_excel.isChecked() and self.ui.lbl_excel.text() == "Ubicación del excel":
            msg_cargar_excel = QMessageBox()
            msg_cargar_excel.setIcon(QMessageBox.Warning)
            msg_cargar_excel.setText("Error, debe cargar un Excel o desmarcar la opción.")
            msg_cargar_excel.setWindowTitle("Error")
            msg_cargar_excel.exec_()
            return
        if self.ui.cant_a_partir.value() > int(self.ui.lbl_cant_pag.text()):
            msg_a_partir_ultima_pag = QMessageBox()
            msg_a_partir_ultima_pag.setIcon(QMessageBox.Warning)
            msg_a_partir_ultima_pag.setText("Error, no se puede a partir de más de la última página")
            msg_a_partir_ultima_pag.setWindowTitle("Error")
            msg_a_partir_ultima_pag.exec_()
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
        pagina_hasta = self.ui.cant_pag_a_recortar.value() + pagina
        intervalo = self.ui.cant_intervalo.value()
        nombre_carpeta = self.ui.txt_carpeta.text()
        bool_excel = self.ui.chkBox_excel.isChecked()

        paginas_generadas = extract_page(self.ui.lbl_pdf.text(), pagina, pagina_hasta, intervalo, nombre_carpeta, bool_excel)

        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText(f"Se han generado {paginas_generadas} archivos PDF.")
        msg.setWindowTitle("Información")
        msg.exec_()
    
    def chkBox_excel_click(self, state):
        if state == QtCore.Qt.Checked:
            self.ui.btn_excel.setEnabled(True)
            self.ui.lbl_cantidad_nombres.setText("0")
        else:
            self.ui.btn_excel.setEnabled(False)
            self.ui.lbl_cantidad_nombres.setText("")
            self.ui.lbl_excel.setText("Ubicación del excel")

    def lbl_cantidad_archivos_update(self):
        x = self.ui.cant_pag_a_recortar.value() // self.ui.cant_intervalo.value()
        self.ui.lbl_cantidad_archivos.setText(str(x))
        if x == 0:
            self.ui.btn_iniciar.setEnabled(False)
        else:
            if self.ui.lbl_pdf.text() == "Ubicación del pdf":
                return
            if self.ui.cant_a_partir.value() > int(self.ui.lbl_cant_pag.text()):
                self.ui.btn_iniciar.setEnabled(False)
                return
            self.ui.btn_iniciar.setEnabled(True)
        try:
            if int(self.ui.lbl_cantidad_nombres.text()) < int(self.ui.lbl_cantidad_archivos.text()):
                self.ui.btn_iniciar.setEnabled(False)
        except:
            pass
    
    def actualizar_lista(self):
        self.ui.listWidget.clear()
        for pdf in self.lista_pdfs:
            pdf = pdf[::-1]
            busca_barra = pdf.find("/")
            pdf = pdf[:busca_barra]
            pdf = pdf[::-1]
            self.ui.listWidget.addItem(pdf)

    def btn_cargar_pdfs_click(self):
        filter = "PDF (*.pdf)"
        names = QtWidgets.QFileDialog.getOpenFileNames(filter=filter, caption="Abir")
        try:
            if names[0]:

                if len(names[0]) == 1:
                    msg = QMessageBox()
                    msg.setIcon(QMessageBox.Warning)
                    msg.setText("Debe seleccionar al menos 2 pdfs")
                    msg.setWindowTitle("Atención")
                    msg.exec_()
                    return

                self.lista_pdfs = names[0]
                self.actualizar_lista()
                self.ui.btn_guardar.setEnabled(True)
                self.ui.btn_arriba.setEnabled(True)
                self.ui.btn_abajo.setEnabled(True)
            else:
                self.ui.btn_guardar.setEnabled(False)
                self.ui.btn_unir.setEnabled(False)
                self.ui.txt_nombre_archivo.setEnabled(False)
                self.ui.txt_nombre_archivo.setText("")
                self.ui.lbl_guardar_en.setText("")
                self.ui.lbl_pdfs.setText("")

        except TypeError:
            pass

    def btn_guardar_en_click(self):
        
        directory = QtWidgets.QFileDialog.getExistingDirectory(None, 'Seleccionar carpeta:', 'C:\\', QtWidgets.QFileDialog.ShowDirsOnly)

        try:
            if directory:
                self.ui.lbl_guardar_en.setText(directory)
                self.ui.btn_unir.setEnabled(True)
                self.ui.txt_nombre_archivo.setEnabled(True)
            else:
                self.ui.btn_unir.setEnabled(False)
                self.ui.txt_nombre_archivo.setEnabled(False)
                self.ui.txt_nombre_archivo.setText("")
                self.ui.lbl_guardar_en.setText("")
        except TypeError:
            pass
    
    def btn_unir_click(self):
        if self.ui.txt_nombre_archivo.text() != "":
            self.ui.lbl_informacion.setText("Cargando...")
            pdf_writer = PdfFileWriter()
            pagina = 0
            for pdf in self.lista_pdfs:
                pdf_reader = PdfFileReader(open(pdf, 'rb'))
                cant_paginas = pdf_reader.getNumPages()

                for i in range(cant_paginas):
                    pdf_writer.addPage(pdf_reader.getPage(i))
            
            with open(self.ui.lbl_guardar_en.text() + '\\' + self.ui.txt_nombre_archivo.text() + '.pdf', 'wb') as doc_file:
                pdf_writer.write(doc_file)

            self.ui.lbl_informacion.setText("")
            self.ui.listWidget.clear()
            self.ui.lbl_guardar_en.setText("")
            self.ui.txt_nombre_archivo.setText("")
            self.ui.btn_guardar.setEnabled(False)
            self.ui.btn_arriba.setEnabled(False)
            self.ui.btn_abajo.setEnabled(False)
            self.ui.btn_unir.setEnabled(False)
            self.lista_pdfs = []
            self.ui.txt_nombre_archivo.setEnabled(False)

            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setText("Proceso finalizado con éxito.")
            msg.setWindowTitle("Información")
            msg.exec_()

        else:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Warning)
            msg.setText("Error: Debe escribir un nombre para el PDF.")
            msg.setWindowTitle("Error")
            msg.exec_()

def main():
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)
        ui = MainUiClass()
        apli = Aplicacion(ui)
        ui.show()
        ui.btn_pdf.clicked.connect(apli.btn_pdf_click)
        ui.btn_iniciar.clicked.connect(apli.btn_iniciar_click)
        ui.btn_excel.clicked.connect(apli.btn_excel_click)
        ui.btn_iniciar.setEnabled(False)
        ui.btn_excel.setEnabled(False)
        ui.cant_pag_a_recortar.setEnabled(False)
        ui.cant_a_partir.setEnabled(False)
        ui.cant_intervalo.setEnabled(False)
        ui.txt_carpeta.setEnabled(False)
        ui.chkBox_excel.setEnabled(False)
        ui.btn_guardar.setEnabled(False)
        ui.btn_unir.setEnabled(False)
        ui.txt_nombre_archivo.setEnabled(False)
        ui.btn_unir.clicked.connect(apli.btn_unir_click)
        ui.cant_a_partir.valueChanged.connect(apli.lbl_cantidad_archivos_update)
        ui.cant_pag_a_recortar.valueChanged.connect(apli.lbl_cantidad_archivos_update)
        ui.cant_intervalo.valueChanged.connect(apli.lbl_cantidad_archivos_update)
        ui.chkBox_excel.stateChanged.connect(apli.chkBox_excel_click)
        ui.btn_cargar_pdfs.clicked.connect(apli.btn_cargar_pdfs_click)
        ui.btn_guardar.clicked.connect(apli.btn_guardar_en_click)
        ui.btn_arriba.clicked.connect(apli.btn_arriba_click)
        ui.btn_abajo.clicked.connect(apli.btn_abajo_click)
        ui.btn_arriba.setEnabled(False)
        ui.btn_abajo.setEnabled(False)
        app.exec_()



if __name__ == "__main__":
    main()