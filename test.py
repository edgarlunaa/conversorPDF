from PyPDF2 import PdfFileWriter, PdfFileReader

def extract_page(doc_name, page_num):
    pdf_reader = PdfFileReader(open(doc_name, 'rb'))
    pdf_writer = PdfFileWriter()
    pdf_writer.addPage(pdf_reader.getPage(page_num))
    with open(f'document-page{page_num}.pdf', 'wb') as doc_file:
        pdf_writer.write(doc_file)

extract_page("sample.pdf", 1)