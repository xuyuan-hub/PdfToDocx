from pdf2docx import Converter

def pdf_to_word(pdf_file,docx_file):
    pdf = Converter(pdf_file)
    pdf.convert(docx_file,start=0,end=None)
    pdf.close()

if __name__ == "__main__":
    pdf_file = 'traningVIT.pdf'
    word_file = 'traningVIT.docx'
    pdf_to_word(pdf_file,word_file)