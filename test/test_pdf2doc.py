from docx import Document
from pdf2docx import Converter
import pandas as pd

def pdf_to_word(pdf_path, word_path):
    """将PDF转换为Word文档"""
    cv = Converter(pdf_path)
    cv.convert(word_path, start=0, end=None)
    tables = cv.extract_tables()
    cv.close()
    return tables

if __name__ == '__main__':
    tables = pdf_to_word('/Users/jaron/Desktop/Python_Programs/ZHANZHI/PDF超链接提取/SGFP Doc/Acquisition Cloud.PDF',
                      '/Users/jaron/Desktop/Python_Programs/ZHANZHI/PDFTranslate/test/test.docx')
    for index,_ in enumerate(tables):
        print(f"""{index}:
              ------------------
              {pd.DataFrame(_)}
              ------------------""")