import sys
import subprocess
import os


PROJECT_ROOT  = os.path.abspath(os.path.dirname(__file__))

def excel_to_pdf(xlsx_file:'str path', timeout=None) -> 'pdf_file path':

    #region run libreoffice
    """
    targeted command
    libreoffice --headless --convert-to pdf --outdir :PROJECT_ROOT :xlsx_file
    """
    args = [libreoffice_exec(), '--headless', '--convert-to', 'pdf', '--outdir', PROJECT_ROOT, xlsx_file]
    _ = subprocess.run(args, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=timeout)
    #endregion

    pdf_file = f'{os.path.splitext(xlsx_file)[0]}.pdf'  # filename without extension ref. https://stackoverflow.com/a/678242/248616
    if not os.path.isfile(pdf_file):
        raise Exception(f'File not found {pdf_file} - Can not convert {xlsx_file}')
    return pdf_file


def libreoffice_exec():
    if sys.platform == 'darwin':
        return '/Applications/LibreOffice.app/Contents/MacOS/soffice'
    return 'libreoffice'


def test_excel2pdf():
    xlsx_file = f'{PROJECT_ROOT}/MOCK_DATA.test.xlsx'
    pdf_file  = excel_to_pdf(xlsx_file)
    print(f'Converted {xlsx_file} to {pdf_file}')


def test_html2pdf():
    html_file = f'{PROJECT_ROOT}/test.html'
    pdf_file  = excel_to_pdf(html_file)
    print(f'Converted {html_file} to {pdf_file}')



if __name__ == '__main__':
    test_excel2pdf()
    test_html2pdf()
