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

    pdf_file = f'{PROJECT_ROOT}/MOCK_DATA.test.pdf'
    if not os.path.isfile(pdf_file): raise Exception("Can not convert")
    return pdf_file


def libreoffice_exec():
    if sys.platform == 'darwin':
        return '/Applications/LibreOffice.app/Contents/MacOS/soffice'
    return 'libreoffice'



if __name__ == '__main__':
    pdf_file = excel_to_pdf(xlsx_file=f'{PROJECT_ROOT}/MOCK_DATA.test.xlsx')
    print(f'Converted to {pdf_file}')
