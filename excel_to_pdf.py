import sys
import subprocess
import os


PROJECT_ROOT  = os.path.abspath(os.path.dirname(__file__))
template_path = f'{PROJECT_ROOT}/MOCK_DATA.test.xlsx'


def convert_to(timeout=None) -> 'pdf_file path':
    # run libreoffice --headless --convert-to pdf --outdir :PROJECT_ROOT :xlsx_file
    args = [libreoffice_exec(), '--headless', '--convert-to', 'pdf', '--outdir', PROJECT_ROOT, template_path]
    _ = subprocess.run(args, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=timeout)

    pdf_file = f'{PROJECT_ROOT}/MOCK_DATA.pdf'
    if not os.path.isfile(pdf_file):
        raise Exception("Can not convert")

    return pdf_file


def libreoffice_exec():
    if sys.platform == 'darwin':
        return '/Applications/LibreOffice.app/Contents/MacOS/soffice'
    return 'libreoffice'



if __name__ == '__main__':
    print('Converted to ' + convert_to())
