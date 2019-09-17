import sys
import subprocess
import os


PROJECT_ROOT = os.path.abspath(os.path.dirname(__file__))

template_path = f'{PROJECT_ROOT}/MOCK_DATA.test.xlsx'
path = f'{PROJECT_ROOT}/'


def convert_to(timeout=None):
    args = [libreoffice_exec(), '--headless', '--convert-to', 'pdf', '--outdir', path, template_path]

    process = subprocess.run(args, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=timeout)
    filename = os.path.isfile(f'{path}MOCK_DATA.pdf')
    if not filename:
        raise Exception("Can not convert")
    else:
        return 'MOCK_DATA.pdf'


def libreoffice_exec():
    if sys.platform == 'darwin':
        return '/Applications/LibreOffice.app/Contents/MacOS/soffice'
    return 'libreoffice'



if __name__ == '__main__':
    print('Converted to ' + convert_to())
