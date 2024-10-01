#!D:\OJD\venv\Scripts\python.exe
# EASY-INSTALL-ENTRY-SCRIPT: 'canoe==0.1.1','console_scripts','canoe'
__requires__ = 'canoe==0.1.1'
import re
import sys
from pkg_resources import load_entry_point

if __name__ == '__main__':
    sys.argv[0] = re.sub(r'(-script\.pyw?|\.exe)?$', '', sys.argv[0])
    sys.exit(
        load_entry_point('canoe==0.1.1', 'console_scripts', 'canoe')()
    )
