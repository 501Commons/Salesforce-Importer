#!c:\python\python36-32\python.exe
# EASY-INSTALL-ENTRY-SCRIPT: 'python-quickbooks==0.7.4','console_scripts','quickbooks-cli'
__requires__ = 'python-quickbooks==0.7.4'
import re
import sys
from pkg_resources import load_entry_point

if __name__ == '__main__':
    sys.argv[0] = re.sub(r'(-script\.pyw?|\.exe)?$', '', sys.argv[0])
    sys.exit(
        load_entry_point('python-quickbooks==0.7.4', 'console_scripts', 'quickbooks-cli')()
    )
