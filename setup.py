from distutils.core import setup
import py2exe
import bs4
from win32process import *

bs4.__version__
setup(windows=[{'script':'AwesomeScraper.py', 'uac_info': "requireAdministrator"}])