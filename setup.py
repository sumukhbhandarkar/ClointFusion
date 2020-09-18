# from distutils.core import setup
import setuptools
from os import path

this_directory = path.abspath(path.dirname(__file__))

with open(path.join(this_directory, "README.md"), "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name='ClointFusion',
    packages=['ClointFusion'],
    version='0.0.14',
    description="Cloint LLC's Python based backend functions for RPA (Automation)",
    long_description=long_description,
    long_description_content_type="text/markdown",
    author='Cloint LLC',
    author_email='automation@cloint.com',
    url='https://github.com/ClointFusion/ClointFusion',
    keywords=['ClointFusion','RPA','Python','Automation'],
      install_requires=[            
          "howdoi","seaborn","texthero","emoji","helium","kaleido", "folium", "zipcodes", "plotly", "PyAutoGUI", "PyGetWindow", "XlsxWriter" ,"PySimpleGUI", "chromedriver-autoinstaller", "gspread", "imutils", "keyboard", "joblib", "opencv-python", "python-imageseach-drov0", "openpyxl", "pandas", "pif", "pytesseract", "scikit-image", "selenium", "xlrd", "clipboard"
      ],
  classifiers=[
    'Development Status :: 3 - Alpha',      
    'Intended Audience :: Developers',      
    'Topic :: Software Development :: Build Tools',
    'License :: OSI Approved :: BSD License',  
    'Natural Language :: English',
  ],
  python_requires='>=3.8',
)