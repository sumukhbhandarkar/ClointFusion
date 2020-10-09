# from distutils.core import setup
import setuptools
from os import path

this_directory = path.abspath(path.dirname(__file__))

with open(path.join(this_directory, "README.md"), "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name='ClointFusion',
    packages=['ClointFusion'],
    version='0.0.20',
    description="Python based functions for RPA (Automation)",
    long_description=long_description,
    long_description_content_type="text/markdown",
    author='Cloint India Pvt. Ltd',
    author_email='automation@cloint.com',
    url='https://github.com/ClointFusion/ClointFusion',
    keywords=['ClointFusion','RPA','Python','Automation','BOT'],
      install_requires=[            
          
      ],
  classifiers=[
    'Development Status :: 3 - Alpha',
    'Intended Audience :: Developers',      
    'Topic :: Software Development :: Build Tools',
    'License :: OSI Approved :: BSD License',
    'Natural Language :: English',
    'Operating System :: Microsoft :: Windows :: Windows 10',
    'Framework :: Robot Framework',
  ],
  python_requires='>=3.8',
)