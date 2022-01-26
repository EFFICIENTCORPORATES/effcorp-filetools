import pathlib
from setuptools import setup


#the Directory Contatining ths file

HERE=pathlib.Path(__file__).parent

#The Text of the Readme FIle

README=(HERE/"README.md").read_text()


#this call to setup() does all the work


setup(
	name = 'effcorp-filetools',
	version = '1.2.0',
	py_modules = ['effcorp-filetools'],
	packages=["filetools"],
	include_package_data=True,
	author = 'Efficient_Corporates',
	author_email = 'efficientcorporates.info@gmail.com',
	install_requires=['pandas','numpy','openpyxl','XlsxWriter','reportlab','PyPDF2','pdfminer','datetime','tabula','tabulate','PyPDF4','fpdf'],
	url = 'https://github.com/EFFICIENTCORPORATES/effcorp-filetools',
	description = 'A python module to simplify various files related work in your day to day office or personal work',
	long_description=README,
	long_description_content_type="text/markdown",
	license="GNU GP License",
	classifiers=[
        "License :: OSI Approved :: GNU General Public License v3 or later (GPLv3+)",
        "Programming Language :: Python :: 3.6",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
    ],
    entry_points={"console_scripts":["filetools=filetools.__main__:main",]},)
