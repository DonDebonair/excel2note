import distribute_setup
distribute_setup.use_setuptools()

from setuptools import setup
from setuptools import find_packages

__author__ = 'Daan Debie'
__version__ = '0.5.0'
__licence__ = 'MIT'

setup(
    name='excel2note',
    version=__version__,
    packages = find_packages(),
    url='https://github.com/DandyDev/excel2note',
    license=__licence__,
    author=__author__,
    author_email='debie.daan@gmail.com',
    description='Utility to convert rows in an Excel sheet to notes dumped into an .enex file that can be imported into Evernote',
    install_requires='openpyxl',
    entry_points={
        'console_scripts': [
            'excel2note = excel2note:main',
        ],
    }
)
