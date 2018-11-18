"""
****************************************
* Author: SIRIUS
* Email: xuqingskywolf@outlook.com
* Created Time: 2018/11/18 23:17
****************************************
"""

from setuptools import setup, find_packages

setup(name='pywin32-excel',
      version="0.1.0",
      description='win32com excel helper',
      author='SIRIUS',
      author_email='xuqingskywolf@outlook.com',
      packages=find_packages(),
      install_requires=['pywin32'],
      keywords=('pip', 'pywin32', 'excel'),
      license="MIT Licence",
      url="https://github.com/skywolf123/pywin32-excel",
      include_package_data=True,
      platforms="windows")
