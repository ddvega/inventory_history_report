from setuptools import setup, find_packages

setup(name='inventory_history',
      version='1.0',
      description='A computer science calculator',
      url='',
      author='darius',
      packages=find_packages(),
      zip_safe=False,
      author_email='',
      license='',
      install_requires=['pandas', 'openpyxl', 'numpy', 'xlsxwriter', 'xlrd']
      )
