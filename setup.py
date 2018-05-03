from setuptools import setup, find_packages

setup(name='felsimetl',
      version='0.1',
      description='ETL for Felsim',
      url='http://github.com/fedegos/felsimetl',
      author='Federico Gosman',
      author_email='federico@equipogsm.com',
      license='MIT',
      packages=find_packages(),
      install_requires=[
            'openpyxl'
      ],
      zip_safe=False)