from setuptools import setup, find_packages

setup(name='felsimetl',
      version='0.1.20',
      description='ETL for Felsim',
      url='http://github.com/fedegos/felsimetl',
      author='Federico Gosman',
      author_email='federico@equipogsm.com',
      license='MIT',
      packages=find_packages(),
      entry_points={
          'console_scripts': [
              'felsim = felsim.__main__:main'
          ]
      },
      install_requires=[
            'openpyxl==2.4.7'
      ],
      zip_safe=False)