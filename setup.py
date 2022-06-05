from setuptools import setup, find_packages
 
classifiers = [
  'Development Status :: 5 - Production/Stable',
  'Intended Audience :: Education',
  'Operating System :: Microsoft :: Windows :: Windows 10',
  'License :: OSI Approved :: MIT License',
  'Programming Language :: Python :: 3'
]
 
setup(
  name='excelwin',
  version='0.0.1',
  description='A package that can be used to manipulate different formats of excel files.',
  long_description=open('README.txt').read() + '\n\n' + open('CHANGELOG.txt').read(),
  url='',  
  author='Peddireddy BALA GOPAL REDDY',
  author_email='balupeddireddy08@gmail.com',
  license='MIT', 
  classifiers=classifiers,
  keywords=['xlsm','xlsx','excel','win32'], 
  # packages=find_packages(),
  py_modules = ['code'],
  install_requires=[] 
)