from setuptools import setup, find_packages

setup(name='poi4py',
      description='Apache POI wrapper for better access of MS Office files',
      author='Youngrok Pak',
      author_email='pak.youngrok@gmail.com',
      keywords= 'excel word powerpoint poi java',
      url='https://github.com/youngrok/poi4py',
      version='0.0.6',
      packages=find_packages(),
      package_data={'poi4py': ['java/*'],},
      install_requires=['JPype1-py3'],
      classifiers = [
                     'Development Status :: 3 - Alpha',
                     'Topic :: System :: Installation/Setup',
                     'License :: OSI Approved :: BSD License']
      )
