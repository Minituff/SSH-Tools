from setuptools import (setup, find_packages)

setup(name='SSH-Tools',
      version='0.0.1',
      description='SSH-Tools',
      url='https://github.com/Minituff/SSH-Tools',
      author='James Tufarelli',
      author_email='james_tufarelli@bah.com',
      license='MIT',
      packages=find_packages(),
      zip_safe=False,
      install_requires=[
          'netmiko',
          'openpyxl',
      ])