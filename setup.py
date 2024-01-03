from cx_Freeze import setup, Executable # pip install cx_Freeze

# Dependencies are automatically detected, but it might need
# fine tuning.
build_options = {'packages': [], 'excludes': []}

import sys
base = 'Win32GUI' if sys.platform=='win32' else None

executables = [
    Executable('main.py', base=base)
]

setup(name='Converter',
      version = '1.0',
      description = 'A simple converter',
      options = {'build_exe': build_options},
      executables = executables)

#  python setup.py build