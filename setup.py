from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need
# fine tuning.
build_options = {'packages': ['os','xlwings','datetime'], 'excludes': []}

# IF UI USE THIS
base = 'Win32GUI'
# IF CONSOLE USE THIS
# base = 'console'

executables = [
    Executable('main.py', base=base)
]

setup(name='AT&T',
      version = '1.2',
      description = 'A tool for AT&T used by Wideout/Innovid Team',
      options = {'build_exe': build_options},
      executables = executables)

# python setup.py build