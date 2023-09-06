# -*- coding: utf-8 -*-
"""
Created on Tue Oct 25 18:10:57 2022

@author: hwlee
"""

from cx_Freeze import setup, Executable
import sys

# A list of packages to include in the build (this is to safeguard against cx_freeze missing a package since it automatically detects required packages).
buildOptions = dict(packages = [],
                    excludes = ['numpy_mkl'],
                    include_files = ['PBD_p3d.ui', 'BR_setting.ui', 'print_setting.ui', 'about.ui', './images', './docs', './template'])

base = None
if sys.platform == "win32":
    base = "Win32GUI"

# Assigns default installation path while creating msi file
if 'bdist_msi' in sys.argv:
    sys.argv += ['--initial-target-dir', 'C:\PBD with Perform-3D']
    
exe = [Executable(script="GUI_main.py", base=base, targetName='PBD-p3d.exe', icon='./images/icon_earthquake.ico')]

setup(
    name='PBD-p3d',
    version = '2.3.0',
    author = "CNP Dongyang",
    description = "Performance-Based Design with Perform-3D",
    options = dict(build_exe = buildOptions),
    executables = exe
)

# 359 mb, 1.02 gb
# 324 mb, 941 mb
# 324 mb, 929 mb
# 320 mb, 908 mb