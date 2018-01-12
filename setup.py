from cx_Freeze import setup, Executable
import os

includefiles = ['logo.ico','multi_comp_settings.cfg','other_applications.prg']
bdist_msi_options = {'upgrade_code': '{ABEC9170-7047-4F17-87A7-42A1E6330C70}'}
build_options = {'include_files':includefiles,"include_msvcr": True}

setup(name = "Computer Info",
	version = "2.0",
	description = "",
	options = {"build_exe":build_options,'bdist_msi': bdist_msi_options},
	executables = [Executable("Computer Info.py",base="Win32GUI",icon="logo.ico",shortcutName="Computer Info",shortcutDir="ProgramMenuFolder")])
