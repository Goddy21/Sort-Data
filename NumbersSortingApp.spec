# -*- mode: python ; coding: utf-8 -*-

import os
import sys

block_cipher = None

# Get the current working directory
py_source = 'sort_kivy.py'
icon_source = r"C:\Goddie\Numbers-sort\icon.png"

a = Analysis(
    [py_source],
    pathex=[],
    binaries=[],
    datas=[(icon_source, '.')],  # Include the icon file in the output
    hiddenimports=[
        'kivy',
        'kivy.app',
        'kivy.uix.screenmanager',
        'kivy.uix.boxlayout',
        'kivy.uix.gridlayout',  # ADD THIS
        'kivy.uix.label',  # ADD THIS
        'kivy.uix.textinput',  # ADD THIS
        'kivy.uix.button',  # ADD THIS
        'kivy.uix.popup',  # ADD THIS
        'kivy.core.window',
        'kivy.utils',  # ADD THIS
        'kivy.graphics',  # ADD THIS
        'plyer',
        'plyer.platforms',  # Explicitly include plyer.platforms
        'plyer.platforms.win',  # Explicitly include plyer.platforms.win (or whatever platform)
        'plyer.platforms.win.filechooser',  # ADD THIS
        'comtypes',
        'win32com',
        'pandas',  # ADD THIS
        'datetime',  # ADD THIS
        're',  # ADD THIS
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='NumbersSortingApp',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon_source,  # ADD THIS
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='NumbersSortingApp',
)
