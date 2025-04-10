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
    datas=[(icon_source, '.')], 
    hiddenimports=[
        'kivy',
        'kivy.app',
        'kivy.uix.screenmanager',
        'kivy.uix.boxlayout',
        'kivy.uix.gridlayout',
        'kivy.uix.label',  
        'kivy.uix.textinput',  
        'kivy.uix.button',  
        'kivy.uix.popup',
        'kivy.core.window',
        'kivy.utils', 
        'kivy.graphics',  
        'plyer',
        'plyer.platforms',  
        'plyer.platforms.win',  
        'plyer.platforms.win.filechooser',  
        'comtypes',
        'win32com',
        'pandas',  
        'datetime',  
        're',  
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
    icon=icon_source, 
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
