import os
import sys
import subprocess

def build():
    print("🛠️ Generating PyInstaller Spec File...")
    
    # This spec file tells PyInstaller EXACTLY how to bundle Streamlit correctly
    spec_content = """# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files, copy_metadata

datas = []
# Automatically grabs the static HTML/CSS files and necessary metadata
datas += collect_data_files('streamlit')
datas += copy_metadata('streamlit')

a = Analysis(
    ['run_app.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    # CRITICAL FIX: Explicitly tell PyInstaller to include the hidden magic_funcs module
    hiddenimports=[
        'streamlit', 
        'streamlit.runtime.scriptrunner.magic_funcs'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='run_app',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
"""
    # Write the spec file
    with open("run_app.spec", "w", encoding="utf-8") as f:
        f.write(spec_content)

    print("🚀 Running PyInstaller...")
    # Run PyInstaller using the newly generated spec file
    subprocess.run([sys.executable, "-m", "PyInstaller", "run_app.spec", "--clean", "-y"])
    print("\n🎉 Build complete! Check the 'dist' folder for your new run_app.exe.")

if __name__ == "__main__":
    build()