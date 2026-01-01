# EMF to PNG Converter (Windows)

Windows-only Python utility to convert EMF (Enhanced Metafile) files to PNG using native Win32 GDI rendering.

## Requirements
- Windows
- Python 3.8+
- Dependencies: pillow, pywin32

Install:
pip install pillow pywin32

## Files
main.py — EMF to PNG converter implementation

## Basic Usage
from main import EMFToPNGConverter

conv = EMFToPNGConverter()
conv.emf_file_to_png_file(
    "sample/input.emf",
    "output/output.png"
)

## Features
- Accurate EMF rendering via Windows GDI
- DPI-aware output (default: 300)
- File-based and in-memory conversion
- Base64 and data-URI helpers

## Limitations
- Windows only
- White background only
- Requires desktop GDI session