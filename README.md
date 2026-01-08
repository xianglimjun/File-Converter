# Universal File Converter

A lightweight Windows utility to convert images and documents.

## Features
- **Image Conversion**: JPG, PNG, BMP, TGA, GIF → PNG, JPG, BMP, TGA, PDF
- **Document Conversion**: DOCX, PPTX → PDF (requires Microsoft 365)
- **Save as Copy**: Option to preserve original files

## Minimum Files Required

To run the converter, you only need:

| File | Purpose |
|------|---------|
| `converter.exe` | Main application (all-in-one) |

## Building from Source

Additional files needed for compilation:

| File | Purpose |
|------|---------|
| `main.c` | Source code |
| `stb_image.h` | Image loading library |
| `stb_image_write.h` | Image writing library |
| `pdf_writer.h` | Custom PDF generator |
| `build.bat` | Build script (requires Visual Studio) |

## Usage

1. Run `converter.exe` (or `run.bat`)
2. Click "Select File"
3. Choose target format
4. Click "Convert"

## Requirements

- Windows 10/11
- Microsoft 365 (Word/PowerPoint) for document conversion
- Visual Studio 2022 (for building from source only)