# Universal File Converter

A lightweight Windows utility to convert images and documents (single-executable).

## Features
- **Image Conversion**: JPG, PNG, WebP, GIF → PNG, JPG, WebP, PDF, Animated GIF*(multiple image select only)
- **Document Conversion**: DOCX, PPTX → PDF (requires Microsoft 365)
- **Multi-File support**: Batch convert multiple files at once.
- **Single-Executable**: Only the .exe file is required.

## Building from Source

To build the converter, you only need these essential files:

| File / Folder | Purpose |
|---------------|---------|
| `main.c` | Core logic |
| `pdf_writer.h` | PDF generator header |
| `stb_image.h` | Image loading |
| `stb_image_write.h` | Image writing |
| `msf_gif.h` | GIF orientation |
| `libwebp/` | WebP library & headers |
| `build.bat` | Compilation script |
| `LICENSE` / `README.md` | Standard docs |

## Usage

1. Download 'File-Converter.exe' from releases or click https://github.com/xianglimjun/File-Converter/releases/latest
2. Run `File-Converter.exe`
3. Click "Select File(s)"
4. Choose target format
5. Click "Convert"

## Requirements

- Windows 10/11
- Microsoft 365 (Word/PowerPoint) for document conversion

- Visual Studio 2022* (for building from source only)

