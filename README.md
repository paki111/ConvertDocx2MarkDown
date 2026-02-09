# DOCX to Markdown Conversion

PowerShell script for converting DOCX documents to Markdown format with advanced image processing.

## Requirements

- **PowerShell 5.1** or newer (default in Windows 10/11)
- **Pandoc 2.x** or newer - document conversion tool
  - Installation: Download from https://pandoc.org/installing.html
  - Script expects `pandoc.exe` at `C:\pandoc\pandoc.exe` (can be changed in script)
- **.NET Framework** with System.Drawing (default in Windows)

## ConvertDocx2MarkDown.ps1

Converts DOCX files to Markdown with advanced image processing.

**Directories:**
- Source: `c:\temp\docx` (hardcoded)
- Destination: `c:\temp\docx_markdown` (hardcoded)
- Images: `c:\temp\docx_markdown\assets` (hardcoded)

**Features:**
- DOCX → Markdown conversion (GFM with pipe tables)
- Image extraction to `/assets` directory
- Automatic EMF → PNG conversion
- Unique image prefixes (first 12 characters of document name)
- Polish characters removed from prefixes
- Relative image paths (`assets/...`)
- Tables in Markdown format instead of HTML

**Usage:**
```powershell
# 1. Place DOCX files in source directory
Copy-Item *.docx c:\temp\docx\

# 2. Run script
.\ConvertDocx2MarkDown.ps1

# 3. Markdown files will be in: c:\temp\docx_markdown\
# 4. Images will be in: c:\temp\docx_markdown\assets\
```

**Example image prefixes:**
- `Document_name.docx` → `Document_nam_imageX.png`
- `SRS Payments on microservice.docx` → `SRS_Payment_imageX.png`

## Notes

- DOCX script automatically creates directories if they don't exist
- EMF images are automatically converted to PNG for better compatibility
- Complex tables with merged cells are converted to markdown (without rowspan/colspan)
- Polish characters in file names are replaced with ASCII in image prefixes
- Scripts support long paths and OneDrive files

## Troubleshooting

**Error: "execution policy"**
```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
```

**Error: "Pandoc not found"**
- Check if Pandoc is installed: `C:\pandoc\pandoc.exe --version`
- Update path in script if Pandoc is in different location

**Images not displaying in VS Code**
- Check if they are in `assets/` directory
- Check if paths in markdown are relative (`assets/...`, not `c:\temp\...`)

## License

Script is available under MIT License. See file header for details.

Inspired by: [ConvertOneNote2MarkDown](https://github.com/SjoerdV/ConvertOneNote2MarkDown) by Sjoerd de Valk
