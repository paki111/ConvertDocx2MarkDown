<#
.SYNOPSIS
Convert DOCX to Markdown with Advanced Image Processing

.DESCRIPTION
This script converts DOCX files to Markdown format with automatic image extraction,
EMF to PNG conversion, and unique image prefixes to avoid naming conflicts.

Inspired by: https://github.com/SjoerdV/ConvertOneNote2MarkDown

.NOTES
Author: [Your Name]
License: MIT License

Copyright (c) 2026 [Your Name]

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

Acknowledgment: This script was inspired by the ConvertOneNote2MarkDown project
by Sjoerd de Valk (https://github.com/SjoerdV/ConvertOneNote2MarkDown).

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

.EXAMPLE
.\ConvertDocx2MarkDown.ps1
#>

Function Remove-InvalidFileNameChars {
  param(
    [Parameter(Mandatory = $true,
      Position = 0,
      ValueFromPipeline = $true,
      ValueFromPipelineByPropertyName = $true)]
    [String]$Name
  )
  $newName = $Name.Split([IO.Path]::GetInvalidFileNameChars()) -join '_'
  return ($newName.Replace(" ", "_"))
}

# hardcoded paths
$sourcepath = "c:\temp\docx"
$destpath = "c:\temp\docx_markdown"

# create folders if they don't exist
if (-not (Test-Path -Path $sourcepath)) {
  New-Item -Path $sourcepath -ItemType Directory -Force | Out-Null
  Write-Host "Created source folder: $sourcepath" -ForegroundColor Green
  Write-Host "Please place your .docx files in $sourcepath and run the script again." -ForegroundColor Yellow
  exit
}

if (-not (Test-Path -Path $destpath)) {
  New-Item -Path $destpath -ItemType Directory -Force | Out-Null
  Write-Host "Created destination folder: $destpath" -ForegroundColor Green
}

# create assets folder
$assetspath = "$destpath\assets"
if (-not (Test-Path -Path $assetspath)) {
  New-Item -Path $assetspath -ItemType Directory -Force | Out-Null
  Write-Host "Created assets folder: $assetspath" -ForegroundColor Green
}

# check if pandoc is available
$pandocPath = "C:\pandoc\pandoc.exe"
if (-not (Test-Path -Path $pandocPath)) {
  Write-Host "Pandoc not found at $pandocPath" -ForegroundColor Red
  Write-Host "Please install Pandoc from https://pandoc.org/installing.html" -ForegroundColor Yellow
  exit
}

Write-Host "Starting conversion..." -ForegroundColor Cyan
$totalerr = ""
$convertedCount = 0

# get all docx files from source folder
$docxFiles = Get-ChildItem -Path $sourcepath -Filter "*.docx" -Recurse

if ($docxFiles.Count -eq 0) {
  Write-Host "No .docx files found in $sourcepath" -ForegroundColor Yellow
  exit
}

Write-Host "Found $($docxFiles.Count) .docx file(s) to convert" -ForegroundColor Green

foreach ($docxFile in $docxFiles) {
  try {
    $fileName = $docxFile.BaseName | Remove-InvalidFileNameChars
    # use -replace with regex for case-insensitive path removal
    $relativePath = $docxFile.DirectoryName -replace [regex]::Escape($sourcepath), "" 
    $relativePath = $relativePath.TrimStart("\")
    
    # create corresponding folder structure in destination
    if ($relativePath) {
      $targetFolder = Join-Path -Path $destpath -ChildPath $relativePath
      if (-not (Test-Path -Path $targetFolder)) {
        New-Item -Path $targetFolder -ItemType Directory -Force | Out-Null
      }
      $mdPath = Join-Path -Path $targetFolder -ChildPath "$fileName.md"
    } else {
      $mdPath = Join-Path -Path $destpath -ChildPath "$fileName.md"
    }

    Write-Host "Converting: $($docxFile.Name)" -ForegroundColor Cyan

    # convert docx to markdown with tables in markdown format and extract media
    # --extract-media extracts images to specified folder
    # -t gfm+pipe_tables ensures markdown tables instead of HTML
    # --wrap=none prevents line wrapping
    # --markdown-headings=atx uses # style headings
    # Note: Complex tables with merged cells may still render as HTML
    & $pandocPath -f docx -t gfm+pipe_tables-raw_html `
      -i "$($docxFile.FullName)" `
      -o "$mdPath" `
      --wrap=none `
      --markdown-headings=atx `
      --extract-media="$destpath" 2>&1 | Out-Null

    if ($LASTEXITCODE -eq 0) {
      Write-Host "  [OK] Converted to: $fileName.md" -ForegroundColor Green
      
      # create image prefix from document name (max 12 chars) without Polish characters
      $imagePrefix = $fileName
      # replace Polish characters with ASCII equivalents using regex
      $imagePrefix = $imagePrefix -replace '\u0105|\u0104', 'a'  # ą Ą
      $imagePrefix = $imagePrefix -replace '\u0107|\u0106', 'c'  # ć Ć
      $imagePrefix = $imagePrefix -replace '\u0119|\u0118', 'e'  # ę Ę
      $imagePrefix = $imagePrefix -replace '\u0142|\u0141', 'l'  # ł Ł
      $imagePrefix = $imagePrefix -replace '\u0144|\u0143', 'n'  # ń Ń
      $imagePrefix = $imagePrefix -replace '\u00f3|\u00d3', 'o'  # ó Ó
      $imagePrefix = $imagePrefix -replace '\u015b|\u015a', 's'  # ś Ś
      $imagePrefix = $imagePrefix -replace '\u017a|\u0179', 'z'  # ź Ź
      $imagePrefix = $imagePrefix -replace '\u017c|\u017b', 'z'  # ż Ż
      
      if ($imagePrefix.Length -gt 12) {
        $imagePrefix = $imagePrefix.Substring(0, 12)
      }
      $imagePrefix = $imagePrefix + "_"
      
      # move images from media to assets folder with prefix
      $mediaFolder = "$destpath\media"
      $imageMapping = @{}  # store original name -> prefixed name mapping
      
      if (Test-Path -Path $mediaFolder) {
        Get-ChildItem -Path $mediaFolder | ForEach-Object {
          $originalName = $_.Name
          $newImageName = "$imagePrefix$originalName"
          $targetPath = Join-Path -Path $assetspath -ChildPath $newImageName
          Move-Item -Path $_.FullName -Destination $targetPath -Force
          
          # convert EMF to PNG
          if ($_.Extension -eq ".emf") {
            try {
              Add-Type -AssemblyName System.Drawing
              $emfPath = $targetPath
              $pngPath = $targetPath -replace '\.emf$', '.png'
              
              $image = [System.Drawing.Image]::FromFile($emfPath)
              $image.Save($pngPath, [System.Drawing.Imaging.ImageFormat]::Png)
              $image.Dispose()
              
              Remove-Item -Path $emfPath -Force
              Write-Host "  Converted $newImageName to PNG" -ForegroundColor DarkGray
              
              # update mapping to reflect PNG extension (original .emf -> prefixed .png)
              $originalNamePng = $originalName -replace '\.emf$', '.png'
              $newImageNamePng = $newImageName -replace '\.emf$', '.png'
              $imageMapping[$originalNamePng] = $newImageNamePng
            }
            catch {
              Write-Host "  Warning: Could not convert $newImageName to PNG: $($_.Exception.Message)" -ForegroundColor Yellow
              $imageMapping[$originalName] = $newImageName
            }
          }
          else {
            # non-EMF files - direct mapping
            $imageMapping[$originalName] = $newImageName
          }
        }
        Remove-Item -Path $mediaFolder -Force -ErrorAction SilentlyContinue
        Write-Host "  Moved images to assets folder with prefix: $imagePrefix" -ForegroundColor Green
      }
      
      # fix image paths in markdown
      if (Test-Path -Path $mdPath) {
        $content = Get-Content -LiteralPath $mdPath -Raw
        
        # FIRST: replace absolute paths to relative (both media and assets)
        # Handle both forward and backward slashes (Pandoc creates mixed paths like c:\path/media/)
        $destPathEscaped = [regex]::Escape($destpath)
        
        # Replace various path patterns to assets
        $content = $content -replace "$destPathEscaped[/\\]media[/\\]", "assets/"
        $content = $content -replace "$destPathEscaped[/\\]assets[/\\]", "assets/"
        $content = $content -replace "media/", "assets/"
        
        # Remove remaining absolute path prefix (both slashes)
        $content = $content -replace "$destPathEscaped[/\\]", ""
        
        # SECOND: replace .emf extensions with .png in all image references
        $content = $content -replace '(assets/[^"'')\s]+)\.emf', '$1.png'
        
        # THEN: replace each original image name with prefixed version
        foreach ($originalName in $imageMapping.Keys) {
          $prefixedName = $imageMapping[$originalName]
          $originalNameEscaped = [regex]::Escape($originalName)
          
          # replace in various contexts: src="...", src='...', ](...)
          $content = $content -replace "([""'(]assets/)$originalNameEscaped", "`$1$prefixedName"
        }
        
        Set-Content -LiteralPath $mdPath -Value $content -NoNewline
        Write-Host "  Fixed image paths" -ForegroundColor Green
      }
      
      $convertedCount++
    } else {
      Write-Host "  Conversion failed" -ForegroundColor Red
      $totalerr += "Error converting '$($docxFile.Name)': Pandoc returned exit code $LASTEXITCODE`r`n"
    }
  }
  catch {
    Write-Host "  Error: $($Error[0].ToString())" -ForegroundColor Red
    $totalerr += "Error converting '$($docxFile.Name)': $($Error[0].ToString())`r`n"
  }
  
  Write-Host ""
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Conversion complete!" -ForegroundColor Green
Write-Host "Converted: $convertedCount of $($docxFiles.Count) file(s)" -ForegroundColor Green
Write-Host "Output location: $destpath" -ForegroundColor Green
Write-Host "Assets location: $assetspath" -ForegroundColor Green

if ($totalerr) {
  Write-Host "`nErrors encountered:" -ForegroundColor Yellow
  Write-Host $totalerr -ForegroundColor Red
}
