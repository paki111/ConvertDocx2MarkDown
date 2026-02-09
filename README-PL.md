# Konwersja DOCX do Markdown

Skrypt PowerShell do konwersji dokumentów DOCX do formatu Markdown z zaawansowanym przetwarzaniem obrazów.

## Wymagania

- **PowerShell 5.1** lub nowszy (domyślnie w Windows 10/11)
- **Pandoc 2.x** lub nowszy - narzędzie do konwersji dokumentów
  - Instalacja: Pobierz z https://pandoc.org/installing.html
  - Skrypt oczekuje `pandoc.exe` w `C:\pandoc\pandoc.exe` (możesz zmienić ścieżkę w skrypcie)
- **.NET Framework** z System.Drawing (domyślnie w Windows)

## ConvertDocx2MarkDown.ps1

Konwertuje pliki DOCX do Markdown z zaawansowanym przetwarzaniem obrazów.

**Katalogi:**
- Źródłowy: `c:\temp\docx` (zakodowany na stałe)
- Docelowy: `c:\temp\docx_markdown` (zakodowany na stałe)
- Obrazy: `c:\temp\docx_markdown\assets` (zakodowany na stałe)

**Funkcje:**
- Konwersja DOCX → Markdown (GFM z pipe tables)
- Ekstrakcja obrazów do katalogu `/assets`
- Automatyczna konwersja EMF → PNG
- Unikalne prefiksy obrazów (pierwsze 12 znaków nazwy dokumentu)
- Usuwanie polskich znaków z prefiksów
- Względne ścieżki do obrazów (`assets/...`)
- Tabele w formacie Markdown zamiast HTML

**Użycie:**
```powershell
# 1. Umieść pliki DOCX w katalogu źródłowym
Copy-Item *.docx c:\temp\docx\

# 2. Uruchom skrypt
.\ConvertDocx2MarkDown.ps1

# 3. Pliki markdown będą w: c:\temp\docx_markdown\
# 4. Obrazy będą w: c:\temp\docx_markdown\assets\
```

**Przykładowe prefiksy obrazów:**
- `Dokument_nazwa.docx` → `Dokument_naz_imageX.png`
- `SRS Płatności na mikroserwisie.docx` → `SRS_Platnos_imageX.png`

## Uwagi

- Skrypt DOCX automatycznie tworzy katalogi jeśli nie istnieją
- Obrazy EMF są automatycznie konwertowane na PNG dla lepszej kompatybilności
- Złożone tabele z merged cells są konwertowane na markdown (bez rowspan/colspan)
- Polskie znaki w nazwach plików są zastępowane ASCII w prefiksach obrazów
- Skrypty obsługują długie ścieżki i pliki na OneDrive

## Rozwiązywanie problemów

**Błąd: "execution policy"**
```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
```

**Błąd: "Pandoc not found"**
- Sprawdź czy Pandoc jest zainstalowany: `C:\pandoc\pandoc.exe --version`
- Zaktualizuj ścieżkę w skrypcie jeśli Pandoc jest w innej lokalizacji

**Obrazy nie wyświetlają się w VS Code**
- Sprawdź czy są w katalogu `assets/`
- Sprawdź czy ścieżki w markdown są względne (`assets/...`, nie `c:\temp\...`)

## Licencja

Skrypt jest dostępny na licencji MIT. Zobacz nagłówek pliku dla szczegółów.

Inspiracja: [ConvertOneNote2MarkDown](https://github.com/SjoerdV/ConvertOneNote2MarkDown) by Sjoerd de Valk
