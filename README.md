## PDFToImage
### Extract and convert specific pages from a PDF to various image formats.
---
I wrote this script to use _ExifTool_ for counting the total pages in the PDF and then _pdftocairo_ to convert the cover (page 1) and the back (last page) into images, but it can be edited to process any number or range of pages in a PDF.

This script uses the following Windows command line tools:
- [Windows PowerShell](https://en.wikipedia.org/wiki/PowerShell)
- [Poppler Packaged for Windows (pdftocairo)](https://github.com/oschwartz10612/poppler-windows)
- [ExifTool](https://exiftool.org/)