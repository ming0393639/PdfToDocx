# PdfToDocx
## Description
Read images in the pdf files from 地政資訊網路E點通資訊服務 (https://www.ttt.nat.gov.tw/), and recognize address words in images using Tesseract, finally export .docx file using OpenXml.

## Getting Started

### Dependencies
* .NET Core 3.1
* Open-XML-SDK 2.16.0
* PdfPig 0.1.5
* tesseract 4.1.1

### Executing program

* Build projects and root folder will contain PdfToDocx.exe and template.docx.
* The template.docx could be modified for satisfying the requirement.
* Executing PdfToDocx.exe requires a parameter: pdf file path.
* The testing pdf file is in PdfToDocx/TestPdf folder.
* Example:
```
PdfToDocx.exe "TestPdf\成家美地整棟111H6000278REGA.pdf"
```

## Authors
Ming Wu (ming0393639@hotmail.com)