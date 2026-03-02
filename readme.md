# OCRBulkAdd

windows app (wpf, .NET), that consumes images from **clipboard (ctrl+v)** or per **drag & drop** einzufügen, puts them through **OCR**, output ends up in editable textbox to **add all number to each other**.

## Features

- input image by:
  - **strg+v** (screenshot from clipboard)
  - **drag & drop** (files supported: `.png, .jpg, .jpeg, .bmp, .gif, .tif, .tiff, .webp`)
- OCR (Tesseract) → result displayed in **editable textbox** to correct OCR mistakes
- button: **sum numbers** → sums all numbers found in OCR textbox (also recognizes german format `12.345,67`)
- 
## Requirements

- windows 10/11
- .NET 9.0
- package: `TesseractOCR`
- `tessdata` folder including:
  - `eng.traineddata`
  - optional `deu.traineddata`

### Tessdata Ordner

app expects language files in:

- `.\tessdata\` (root level of project)
