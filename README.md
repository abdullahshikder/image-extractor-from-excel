# Excel Image Extractor

A Python tool to extract images from Excel (.xlsx) files and rename them based on cell values. This tool extracts images from a specified column (default: column C) and renames them using values from another column (default: column B).

## Features

- ✅ Extracts embedded images from Excel workbooks
- ✅ Automatically renames images based on cell values
- ✅ Supports multiple image formats (PNG, JPG, GIF, BMP, TIFF)
- ✅ Handles duplicate filenames automatically
- ✅ Works with multiple sheets
- ✅ Maps images to their exact cell positions via XML parsing
- ✅ Fallback matching by row order if XML parsing fails

## How It Works

Excel `.xlsx` files are actually ZIP archives containing XML files and media. This tool:
1. Opens the Excel file as a ZIP archive
2. Extracts images from the `xl/media/` folder
3. Parses the drawing XML to map images to their cell positions
4. Reads product names from column B
5. Renames and saves images with sanitized product names

## Installation

1. Clone this repository:
```bash
git clone https://github.com/yourusername/excel-image-extractor.git
cd excel-image-extractor
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage

```python
from extract_images import extract_images_from_excel

# Extract images from column C, rename with column B values
extract_images_from_excel("your_file.xlsx")
```

### Command Line

```bash
python extract_images.py
```

Make sure your Excel file is named `Hermizon images.xlsx` or modify the filename in the script.

### Custom Usage

```python
from extract_images import extract_images_from_excel

# Custom output directory
extract_images_from_excel("your_file.xlsx", output_dir="my_images")
```

## Excel File Structure

The script expects:
- **Column B**: Product names (used for renaming images)
- **Column C**: Images (images to be extracted)

You can modify the script to use different columns by changing the column numbers in the code.

## Output

Images are extracted to the `extracted_images/` directory (or your specified output directory) with filenames based on the product names from column B.

Example output:
```
extracted_images/
├── Product_Name_1.png
├── Product_Name_2.jpg
├── Product_Name_3.png
└── ...
```

## Requirements

- Python 3.6+
- openpyxl >= 3.1.0

## Supported Image Formats

- PNG (.png)
- JPEG (.jpg, .jpeg)
- GIF (.gif)
- BMP (.bmp)
- TIFF (.tiff)

## License

MIT License - feel free to use this project for any purpose.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Acknowledgments

Inspired by the approach used in online Excel image extraction tools, this tool reverse-engineers the Excel file structure to extract embedded images efficiently.

