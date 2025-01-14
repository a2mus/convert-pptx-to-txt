# PPTX to Text Converter

A Python application with a PyQt6 GUI that converts PowerPoint (.pptx) files to text using three different methods.

## Features
- **PPTX Conversion**: Uses python-pptx library for basic text extraction
- **MarkItDown Conversion**: Uses custom MarkItDown converter for markdown-formatted output
- **Spire Conversion**: Uses Spire.Presentation library for advanced text extraction including SmartArt

## Installation

1. Clone the repository:
```bash
git clone https://github.com/a2mus/pptx-to-text-converter.git
cd pptx-to-text-converter
```

2. Create a virtual environment:
```bash
python -m venv venv
```

3. Activate the virtual environment:
- Windows: `venv\Scripts\activate`
- macOS/Linux: `source venv/bin/activate`

4. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

Run the application:
```bash
python pptx_to_text_qt.py
```

1. Select a .pptx file using the "Browse" button
2. Choose your preferred conversion method:
   - PPTX Conversion: Basic text extraction
   - MarkItDown Conversion: Markdown-formatted output
   - Spire Conversion: Advanced text extraction with SmartArt support
3. View the extracted text in the text area
4. Save the output as:
   - Text (.txt)
   - Markdown (.md)
   - HTML (.html)

## Technical Details

### Conversion Methods

#### 1. PPTX Conversion (python-pptx)
- Extracts basic text content
- Handles tables and grouped shapes
- Limited SmartArt support
- Pros: Lightweight, fast
- Cons: Limited formatting preservation

#### 2. MarkItDown Conversion
- Custom converter for markdown output
- Preserves some formatting
- Pros: Better formatting than basic text
- Cons: Limited SmartArt support

#### 3. Spire Conversion
- Uses Spire.Presentation library
- Full SmartArt support
- Better text extraction from complex layouts
- Pros: Most comprehensive text extraction
- Cons: Requires Spire.Presentation library

## Challenges and Solutions

1. **SmartArt Extraction**
   - Challenge: Different libraries handle SmartArt differently
   - Solution: Implemented robust SmartArt handling in Spire conversion
   - Used attribute checking for better compatibility

2. **UI Layout**
   - Challenge: Radio buttons not appearing correctly
   - Solution: Fixed layout management to ensure all options are visible

3. **File Format Support**
   - Challenge: Different output formats required different handling
   - Solution: Implemented separate save handlers for each format

4. **Performance**
   - Challenge: Large files could cause UI freezing
   - Solution: Added progress bar to show conversion progress

## Future Improvements

- Add batch processing for multiple files
- Implement PDF export option
- Add support for other presentation formats
- Improve error handling and user feedback
- Add dark/light theme switching
- Implement text formatting options

## Requirements

- Python 3.8+
- PyQt6
- python-pptx
- Spire.Presentation
- markdown
- markitdown

## License

MIT License
