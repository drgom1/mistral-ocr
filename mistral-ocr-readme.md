# 🔍 Mistral OCR Tool

A modern, user-friendly desktop application for extracting text from documents using Mistral AI's OCR API. Features a sleek dark theme UI with drag-and-drop functionality.

![Python](https://img.shields.io/badge/python-3.7+-blue.svg)
![Platform](https://img.shields.io/badge/platform-windows%20%7C%20macos%20%7C%20linux-lightgrey)
![License](https://img.shields.io/badge/license-MIT-green.svg)

## ✨ Features

- **🎯 Drag & Drop Interface**: Simply drag files into the application or browse for them
- **📄 Multiple Format Support**: Process PDFs, images (JPG, PNG, GIF), Word documents, and PowerPoint presentations
- **🎨 Modern Dark Theme**: Easy on the eyes with a professional appearance
- **📁 Batch Processing**: Process multiple files at once
- **📝 Flexible Output**: Save results as plain text (.txt) or Word documents (.docx)
- **🖼️ Image Extraction**: Optionally include images from documents with configurable limits
- **📊 Activity Logging**: Real-time processing status and clickable output links
- **📂 Recent Outputs**: Quick access to your last 10 processed files
- **🔐 Secure API Key Storage**: Toggle visibility for API key input
- **⚡ Cross-Platform**: Works on Windows, macOS, and Linux

## 📋 Requirements

- Python 3.7 or higher
- Mistral AI API key ([Get one here](https://console.mistral.ai/))

## 🚀 Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/yourusername/mistral-ocr-tool.git
   cd mistral-ocr-tool
   ```

2. **Create a virtual environment** (recommended)
   ```bash
   python -m venv venv
   
   # On Windows
   venv\Scripts\activate
   
   # On macOS/Linux
   source venv/bin/activate
   ```

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

### Dependencies

Create a `requirements.txt` file with:
```
tkinterdnd2>=0.3.0
requests>=2.28.0
python-docx>=0.8.11
```

## 🎮 Usage

1. **Run the application**
   ```bash
   python mistral_ocr_tool.py
   ```

2. **Enter your Mistral API key** in the secure input field

3. **Add files** by:
   - Dragging and dropping them into the designated area
   - Clicking the drop area to browse for files
   - Using Ctrl+O (Cmd+O on macOS) shortcut

4. **Configure options**:
   - Choose output format (Text or Word)
   - Enable/disable image inclusion
   - Set image limit (0-50)

5. **Click "Process Documents"** to start OCR processing

6. **Access your results**:
   - Click on processed files in the Recent Outputs section
   - Use the "📁 Output Folder" button
   - Check the activity log for clickable links

## 🎯 Supported File Formats

| Format | Extensions | Description |
|--------|------------|-------------|
| PDF | `.pdf` | Portable Document Format |
| Images | `.jpg`, `.jpeg`, `.png`, `.gif` | Common image formats |
| Word | `.docx` | Microsoft Word documents |
| PowerPoint | `.pptx` | Microsoft PowerPoint presentations |

## ⚙️ Configuration Options

### Output Format
- **Text (.txt)**: Plain text with markdown formatting preserved
- **Word (.docx)**: Formatted document with proper headings

### Image Processing
- **Include Images**: Extract and include base64-encoded images
- **Image Limit**: Control the maximum number of images per document (0-50)

## 🔧 Keyboard Shortcuts

- `Ctrl+O` - Add files
- `Right-click on log` - Show context menu (Copy/Clear)

## 📸 Screenshots

### Main Interface
![Main Interface](screenshots/main-interface.png)
*The main application window with drag-and-drop area*

### Processing Files
![Processing](screenshots/processing.png)
*Multiple files being processed with progress indicator*

### Results View
![Results](screenshots/results.png)
*Completed processing with recent outputs displayed*

## 🏗️ Project Structure

```
mistral-ocr-tool/
├── mistral_ocr_tool.py    # Main application file
├── requirements.txt        # Python dependencies
├── README.md              # This file
└── screenshots/           # Application screenshots
    ├── main-interface.png
    ├── processing.png
    └── results.png
```

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request. For major changes, please open an issue first to discuss what you would like to change.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 🐛 Known Issues

- Large files (>100MB) are not supported due to API limitations
- Processing time varies based on document complexity and size
- Network connectivity is required for API calls

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- [Mistral AI](https://mistral.ai/) for providing the OCR API
- [tkinterdnd2](https://github.com/pmgagne/tkinterdnd2) for drag-and-drop functionality
- [python-docx](https://python-docx.readthedocs.io/) for Word document generation

## 📞 Support

If you encounter any issues or have questions:
- Open an issue on GitHub
- Check existing issues for solutions
- Ensure your API key is valid and has sufficient credits

---

**Note**: This tool requires an active internet connection and a valid Mistral AI API key to function. OCR accuracy depends on document quality and the Mistral AI OCR model's capabilities.