# Document Merge

A Python-based GUI application for merging multiple Word documents (.docx files) into a single document. Features drag-and-drop functionality for easy file management.

## Features

-  **Drag-and-Drop Interface**: Easily add files by dragging and dropping them into the application
-  **Word Document Merging**: Combine multiple .docx files into a single document
-  **File Management**: Remove individual files or clear all files from the list
-  **User-Friendly**: Simple and intuitive graphical interface

## Screenshots

The application provides a clean interface where you can:
- Drag and drop .docx files into the list
- Combine documents into a single Word file
- Manage your file list

## Installation

### Prerequisites

- Python 3.6 or higher

### Step 1: Clone the Repository

```bash
git clone https://github.com/yourusername/documentMerge.git
cd documentMerge
```

### Step 2: Install Dependencies

```bash
pip install -r requirements.txt
```

### Step 3: Run the Application

```bash
python merger.py
```

## Usage

1. **Launch the Application**
   - Run `python merger.py` from the project directory

2. **Add Files**
   - Drag and drop .docx files into the application window
   - Files will appear in the list below

3. **Combine Documents**
   - Click "Combine to Word" to merge all files
   - Choose a save location for the combined document
   - The merged file will be created and saved

4. **Manage Files**
   - Use "Remove Selected File" to remove individual files from the list
   - Use "Clear All Files" to remove all files from the list

## Requirements

The following Python packages are required (see `requirements.txt`):

- `python-docx>=1.1.0` - For reading and writing Word documents
- `tkinterdnd2>=0.3.0` - For drag-and-drop functionality
- `docxcompose>=1.3.6` - For composing/merging documents

## System Requirements

- **Operating System**: Windows, macOS, or Linux
- **Python**: 3.6 or higher

## Troubleshooting

### Common Issues

1. **Drag-and-Drop Not Working**
   - Ensure you're dragging .docx files only
   - Check that the files aren't corrupted or password-protected

2. **Import Errors**
   - Make sure all dependencies are installed: `pip install -r requirements.txt`
   - Verify you're using Python 3.6 or higher

3. **File Permission Errors**
   - Ensure you have write permissions in the directory where you're saving files
   - Close any open Word documents before merging

### Error Messages

- **"No files to combine"**: Add .docx files to the list before combining

## Development

### Project Structure

```
documentMerge/
├── merger.py          # Main application file
├── requirements.txt   # Python dependencies
└── README.md         # This file
```

### Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

---

**Note**: This application is designed for merging Word documents (.docx format). For other file formats, consider converting them to .docx first.

