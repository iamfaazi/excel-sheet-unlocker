# 🔓 Excel Sheet Unlocker

[![Python Version](https://img.shields.io/badge/python-3.7+-blue.svg)](https://python.org)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey.svg)](https://github.com/iamfaazi/excel-sheet-unlocker)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Downloads](https://img.shields.io/github/downloads/iamfaazi/excel-sheet-unlocker/total.svg)](https://github.com/iamfaazi/excel-sheet-unlocker/releases)

A powerful, cross-platform tool to remove sheet protection from Excel files **without needing the password**. Works on Windows, macOS, and Linux with multiple installation options including **no-Python-required** solutions.

## ✨ Features

- 🔐 **Password-free unlocking** - Remove sheet protection without knowing the password
- 🌐 **Cross-platform** - Works on Windows, macOS, and Linux
- 📁 **Batch processing** - Unlock multiple Excel files at once
- 🚀 **Multiple installation options** - From auto-installers to portable solutions
- 🛡️ **Safe operation** - Creates new files, never modifies originals
- 📊 **Multiple formats** - Supports `.xlsx`, `.xlsm`, and `.xls` files
- 🎯 **User-friendly** - Interactive mode with drag-and-drop support
- ⚡ **Fast processing** - Efficient unlocking algorithm with fallback methods

## 🚀 Quick Start
1. Download [`excel_unlocker.py`](excel_unlocker.py)
2. Install openpyxl: `pip install openpyxl`
3. Run: `python excel_unlocker.py`

## 📋 Usage

### Interactive Mode
```bash
python excel_unlocker.py
```
Then choose:
- **Option 1**: Process a single file (supports drag & drop)
- **Option 2**: Process all Excel files in a folder

### Command Line Mode
```bash
# Single file
python excel_unlocker.py "/path/to/your/file.xlsx"

# Single file with custom output
python excel_unlocker.py "/path/to/file.xlsx" "/path/to/output.xlsx"

# Batch process folder
python excel_unlocker.py "/path/to/folder/"
```

### Examples
```bash
# Windows
python excel_unlocker.py "C:\Documents\protected_file.xlsx"

# macOS
python3 excel_unlocker.py "/Users/john/Documents/protected_file.xlsx"

# Linux
python3 excel_unlocker.py "/home/user/documents/protected_file.xlsx"

# Process entire folder
python excel_unlocker.py "/path/to/excel/files/"
```

## 🛠️ Installation Options

### For Users Without Python

#### macOS Users
```bash
# Install Homebrew (if not installed)
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"

# Install Python
brew install python

# Install dependencies and run
pip3 install openpyxl
python3 excel_unlocker.py
```

#### Linux Users
```bash
# Ubuntu/Debian
sudo apt update && sudo apt install python3 python3-pip
pip3 install openpyxl

# CentOS/RHEL/Fedora
sudo yum install python3 python3-pip  # or dnf
pip3 install openpyxl

# Arch Linux
sudo pacman -S python python-pip
pip install openpyxl
```

### For Excel Users (VBA Alternative)
If you have Excel but no Python:
1. Open Excel
2. Press `Alt + F11`
3. Insert → Module
4. Paste the [VBA code from our documentation](docs/vba-solution.md)
5. Run the macro

## 📖 How It Works

The tool uses the `openpyxl` library to:

1. **Load the Excel file** - Bypasses sheet-level password protection
2. **Identify protected sheets** - Scans all worksheets for protection
3. **Remove protection** - Clears all protection settings using multiple methods:
   - Method 1: Clean protection object replacement
   - Method 2: Direct protection attribute modification  
   - Method 3: Minimal intervention fallback
4. **Save unlocked file** - Creates new file with `_unlocked` suffix
5. **Preserve data integrity** - Only removes restrictions, keeps all data

## 🔧 Advanced Features

### Batch Processing
Process entire folders of Excel files:
```bash
python excel_unlocker.py "/path/to/folder/"
```

### Custom Output Paths
Specify where to save unlocked files:
```bash
python excel_unlocker.py "input.xlsx" "custom_output.xlsx"
```

### Error Handling
The tool includes robust error handling:
- Multiple unlocking methods with fallbacks
- Graceful handling of corrupted protection objects
- Detailed progress reporting
- Safe processing that won't crash on problematic files

## 📁 File Structure

```
excel-sheet-unlocker/
├── excel_unlocker.py          # Main Python script
├── windows_auto_installer.bat # Windows one-click installer
├── unlock_excel.sh           # Linux/macOS installer script
├── docs/
│   ├── vba-solution.md       # VBA macro alternative
│   ├── troubleshooting.md    # Common issues and solutions
│   └── advanced-usage.md     # Advanced features guide
├── examples/
│   ├── sample_protected.xlsx # Sample protected file for testing
│   └── screenshots/          # Usage screenshots
├── README.md
├── LICENSE
└── requirements.txt
```

## ⚠️ Important Notes

### ✅ What This Tool Does
- Removes Excel sheet protection (cell locking, editing restrictions)
- Works with password-protected **sheets** (not workbook passwords)
- Preserves all data, formulas, and formatting
- Creates backup copies automatically

### ❌ What This Tool Cannot Do
- Cannot unlock **workbook-level** passwords (file encryption)
- Cannot recover lost data or repair corrupted files
- Does not work with Excel files encrypted at the file level

### 🔒 Security Considerations
- This tool is intended for legitimate use (recovering your own files)
- Only use on files you own or have permission to modify
- The tool bypasses protection, not encryption

## 🐛 Troubleshooting

### Common Issues

**"Permission Error" - File is open in Excel**
```bash
# Solution: Close Excel and try again
```

**"Module not found: openpyxl"**
```bash
# Solution: Install the package
pip install openpyxl
```

**"Python not recognized"**
```bash
# Windows: Use the auto-installer batch file
# Or add Python to PATH during installation
```

**For more issues, see [Troubleshooting Guide](docs/troubleshooting.md)**

## 🧪 Testing

Test the tool with the provided sample file:
```bash
python excel_unlocker.py examples/sample_protected.xlsx
```

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request. For major changes:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

### Development Setup
```bash
git clone https://github.com/iamfaazi/excel-sheet-unlocker.git
cd excel-sheet-unlocker
pip install -r requirements.txt
python -m pytest tests/
```

## 📊 Compatibility

| Platform | Python Version | Status | Notes |
|----------|----------------|--------|-------|
| Windows 10/11 | 3.7+ | ✅ Fully Supported | Auto-installer available |
| macOS 10.14+ | 3.7+ | ✅ Fully Supported | Homebrew recommended |
| Ubuntu 18.04+ | 3.7+ | ✅ Fully Supported | |
| CentOS 7+ | 3.7+ | ✅ Fully Supported | |
| Debian 9+ | 3.7+ | ✅ Fully Supported | |
| Arch Linux | 3.7+ | ✅ Fully Supported | |

### Excel Formats
- ✅ `.xlsx` (Excel 2007+)
- ✅ `.xlsm` (Excel 2007+ with macros)
- ✅ `.xls` (Excel 97-2003)

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Built with [openpyxl](https://openpyxl.readthedocs.io/) - Python library for Excel files
- Inspired by the need for accessible Excel sheet recovery tools
- Thanks to all contributors and users who provided feedback

## 📞 Support

- 🐛 **Bug Reports**: [Open an issue](https://github.com/iamfaazi/excel-sheet-unlocker/issues)
- 💡 **Feature Requests**: [Discussions](https://github.com/iamfaazi/excel-sheet-unlocker/discussions)
- 📧 **Direct Contact**: faaziahamed@gmail.com

## ⭐ Star History

[![Star History Chart](https://api.star-history.com/svg?repos=iamfaazi/excel-sheet-unlocker&type=Date)](https://star-history.com/#iamfaazi/excel-sheet-unlocker&Date)

---

<div align="center">

**Made with ❤️ for the open source community**

[⬆ Back to Top](#-excel-sheet-unlocker)

</div>
