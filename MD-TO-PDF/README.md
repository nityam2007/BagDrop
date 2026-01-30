# MD-TO-PDF

A universal command-line tool to convert Markdown files to beautifully formatted PDFs.

---

## Features

- Clean, professional PDF output
- Table of Contents support
- Customizable margins, fonts, and paper size
- Batch convert all `.md` files at once
- A4 and Letter paper support
- Colored terminal output

---

## Requirements

Install these dependencies:

```bash
# Ubuntu/Debian
sudo apt install pandoc texlive-xetex texlive-fonts-extra

# Arch Linux
sudo pacman -S pandoc texlive-xetex

# macOS (Homebrew)
brew install pandoc
brew install --cask mactex
```

---

## Installation

1. Copy the `MD-TO-PDF` folder to your desired location
2. Make the script executable:

```bash
chmod +x md2pdf.sh
```

3. (Optional) Add to PATH for global access:

```bash
# Add to ~/.bashrc or ~/.zshrc
export PATH="$PATH:/path/to/MD-TO-PDF"
```

---

## Usage

### Basic Conversion

```bash
./md2pdf.sh document.md
# Output: document.pdf
```

### Specify Output Name

```bash
./md2pdf.sh document.md output.pdf
```

### With Table of Contents

```bash
./md2pdf.sh document.md --toc
./md2pdf.sh document.md -t
```

### Custom Margins

```bash
./md2pdf.sh document.md --margin 2cm
./md2pdf.sh document.md -m 1cm
```

### Custom Font Size

```bash
./md2pdf.sh document.md --size 10pt
./md2pdf.sh document.md -s 12pt
```

### Batch Convert All Markdown Files

```bash
./md2pdf.sh --all
./md2pdf.sh --all --toc
```

### Combined Options

```bash
./md2pdf.sh document.md output.pdf --toc --margin 1.5cm --size 11pt
```

---

## Options

| Option | Short | Default | Description |
|--------|-------|---------|-------------|
| `--margin` | `-m` | 1.5cm | Page margin |
| `--size` | `-s` | 11pt | Font size |
| `--font` | `-f` | DejaVu Sans | Font family |
| `--toc` | `-t` | off | Include Table of Contents |
| `--paper` | `-p` | a4paper | Paper size (a4paper/letterpaper) |
| `--all` | - | - | Batch convert all .md files |
| `--help` | `-h` | - | Show help |

---

## Examples

```bash
# Simple conversion
./md2pdf.sh README.md

# Business document with TOC
./md2pdf.sh business_plan.md --toc --margin 2cm

# Compact formatting
./md2pdf.sh notes.md -m 1cm -s 10pt

# Convert all markdown files in folder
./md2pdf.sh --all --toc

# Letter paper (US)
./md2pdf.sh doc.md --paper letterpaper
```

---

## Output Quality

The script uses:
- **XeLaTeX** engine for better Unicode support
- **DejaVu Sans** font (supports most characters including Rs. symbol)
- **Clickable links** in blue color
- **Clean formatting** with sensible defaults

---

## Troubleshooting

### Missing Characters
If you see missing characters (like ₹), install additional fonts:
```bash
sudo apt install fonts-dejavu fonts-noto
```

### Emoji Support
For emoji support, install:
```bash
sudo apt install fonts-noto-color-emoji
```

### Permission Denied
Make the script executable:
```bash
chmod +x md2pdf.sh
```

---

## License

Free to use and modify.

---

*Created by NSheth.in*
