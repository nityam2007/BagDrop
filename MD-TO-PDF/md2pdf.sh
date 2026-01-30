#!/bin/bash
# ============================================================
# MD-TO-PDF: Universal Markdown to PDF Converter
# ============================================================
# Usage: ./md2pdf.sh <input.md> [output.pdf] [options]
# ============================================================

set -e

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# Default settings
MARGIN="1.5cm"
FONT_SIZE="11pt"
FONT="DejaVu Sans"
TOC=false
PAPER="a4paper"

# Help function
show_help() {
    echo -e "${BLUE}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${NC}"
    echo -e "${YELLOW}  MD-TO-PDF - Universal Markdown to PDF Converter${NC}"
    echo -e "${BLUE}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${NC}"
    echo ""
    echo -e "${GREEN}Usage:${NC}"
    echo "  ./md2pdf.sh <input.md> [output.pdf] [options]"
    echo ""
    echo -e "${GREEN}Options:${NC}"
    echo "  -m, --margin <size>    Set margin (default: 1.5cm)"
    echo "  -f, --font <name>      Set font (default: DejaVu Sans)"
    echo "  -s, --size <pt>        Set font size (default: 11pt)"
    echo "  -t, --toc              Include Table of Contents"
    echo "  -p, --paper <size>     Paper size: a4paper, letterpaper (default: a4paper)"
    echo "  -h, --help             Show this help"
    echo ""
    echo -e "${GREEN}Examples:${NC}"
    echo "  ./md2pdf.sh README.md"
    echo "  ./md2pdf.sh doc.md output.pdf --toc"
    echo "  ./md2pdf.sh doc.md --margin 2cm --toc"
    echo "  ./md2pdf.sh doc.md -m 1cm -s 10pt -t"
    echo ""
    echo -e "${GREEN}Batch Convert (all .md files):${NC}"
    echo "  ./md2pdf.sh --all"
    echo "  ./md2pdf.sh --all --toc"
    echo ""
    exit 0
}

# Check dependencies
check_deps() {
    if ! command -v pandoc &> /dev/null; then
        echo -e "${RED}Error: pandoc is not installed${NC}"
        echo "Install with: sudo apt install pandoc"
        exit 1
    fi
    
    if ! command -v xelatex &> /dev/null; then
        echo -e "${RED}Error: xelatex is not installed${NC}"
        echo "Install with: sudo apt install texlive-xetex"
        exit 1
    fi
}

# Convert single file
convert_file() {
    local input="$1"
    local output="$2"
    
    if [[ ! -f "$input" ]]; then
        echo -e "${RED}Error: File not found: $input${NC}"
        exit 1
    fi
    
    # Auto-generate output name if not provided
    if [[ -z "$output" ]]; then
        output="${input%.md}.pdf"
    fi
    
    echo -e "${BLUE}Converting:${NC} $input -> $output"
    
    # Build pandoc command
    local cmd="pandoc \"$input\" -o \"$output\" --pdf-engine=xelatex"
    cmd+=" -V geometry:\"$PAPER,margin=$MARGIN\""
    cmd+=" -V fontsize=$FONT_SIZE"
    cmd+=" -V mainfont=\"$FONT\""
    cmd+=" -V colorlinks=true"
    cmd+=" -V linkcolor=blue"
    
    if [[ "$TOC" == true ]]; then
        cmd+=" --toc --toc-depth=3"
    fi
    
    # Execute
    eval $cmd 2>/dev/null
    
    if [[ $? -eq 0 ]]; then
        local size=$(du -h "$output" | cut -f1)
        echo -e "${GREEN}✓ Success:${NC} $output ($size)"
    else
        echo -e "${RED}✗ Failed:${NC} $input"
        return 1
    fi
}

# Batch convert all .md files
convert_all() {
    echo -e "${BLUE}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${NC}"
    echo -e "${YELLOW}  Batch Converting All Markdown Files${NC}"
    echo -e "${BLUE}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${NC}"
    
    local count=0
    local success=0
    
    for file in *.md; do
        if [[ -f "$file" ]]; then
            ((count++))
            if convert_file "$file" ""; then
                ((success++))
            fi
        fi
    done
    
    echo -e "${BLUE}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${NC}"
    echo -e "${GREEN}Converted: $success/$count files${NC}"
}

# Main
check_deps

# Parse arguments
INPUT=""
OUTPUT=""
BATCH=false

while [[ $# -gt 0 ]]; do
    case $1 in
        -h|--help)
            show_help
            ;;
        -m|--margin)
            MARGIN="$2"
            shift 2
            ;;
        -f|--font)
            FONT="$2"
            shift 2
            ;;
        -s|--size)
            FONT_SIZE="$2"
            shift 2
            ;;
        -t|--toc)
            TOC=true
            shift
            ;;
        -p|--paper)
            PAPER="$2"
            shift 2
            ;;
        --all)
            BATCH=true
            shift
            ;;
        *)
            if [[ -z "$INPUT" ]]; then
                INPUT="$1"
            elif [[ -z "$OUTPUT" ]]; then
                OUTPUT="$1"
            fi
            shift
            ;;
    esac
done

# Execute
if [[ "$BATCH" == true ]]; then
    convert_all
elif [[ -n "$INPUT" ]]; then
    echo -e "${BLUE}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${NC}"
    echo -e "${YELLOW}  MD-TO-PDF Converter${NC}"
    echo -e "${BLUE}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━${NC}"
    convert_file "$INPUT" "$OUTPUT"
else
    show_help
fi
