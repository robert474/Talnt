#!/bin/bash
# Quick launcher for resume formatter

cd "$(dirname "$0")"

echo "===================================="
echo "Data Center TALNT Resume Formatter"
echo "===================================="
echo ""

# Check if input folder has files
if [ "$(ls -A input 2>/dev/null)" ]; then
    echo "Found resumes in input folder. Processing..."
    python3 format_resume.py
else
    echo "No resumes found in input folder."
    echo ""
    echo "Usage:"
    echo "  1. Place PDF or DOCX resumes in the 'input' folder"
    echo "  2. Run this script again"
    echo "  3. Get formatted resumes from 'output' folder"
    echo ""
    echo "Or process a single file:"
    echo "  python format_resume.py /path/to/resume.pdf"
fi
