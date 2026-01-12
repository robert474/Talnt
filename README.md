# Data Center TALNT Resume Formatter

Automatically format resumes to company standard with logo and Arial 9 font.

**Time savings: 20 minutes → 30 seconds per resume**

## Quick Start

### Option 1: Batch Process Multiple Resumes

1. Drop all resumes (PDF or DOCX) into the `input` folder
2. Run: `python format_resume.py`
3. Get formatted resumes from the `output` folder

### Option 2: Format Single Resume

```bash
python format_resume.py path/to/resume.pdf
```

## Setup (Optional - For Best Results)

The tool works out of the box with a simple parser. For **best results** with complex resumes:

1. Get an Anthropic API key from https://console.anthropic.com/
2. Set it as an environment variable:

**Mac/Linux:**
```bash
export ANTHROPIC_API_KEY="your-api-key-here"
```

**Windows (Command Prompt):**
```cmd
set ANTHROPIC_API_KEY=your-api-key-here
```

**Windows (PowerShell):**
```powershell
$env:ANTHROPIC_API_KEY="your-api-key-here"
```

Or edit `format_resume.py` and add your key on line 27:
```python
ANTHROPIC_API_KEY = 'your-api-key-here'
```

## What It Does

✓ Extracts text from any PDF or Word resume
✓ Uses Claude AI to intelligently parse sections (name, experience, education, etc.)
✓ Reformats to company standard:
  - Arial 9 font throughout
  - Data Center TALNT logo in upper left
  - Consistent section headers and bullet formatting
  - Auto-highlights: AWS, Amazon, Google, Data Center, Microsoft
✓ Outputs both DOCX and PDF formats

## Folder Structure

```
Talnt/
├── format_resume.py       # Main script
├── input/                 # Drop resumes here
├── output/                # Formatted resumes appear here
└── assets/                # Logo files (don't modify)
    ├── logo_icon.png
    └── logo_text.png
```

## Requirements

**Already Installed:**
- Python 3
- Node.js (for docx generation)
- LibreOffice (for PDF conversion)

**Auto-installs on first run:**
- pypdf
- pdfplumber
- python-docx

## Examples

**Batch process:**
```bash
# Put resumes in input folder, then:
python format_resume.py

# Output:
# Processing: John_Doe_Resume.pdf
# ✓ Extracted text
# ✓ Parsed with Claude
# ✓ Created formatted DOCX
# ✓ Created formatted PDF
```

**Single file:**
```bash
python format_resume.py ~/Downloads/candidate_resume.pdf
```

## Troubleshooting

**"No module named 'pypdf'"**
- Script will auto-install on first run

**"Could not extract text"**
- Resume might be scanned image - try converting to text first

**"API Error"**
- Claude API might need configuration
- Check internet connection

## Customization

To modify formatting:
1. Edit the `generate_docx.js` section in `format_resume.py`
2. Adjust `fontSize` variable (18 = 9pt, 20 = 10pt)
3. Modify spacing values for tighter/looser layout

## Support

Questions? Issues? 
- Check that logos are in `assets/` folder
- Verify input files are readable PDF or DOCX format
- Test with a simple resume first

---

**Created for Data Center TALNT**
Streamlining resume formatting since 2025
