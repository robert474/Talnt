# QUICK START GUIDE - Data Center TALNT Resume Formatter

## 30-Second Setup

1. **Download** the Talnt folder to your Desktop
2. **Open Terminal** (Mac) or **Command Prompt** (Windows)
3. **Navigate** to the folder:
   ```
   cd Desktop/Talnt
   ```

## Daily Use (30 seconds per resume)

### Method 1: Drag and Drop (Easiest)

1. Drop all PDFs/Word docs into the `input` folder
2. Double-click `run.sh` (Mac) or run `python format_resume.py` (Windows)
3. Get formatted resumes from `output` folder

### Method 2: Single File

```bash
python format_resume.py /path/to/resume.pdf
```

## What You Get

✅ Arial 9 font
✅ Data Center TALNT logo
✅ Consistent formatting
✅ Bold AWS, Google, Data Center, Microsoft
✅ Both DOCX and PDF output

## Folder Structure

```
Talnt/
├── input/     ← Drop resumes here
├── output/    ← Get formatted resumes here
└── assets/    ← Logo files (don't touch)
```

## Troubleshooting

**"python: command not found"**
- Try `python3` instead of `python`

**Resume looks wrong**
- For better results, get free API key from https://console.anthropic.com/
- Set it: `export ANTHROPIC_API_KEY="your-key"`

**Logo missing**
- Make sure `assets/` folder has logo files

## That's It!

Drop resumes → Run script → Get formatted resumes

**Questions?** Check README.md for full documentation
