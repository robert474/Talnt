#!/usr/bin/env python3
"""
Data Center TALNT Resume Formatter
Automatically formats resumes to company standard with logo and Arial 9 font
"""

import os
import sys
import json
import subprocess
from pathlib import Path
import requests

# Import PDF/DOCX libraries
try:
    from pypdf import PdfReader
    import pdfplumber
except ImportError:
    print("Installing required packages...")
    subprocess.run([sys.executable, "-m", "pip", "install", "pypdf", "pdfplumber", "--break-system-packages"], check=True)
    from pypdf import PdfReader
    import pdfplumber

try:
    import docx
except ImportError:
    print("Installing python-docx...")
    subprocess.run([sys.executable, "-m", "pip", "install", "python-docx", "--break-system-packages"], check=True)
    import docx

# Configuration
SCRIPT_DIR = Path(__file__).parent
ASSETS_DIR = SCRIPT_DIR / "assets"
OUTPUT_DIR = SCRIPT_DIR / "output"
LOGO_ICON = ASSETS_DIR / "logo_icon.png"
LOGO_TEXT = ASSETS_DIR / "logo_text.png"

# API Key - set this or use environment variable ANTHROPIC_API_KEY
ANTHROPIC_API_KEY = os.environ.get('ANTHROPIC_API_KEY', '')

# Keywords to highlight
HIGHLIGHT_KEYWORDS = ["AWS", "Amazon", "Google", "Data Center", "Microsoft"]

def extract_text_from_pdf(pdf_path):
    """Extract text from PDF file"""
    text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                # Try layout mode first (better for multi-column resumes)
                page_text = page.extract_text(layout=True)
                if page_text:
                    # Clean up excessive whitespace from layout mode
                    # but preserve line structure
                    lines = page_text.split('\n')
                    cleaned_lines = []
                    for line in lines:
                        # Collapse multiple spaces to single space
                        line = ' '.join(line.split())
                        if line.strip():
                            cleaned_lines.append(line)
                    text += '\n'.join(cleaned_lines) + "\n"
    except Exception as e:
        print(f"Error extracting from PDF: {e}")
        # Fallback to pypdf
        try:
            reader = PdfReader(pdf_path)
            for page in reader.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
        except Exception as e2:
            print(f"Fallback also failed: {e2}")

    # Clean up null bytes and other problematic characters
    text = text.replace('\x00', '')
    # Remove other control characters except newlines and tabs
    text = ''.join(char for char in text if char == '\n' or char == '\t' or (ord(char) >= 32 and ord(char) < 127) or ord(char) > 127)

    return text

def extract_text_from_docx(docx_path):
    """Extract text from DOCX file"""
    try:
        doc = docx.Document(docx_path)
        text = "\n".join([paragraph.text for paragraph in doc.paragraphs])
        return text
    except Exception as e:
        print(f"Error extracting from DOCX: {e}")
        return ""

def simple_parse_resume(resume_text):
    """Enhanced parser that handles multiple resume formats"""

    import re

    data = {
        "name": "",
        "contact": {"location": "", "phone": "", "email": ""},
        "summary": "",
        "experience": [],
        "education": [],
        "certifications": [],
        "skills": ""
    }

    lines = resume_text.split('\n')

    # Extract name (usually first non-empty line that looks like a name)
    for line in lines[:10]:
        line = line.strip()
        if line and len(line) < 60 and len(line) > 3:
            # Skip lines that don't look like names
            skip_patterns = [
                'resume', 'cv', 'page', 'professional', 'summary', 'email', '@', 'phone',
                'scheduler', 'manager', 'engineer', 'director', 'specialist'
            ]
            # Skip addresses (start with numbers followed by letters - like "5013RollingwoodDr")
            looks_like_address = re.match(r'^\d+\s*[A-Za-z]', line)
            # Skip lines with city/state patterns like "Austin,TX" or phone numbers
            looks_like_contact = re.search(r',\s*[A-Z]{2}\s*\d{5}|^\(\d{3}\)', line)

            if not any(x in line.lower() for x in skip_patterns) and not looks_like_address and not looks_like_contact:
                data['name'] = line
                break

    # Extract contact info
    email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    phone_pattern = r'[\+\(]?[1-9][0-9 .\-\(\)]{8,}[0-9]'

    for line in lines[:30]:
        email_match = re.search(email_pattern, line)
        if email_match:
            data['contact']['email'] = email_match.group()

        phone_match = re.search(phone_pattern, line)
        if phone_match:
            data['contact']['phone'] = phone_match.group()

    text = resume_text

    # Date pattern to detect job header lines - handles multiple formats:
    # "Feb 2024 – Present", "Jan 2022 to present", "2023 - 2024", "2015 - Present"
    # "3/10/25 to Present", "03/05/20 to 2/20/25" (MM/DD/YY format)
    # "04/2001 to 02/2008" (MM/YYYY format)
    # "04/2012 to 4/20/18" (mixed MM/YYYY to M/DD/YY format)
    # For slash dates, require "to" or dash between dates (not just any number)
    date_pattern_month = r'(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+\d{4}\s*[-–—to\s]+\s*(?:Present|present|Current|current|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s*\d{4})'
    date_pattern_year = r'\d{4}\s*[-–—to\s]+\s*(?:Present|present|Current|current|\d{4})'
    date_pattern_slash_full = r'\d{1,2}/\d{1,2}/\d{2,4}\s*(?:to|[-–—])\s*(?:Present|present|Current|current|\d{1,2}/\d{1,2}/\d{2,4})'
    date_pattern_slash_my = r'\d{1,2}/\d{4}\s*(?:to|[-–—])\s*(?:Present|present|Current|current|\d{1,2}/\d{4})'
    # Mixed: MM/YYYY to M/DD/YY or vice versa - use a general slash pattern
    date_pattern_slash_any = r'\d{1,2}/\d{2,4}\s*(?:to|[-–—])\s*(?:Present|present|Current|current|\d{1,2}/\d{1,2}/\d{2,4})'
    date_pattern = f'{date_pattern_month}|{date_pattern_slash_full}|{date_pattern_slash_my}|{date_pattern_slash_any}|{date_pattern_year}'

    # Summary - extract until we hit a section header or skills/experience
    summary_patterns = [
        r'(?:PROFESSIONAL\s+SUMMARY|SUMMARY|PROFILE|OBJECTIVE)\s*\n+(.*?)(?=\n\s*(?:TECHNICAL|SKILLS|EDUCATION|EXPERIENCE|EMPLOYMENT|WORK|CORE|PROFESSIONAL\s+EXPERIENCE|SELECTED))',
        r'(?:PROFESSIONAL\s+SUMMARY|SUMMARY)\s*\n+(.*?)(?=\n[A-Z][a-z]+\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec))',
        # Format: Title headline followed by paragraph, then CORE COMPETENCIES or Professional Experience
        r'(?:VICE\s+PRESIDENT|DIRECTOR|MANAGER|EXECUTIVE|ENGINEER|SPECIALIST|CONSULTANT|ANALYST|SENIOR\s+\w+\s+\w+\s+\w+)[^\n]*\n+(.*?)(?=\n\s*(?:CORE\s+COMPETENCIES|SKILLS|TECHNICAL|AREAS\s+OF|PROFESSIONAL\s+EXPERIENCE|EXPERIENCE:|EMPLOYMENT))',
    ]

    for pattern in summary_patterns:
        summary_match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
        if summary_match:
            summary = summary_match.group(1).strip()
            # Clean up - join lines and remove bullets
            summary_lines = []
            for line in summary.split('\n'):
                line = line.strip()
                if line and not line.startswith('•') and not line.startswith('-'):
                    summary_lines.append(line)
            summary = ' '.join(summary_lines)
            summary = re.sub(r'\s+', ' ', summary)
            if len(summary) > 50:
                data['summary'] = summary
                break

    # Skills - look for skills/tools section (usually at end or after summary)
    skills_patterns = [
        r'(?:CORE\s+COMPETENCIES|CORE\s+SKILLS|TECHNICAL\s+TOOLS?|SKILLS|TOOLS)\s*\n+(.*?)(?=\n\s*(?:EDUCATION|CERTIFICATION|PROFESSIONAL\s+EXPERIENCE|EXPERIENCE|EMPLOYMENT))',
        r'(?:CORE\s+SKILLS|TECHNICAL\s+TOOLS?|SKILLS|TOOLS)\s*\n+(.*?)(?=\n\s*(?:EDUCATION|CERTIFICATION)|\Z)',
        r'(?:SKILLS)\s*\n+(.*?)(?=\n\s*\$|\Z)'  # For resumes with $ amounts in job titles
    ]

    for pattern in skills_patterns:
        skills_match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
        if skills_match:
            skills_text = skills_match.group(1).strip()
            # Clean up bullets and newlines
            skills_text = re.sub(r'[•\-\*]\s*', '', skills_text)
            skills_text = re.sub(r'\s*\n\s*', ', ', skills_text)
            skills_text = re.sub(r'\s+', ' ', skills_text)
            skills_text = re.sub(r',\s*,', ',', skills_text)  # Remove double commas
            if len(skills_text) > 10:
                data['skills'] = skills_text
                break

    # Education - parse each degree entry with school on separate line
    edu_patterns = [
        r'EDUCATION(?:\s*(?:&|AND)\s*CERTIFICATIONS?)?\s*\n+(.*?)(?=\n\s*(?:CERTIFICATION|SKILLS|EXPERIENCE|$))',
        r'EDUCATION\s*\n+(.*?)(?=\Z)'
    ]

    for pattern in edu_patterns:
        edu_match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
        if edu_match:
            edu_text = edu_match.group(1).strip()
            edu_lines = [l.strip() for l in edu_text.split('\n') if l.strip()]

            i = 0
            while i < len(edu_lines):
                line = edu_lines[i]
                # Skip certification lines
                if any(x in line.lower() for x in ['pmp', 'safe', 'scrum', 'certified', 'certification']):
                    i += 1
                    continue

                # Check if this is a school name (University, College, etc.)
                if any(x in line for x in ['University', 'College', 'Institute', 'Polytechnic', 'School']):
                    school = line
                    degree = ""
                    # Next line might be the degree
                    if i + 1 < len(edu_lines):
                        next_line = edu_lines[i + 1]
                        if any(x in next_line for x in ['Master', 'Bachelor', 'MSc', 'BSc', 'MBA', 'HND', 'M.S', 'B.S', 'PhD', 'Diploma']):
                            degree = next_line
                            i += 1
                    data['education'].append({"degree": degree, "school": school, "year": ""})
                # Check if this is a degree line
                elif any(x in line for x in ['Master', 'Bachelor', 'MSc', 'BSc', 'MBA', 'HND', 'M.S.', 'B.S.', 'PhD', 'Diploma', 'M.Sc', 'B.Sc']):
                    degree = line
                    school = ""
                    # Next line might be school
                    if i + 1 < len(edu_lines):
                        next_line = edu_lines[i + 1]
                        if any(x in next_line for x in ['University', 'College', 'Institute', 'Polytechnic']):
                            school = next_line
                            i += 1
                    # Or check if degree line contains school (separated by dash/comma)
                    if not school and ('—' in degree or '–' in degree or ' - ' in degree):
                        parts = re.split(r'\s*[—–-]\s*', degree, 1)
                        if len(parts) == 2:
                            degree = parts[0].strip()
                            school = parts[1].strip()
                    data['education'].append({"degree": degree, "school": school, "year": ""})
                i += 1
            break

    # Employment History - detect multiple formats
    # Some resumes have no header, jobs start after Skills section with $ or company names
    exp_patterns = [
        r'(?:PROFESSIONAL\s+EXPERIENCE|EMPLOYMENT\s+HISTORY|WORK\s+HISTORY|WORK\s+EXPERIENCE|EXPERIENCE)\s*\n+(.*?)(?=\n\s*(?:EDUCATION|SELECTED|CORE\s+SKILLS|CERTIFICATIONS?)|\Z)',
    ]

    # First try standard patterns
    exp_text = None
    for pattern in exp_patterns:
        exp_match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
        if exp_match:
            exp_text = exp_match.group(1)
            break

    # If no experience section found, look for lines starting with $ (budget amount format)
    if not exp_text:
        # Find the first line that starts with $ and has a date
        dollar_job_match = re.search(r'(\$[\d.,]+\s*(?:Billion|Million|B|M)?.*?' + date_pattern + r'.*?)(?=\n\s*EDUCATION|\Z)', text, re.DOTALL | re.IGNORECASE)
        if dollar_job_match:
            exp_text = dollar_job_match.group(1)

    if exp_text:
        jobs = []
        current_job = None
        exp_lines = exp_text.split('\n')

        i = 0
        while i < len(exp_lines):
            line = exp_lines[i].strip()

            if not line:
                i += 1
                continue

            # Check for date range in line (indicates job header)
            has_date = re.search(date_pattern, line, re.IGNORECASE)

            # Format 1: "Title — Company | Dates" (em dash)
            if '—' in line and has_date:
                # Save previous job
                if current_job:
                    jobs.append(current_job)

                parts = re.split(r'\s*[—|]\s*', line)
                title = parts[0].strip() if len(parts) > 0 else ""
                company = parts[1].strip() if len(parts) > 1 else ""
                dates = parts[2].strip() if len(parts) > 2 else ""

                current_job = {
                    "company": company,
                    "title": title,
                    "location": "",
                    "dates": dates,
                    "project_details": "",
                    "bullets": []
                }

            # Format 1b: "COMPANY | Title" on line 1, "Dates | Location" on line 2
            # Check if this line has pipe but no date, and next line has date
            elif '|' in line and not has_date and not line.startswith('•'):
                if i + 1 < len(exp_lines):
                    next_line = exp_lines[i + 1].strip()
                    next_has_date = re.search(date_pattern, next_line, re.IGNORECASE)
                    if next_has_date:
                        # Save previous job
                        if current_job:
                            jobs.append(current_job)

                        # Parse "COMPANY | Title" from current line
                        parts = re.split(r'\s*\|\s*', line)
                        company = parts[0].strip() if len(parts) > 0 else ""
                        title = parts[1].strip() if len(parts) > 1 else ""

                        # Parse "Dates | Location" from next line
                        dates_match = re.search(date_pattern, next_line, re.IGNORECASE)
                        dates = dates_match.group(0).strip() if dates_match else ""
                        # Get location from remaining parts after removing date
                        location_parts = re.sub(date_pattern, '', next_line, flags=re.IGNORECASE).strip()
                        location_parts = re.split(r'\s*\|\s*', location_parts)
                        location = ' '.join([p.strip() for p in location_parts if p.strip()])

                        current_job = {
                            "company": company,
                            "title": title,
                            "location": location,
                            "dates": dates,
                            "project_details": "",
                            "bullets": []
                        }
                        i += 1  # Skip the next line since we processed it

            # Format 2: Company name on its own line, then "Title | Location | Dates"
            # Only matches if this line could be a company AND next line has a date
            # Company names typically: start with capital, are short, don't end in punctuation like '.' or ','
            elif not has_date and not line.startswith('•') and not line.startswith('(') and len(line) < 80:
                is_company_line = False
                # Skip if line looks like bullet continuation (ends with '.', starts lowercase, contains 'and', etc.)
                looks_like_continuation = (
                    line.endswith('.') or
                    line.endswith(',') or
                    (line[0].islower() if line else False) or
                    line.startswith('and ') or
                    line.startswith('or ')
                )

                # Check if next line has the date
                if not looks_like_continuation and i + 1 < len(exp_lines):
                    next_line = exp_lines[i + 1].strip()
                    next_has_date = re.search(date_pattern, next_line, re.IGNORECASE)

                    # Check if there's a pipe | OR if there's title|location before the date
                    # (not just a dash within the date itself)
                    has_separator_before_date = False
                    if next_has_date:
                        # Get the part before the date
                        before_date = re.sub(date_pattern, '', next_line, flags=re.IGNORECASE).strip()
                        # Check if it has separators (|, –, —) that would indicate title | location format
                        has_separator_before_date = '|' in before_date or ' – ' in before_date or ' — ' in before_date

                    # Only use this format when there's a PIPE separator (not dash)
                    # Dash is used for "Company – Address" format, not "Title | Location"
                    has_pipe_separator = '|' in before_date if next_has_date else False

                    if next_has_date and has_pipe_separator:
                        is_company_line = True
                        # This line is company, next line is title|location|dates
                        if current_job:
                            jobs.append(current_job)

                        company = line
                        # First extract the date from the line
                        dates_match = re.search(date_pattern, next_line, re.IGNORECASE)
                        dates = dates_match.group(0).strip() if dates_match else ""

                        # Remove date from line for further parsing
                        line_without_date = re.sub(date_pattern, '', next_line, flags=re.IGNORECASE).strip()
                        line_without_date = line_without_date.rstrip('|').strip()

                        # Split ONLY on pipe character to preserve dashes in titles
                        parts = re.split(r'\s*\|\s*', line_without_date)
                        parts = [p.strip() for p in parts if p.strip()]

                        title = parts[0] if len(parts) > 0 else ""
                        location = ""
                        # Remaining parts are location
                        if len(parts) > 1:
                            location = ", ".join(parts[1:])

                        current_job = {
                            "company": company,
                            "title": title,
                            "location": location,
                            "dates": dates,
                            "project_details": "",
                            "bullets": []
                        }
                        i += 1  # Skip the next line since we processed it

                # If not a company line, could be title followed by company+date on next line
                if not is_company_line and not looks_like_continuation:
                    # Check if THIS line is a title and NEXT line is company+date
                    if i + 1 < len(exp_lines):
                        next_line = exp_lines[i + 1].strip()
                        next_has_date = re.search(date_pattern, next_line, re.IGNORECASE)
                        # Next line has date (Company – Address Date format or Company Date format)
                        if next_has_date and '|' not in next_line:
                            # This line is title, next is company+location+date
                            if current_job:
                                jobs.append(current_job)

                            title = line
                            dates_match = re.search(date_pattern, next_line, re.IGNORECASE)
                            dates = dates_match.group(0).strip() if dates_match else ""

                            # Remove dates from line to get company info
                            company_part = re.sub(date_pattern, '', next_line, flags=re.IGNORECASE).strip().rstrip(',').strip()

                            # Check if company has location separated by dash (Company – Address format)
                            # e.g., "AGL Fuel Transportations – 12215 Telegraph Rd Santa Fe Springs, CA 90670"
                            company = company_part
                            location = ""
                            if ' – ' in company_part or ' - ' in company_part:
                                dash_parts = re.split(r'\s*[–-]\s*', company_part, 1)
                                if len(dash_parts) == 2:
                                    company = dash_parts[0].strip()
                                    location = dash_parts[1].strip()

                            current_job = {
                                "company": company,
                                "title": title,
                                "location": location,
                                "dates": dates,
                                "project_details": "",
                                "bullets": []
                            }
                            i += 1  # Skip the next line since we processed it
                            is_company_line = True  # Mark as handled

                # If not a company/title line, check if it's a bullet continuation
                if not is_company_line and current_job and current_job.get('bullets'):
                    if not any(x in line for x in ['University', 'College', 'EDUCATION', 'SKILLS', 'CERTIFICATION']):
                        if not line.isupper():
                            current_job['bullets'][-1] += ' ' + line

            # Format 3: "$X.X Billion Company Project, Location, Dates" then "Title"
            elif line.startswith('$') and has_date:
                if current_job:
                    jobs.append(current_job)

                # Extract company info from this line
                company_match = re.match(r'\$[\d.]+\s*(?:Billion|Million|B|M)?\s*(.*?),?\s*(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)', line, re.IGNORECASE)
                if company_match:
                    company = company_match.group(1).strip().rstrip(',')
                else:
                    company = line

                dates_match = re.search(date_pattern, line, re.IGNORECASE)
                dates = dates_match.group(0) if dates_match else ""

                # Next line should be title
                title = ""
                if i + 1 < len(exp_lines):
                    next_line = exp_lines[i + 1].strip()
                    if not next_line.startswith('•') and not re.search(date_pattern, next_line):
                        title = next_line
                        i += 1

                current_job = {
                    "company": company,
                    "title": title,
                    "location": "",
                    "dates": dates,
                    "project_details": "",
                    "bullets": []
                }

            # Format 4: "Title Dates" or "Company Dates" on one line
            # Need to determine if text before date is title or company
            elif has_date and not line.startswith('•'):
                if current_job:
                    jobs.append(current_job)

                dates_match = re.search(date_pattern, line, re.IGNORECASE)
                dates = dates_match.group(0) if dates_match else ""
                text_before_date = re.sub(date_pattern, '', line, flags=re.IGNORECASE).strip().rstrip(',').strip()

                title = ""
                company = ""
                location = ""

                # Check next line to determine if text_before_date is title or company
                # If next line looks like a company (has Inc., LLC, Corp, Company, or dash with address)
                # then text_before_date is the title
                next_line = ""
                if i + 1 < len(exp_lines):
                    next_line = exp_lines[i + 1].strip()

                company_indicators = ['Inc.', 'Inc,', 'LLC', 'Corp', 'Company', 'Ltd', 'L.L.C', 'Corporation', 'Center', 'Administrators', 'Airlines']
                # Check if next line looks like a company name with address
                # Patterns: "Company – City, State" or "Company – Address" or company indicators
                next_looks_like_company = (
                    next_line and
                    not next_line.startswith('•') and
                    not re.search(date_pattern, next_line, re.IGNORECASE) and
                    (any(ind in next_line for ind in company_indicators) or
                     (' – ' in next_line and re.search(r',\s*[A-Z]{2}', next_line)) or  # Has dash and "City, ST" pattern
                     (' - ' in next_line and re.search(r',\s*[A-Z]{2}', next_line)) or
                     (' – ' in next_line and re.search(r'\d{5}', next_line)))  # Has dash and ZIP code
                )

                if next_looks_like_company:
                    # Current line has title+date, next line has company+address
                    title = text_before_date
                    company_line = next_line
                    # Parse company and location from next line
                    if ' – ' in company_line or ' - ' in company_line:
                        dash_parts = re.split(r'\s*[–-]\s*', company_line, maxsplit=1)
                        if len(dash_parts) == 2:
                            company = dash_parts[0].strip()
                            location = dash_parts[1].strip()
                        else:
                            company = company_line
                    else:
                        company = company_line
                    i += 1  # Skip next line since we consumed it
                else:
                    # Current line has company+date (traditional format)
                    company = text_before_date
                    # Check if previous line was a title
                    if i > 0:
                        prev_line = exp_lines[i - 1].strip()
                        prev_has_date = re.search(date_pattern, prev_line, re.IGNORECASE)
                        if (prev_line and not prev_line.startswith('•') and
                            not prev_has_date and len(prev_line) < 60 and
                            not prev_line.endswith('.') and not prev_line.endswith(',')):
                            title = prev_line
                            if jobs and jobs[-1].get('bullets'):
                                last_bullet = jobs[-1]['bullets'][-1] if jobs[-1]['bullets'] else ""
                                if title in last_bullet:
                                    title = ""
                                elif last_bullet == title:
                                    jobs[-1]['bullets'].pop()

                current_job = {
                    "company": company,
                    "title": title,
                    "location": location,
                    "dates": dates,
                    "project_details": "",
                    "bullets": []
                }

            # Bullet points - with or without bullet markers
            elif current_job and (line.startswith('•') or line.startswith('-') or line.startswith('*') or (line.startswith('(') and 'cid:' in line)):
                # Handle (cid:127) bullet format
                bullet = re.sub(r'^[\•\-\*]\s*', '', line)
                bullet = re.sub(r'^\(cid:\d+\)\s*', '', bullet)
                if bullet and len(bullet) > 10:
                    current_job['bullets'].append(bullet)

            # Lines without bullet markers but part of job description
            # These are either intro paragraphs or unmarked bullets
            elif current_job and line and not has_date:
                # Check it's not a new section header
                is_section_header = any(x in line for x in ['University', 'College', 'EDUCATION', 'SKILLS', 'CERTIFICATION', 'CORE COMPETENCIES'])
                is_all_caps = line.isupper() and len(line) > 3
                has_pipe = '|' in line  # Likely another job header

                if not is_section_header and not is_all_caps and not has_pipe:
                    # This is content for the current job
                    if len(line) > 20:  # Only substantial lines
                        current_job['bullets'].append(line)

            i += 1

        # Don't forget the last job
        if current_job:
            jobs.append(current_job)

        data['experience'] = jobs

    # Certifications - check both dedicated section and education section
    cert_patterns = [
        r'CERTIFICATION[S]?\s*\n+(.*?)(?=\n\s*(?:EDUCATION|EXPERIENCE|SKILLS|TECHNICAL)|\Z)',
        r'(?:EDUCATION\s*(?:&|AND)\s*CERTIFICATIONS?)\s*\n+.*?((?:PMP|PMI|SAFe|OSHA|Certified).*?)(?=\Z)'
    ]

    for pattern in cert_patterns:
        cert_match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
        if cert_match:
            cert_text = cert_match.group(1)
            certs = []
            for line in cert_text.split('\n'):
                line = line.strip()
                line = re.sub(r'^[\•\-\*]\s*', '', line)
                if line and len(line) > 3:
                    # Don't include degree lines
                    if not any(x in line for x in ['University', 'College', 'Bachelor', 'Master', 'MSc', 'BSc', 'MBA']):
                        certs.append(line)
            if certs:
                data['certifications'] = certs[:10]
                break

    return data

def validate_and_clean_data(data):
    """Validate and clean parsed data to remove duplications and errors"""
    
    # Remove duplicate experience entries (check by company+title)
    if data.get('experience'):
        seen = set()
        unique_exp = []
        for exp in data['experience']:
            key = f"{exp.get('company', '')}|{exp.get('title', '')}"
            if key not in seen and exp.get('company'):
                seen.add(key)
                unique_exp.append(exp)
        data['experience'] = unique_exp
    
    # Clean up summary - remove if it contains too many skill keywords
    if data.get('summary'):
        summary = data['summary']
        skill_indicators = ['primavera', 'microsoft project', 'power bi', 'excel', 'oracle', 'sap']
        skill_count = sum(1 for indicator in skill_indicators if indicator in summary.lower())
        
        # If summary has too many skills, extract just the first real paragraph
        if skill_count > 3:
            # Split by double newlines or bullets
            paragraphs = [p.strip() for p in summary.split('\n') if p.strip() and not p.strip().startswith('•')]
            if paragraphs:
                # Find first substantial paragraph (>100 chars)
                for para in paragraphs:
                    if len(para) > 100 and skill_count_in_text(para) < 3:
                        data['summary'] = para
                        break
    
    # Ensure skills is a string
    if data.get('skills'):
        if isinstance(data['skills'], list):
            data['skills'] = ', '.join(data['skills'])
    
    # Clean up education - remove duplicates
    if data.get('education'):
        seen_edu = set()
        unique_edu = []
        for edu in data['education']:
            key = f"{edu.get('degree', '')}|{edu.get('school', '')}"
            if key not in seen_edu and edu.get('degree'):
                seen_edu.add(key)
                unique_edu.append(edu)
        data['education'] = unique_edu
    
    return data

def skill_count_in_text(text):
    """Count how many skill keywords appear in text"""
    skill_indicators = ['primavera', 'microsoft project', 'power bi', 'excel', 'oracle', 'sap']
    return sum(1 for indicator in skill_indicators if indicator in text.lower())

def parse_resume_with_claude(resume_text):
    """Use Claude API to parse resume into structured format"""
    
    # Check if API key is available
    if not ANTHROPIC_API_KEY:
        print("Note: Claude API key not set. Using simple parser.")
        print("For better results, set ANTHROPIC_API_KEY environment variable.\n")
        return simple_parse_resume(resume_text)
    
    prompt = f"""You are parsing a resume to reformat it into a standardized template. Extract ALL information and structure it EXACTLY as specified below.

CRITICAL RULES:
1. Extract EVERY job/position as a separate entry in experience array
2. Do NOT duplicate any content
3. Do NOT merge sections together
4. Preserve ALL bullet points for each job
5. Extract education degrees separately (one object per degree)
6. Skills/tools should be a single comma-separated string

Return ONLY valid JSON with this EXACT structure:

{{
  "name": "Full Name",
  "contact": {{
    "location": "City, State or Country",
    "phone": "phone number",
    "email": "email address"
  }},
  "summary": "Complete professional summary paragraph. Do NOT include technical skills here.",
  "experience": [
    {{
      "company": "Company Name",
      "title": "Job Title",
      "location": "City, State",
      "dates": "Month Year - Month Year",
      "project_details": "Only if there's a 'Project:' or 'Projects:' line, otherwise empty string",
      "bullets": ["First responsibility", "Second responsibility", "etc - include ALL bullets for this job"]
    }},
    {{
      "company": "Next Company Name",
      "title": "Next Job Title",
      "location": "City, State",
      "dates": "Month Year - Month Year",
      "project_details": "",
      "bullets": ["First responsibility", "Second responsibility", "etc"]
    }}
  ],
  "education": [
    {{
      "degree": "Degree Name and Major",
      "school": "University/School Name",
      "year": "Year or empty string"
    }}
  ],
  "certifications": ["Certification 1", "Certification 2", "etc"],
  "skills": "Primavera P6, MS Project, Power BI, Excel, etc - all tools/software comma-separated"
}}

PARSING INSTRUCTIONS:
- Name: Extract the candidate's full name (usually at top)
- Summary: Extract the professional summary/objective paragraph ONLY. Do not include skills list.
- Experience: Create ONE object per job position. Include company, title, location, dates, and ALL bullets for that specific job.
- Project details: Only populate if there's an explicit "Project:" or "Projects:" line under a job
- Education: Create one object per degree (MSc, BSc, HND, etc). Format as "Degree" and "School" separately
- Certifications: Extract all certifications as array items
- Skills: Combine all technical tools/software into one comma-separated string

DO NOT:
- Duplicate any experience entries
- Merge technical skills into the summary
- Include skills in multiple places
- Skip any jobs/positions
- Combine multiple jobs into one entry

Resume text:
{resume_text}

Return ONLY the JSON, no markdown, no explanation, no other text."""

    try:
        response = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "Content-Type": "application/json",
                "x-api-key": ANTHROPIC_API_KEY,
                "anthropic-version": "2023-06-01"
            },
            json={
                "model": "claude-sonnet-4-20250514",
                "max_tokens": 8000,
                "messages": [
                    {"role": "user", "content": prompt}
                ]
            }
        )
        
        if response.status_code == 200:
            result = response.json()
            content = result['content'][0]['text']
            
            # Extract JSON from response (in case there's extra text)
            json_start = content.find('{')
            json_end = content.rfind('}') + 1
            if json_start != -1 and json_end > json_start:
                content = content[json_start:json_end]
            
            parsed_data = json.loads(content)
            
            # Validate and clean the data
            parsed_data = validate_and_clean_data(parsed_data)
            
            return parsed_data
        else:
            print(f"API Error: {response.status_code} - falling back to simple parser")
            return simple_parse_resume(resume_text)
            
    except Exception as e:
        print(f"Error with Claude API, using simple parser: {e}\n")
        return simple_parse_resume(resume_text)

def generate_formatted_docx(parsed_data, output_path):
    """Generate formatted DOCX with company template"""

    # We'll use the existing Node.js script for docx generation
    # Create a temporary JSON file with the parsed data
    json_path = SCRIPT_DIR / "temp_resume_data.json"
    with open(json_path, 'w') as f:
        json.dump(parsed_data, f, indent=2)

    # Use the existing generate_docx.js script (which supports brand selection)
    node_script = SCRIPT_DIR / "generate_docx.js"

    # Run Node.js script - the script reads TALNT_BRAND from environment
    try:
        result = subprocess.run(
            ['node', str(node_script), str(json_path), str(output_path)],
            capture_output=True,
            text=True,
            check=True
        )
        print(result.stdout)
        return True
    except subprocess.CalledProcessError as e:
        print(f"Error generating DOCX: {e}")
        print(e.stderr)
        return False
    finally:
        # Clean up temp file
        if json_path.exists():
            json_path.unlink()

def convert_to_pdf(docx_path, pdf_path):
    """Convert DOCX to PDF using LibreOffice"""
    try:
        subprocess.run([
            'soffice', '--headless', '--convert-to', 'pdf',
            str(docx_path), '--outdir', str(pdf_path.parent)
        ], check=True, capture_output=True)
        return True
    except FileNotFoundError:
        print("LibreOffice not installed - skipping PDF conversion")
        return False
    except subprocess.CalledProcessError as e:
        print(f"Error converting to PDF: {e}")
        return False

def format_resume(input_path):
    """Main function to format a resume"""
    
    input_path = Path(input_path)
    
    if not input_path.exists():
        print(f"Error: File not found: {input_path}")
        return False
    
    print(f"\n{'='*60}")
    print(f"Processing: {input_path.name}")
    print(f"{'='*60}\n")
    
    # Step 1: Extract text
    print("Step 1: Extracting text from resume...")
    if input_path.suffix.lower() == '.pdf':
        text = extract_text_from_pdf(input_path)
    elif input_path.suffix.lower() in ['.docx', '.doc']:
        text = extract_text_from_docx(input_path)
    else:
        print(f"Error: Unsupported file format: {input_path.suffix}")
        return False
    
    if not text.strip():
        print("Error: Could not extract text from resume")
        return False
    
    print(f"✓ Extracted {len(text)} characters\n")

    # Step 2: Parse resume
    # Always use Claude API when available - handles any format intelligently
    if ANTHROPIC_API_KEY:
        print("Step 2: Parsing resume with Claude AI...")
        parsed_data = parse_resume_with_claude(text)
    else:
        print("Step 2: Parsing resume...")
        print("⚠ Warning: No API key. Using simple parser (may have format issues).")
        print("  Set ANTHROPIC_API_KEY for better results.\n")
        parsed_data = simple_parse_resume(text)
    
    if not parsed_data:
        print("Error: Could not parse resume")
        return False
    
    print(f"✓ Parsed resume structure\n")
    
    # Step 3: Generate formatted DOCX
    print("Step 3: Generating formatted DOCX...")
    
    # Create output filename
    name = parsed_data.get('name', input_path.stem).replace(' ', '_')
    output_docx = OUTPUT_DIR / f"{name}_Formatted.docx"
    
    if not generate_formatted_docx(parsed_data, output_docx):
        print("Error: Could not generate formatted DOCX")
        return False
    
    print(f"✓ Created: {output_docx}\n")
    
    # Step 4: Convert to PDF
    print("Step 4: Converting to PDF...")
    output_pdf = output_docx.with_suffix('.pdf')
    
    if convert_to_pdf(output_docx, output_pdf):
        print(f"✓ Created: {output_pdf}\n")
    else:
        print("Warning: Could not convert to PDF, but DOCX is available\n")
    
    print(f"{'='*60}")
    print(f"✓ SUCCESS! Resume formatted and saved to output folder")
    print(f"{'='*60}\n")
    
    return True

def batch_process():
    """Process all resumes in input folder"""
    input_dir = SCRIPT_DIR / "input"
    
    resumes = list(input_dir.glob("*.pdf")) + list(input_dir.glob("*.docx"))
    
    if not resumes:
        print("No resumes found in input folder")
        print("Please place PDF or DOCX files in the 'input' folder")
        return
    
    print(f"\nFound {len(resumes)} resume(s) to process\n")
    
    success_count = 0
    for resume in resumes:
        if format_resume(resume):
            success_count += 1
        print()
    
    print(f"\n{'='*60}")
    print(f"Batch processing complete!")
    print(f"Successfully formatted {success_count}/{len(resumes)} resumes")
    print(f"{'='*60}\n")

def main():
    """Main entry point"""
    
    # Ensure output directory exists
    OUTPUT_DIR.mkdir(exist_ok=True)
    
    if len(sys.argv) > 1:
        # Process specific file
        input_file = sys.argv[1]
        format_resume(input_file)
    else:
        # Batch process input folder
        batch_process()

if __name__ == "__main__":
    main()
