#!/usr/bin/env python3
"""
Talnt Document Generator Web App
- Resume Formatter: Format resumes with company branding
- RFQ Proposal Generator: Create proposals with pricing tables and formatted resumes
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path
from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename

# Add parent directory to path for importing format_resume
SCRIPT_DIR = Path(__file__).parent.resolve()
PARENT_DIR = SCRIPT_DIR.parent
sys.path.insert(0, str(PARENT_DIR))

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['UPLOAD_FOLDER'] = SCRIPT_DIR / 'uploads'

# Ensure folders exist
app.config['UPLOAD_FOLDER'].mkdir(exist_ok=True)
(PARENT_DIR / 'input').mkdir(exist_ok=True)
(PARENT_DIR / 'output').mkdir(exist_ok=True)
(SCRIPT_DIR / 'output').mkdir(exist_ok=True)

ALLOWED_EXTENSIONS = {'pdf', 'docx'}


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def format_resume_file(input_path, brand='dc'):
    """Run the resume formatter on a file and return the output path."""
    format_script = PARENT_DIR / "format_resume.py"
    input_folder = PARENT_DIR / "input"
    output_folder = PARENT_DIR / "output"

    # Clear input folder and copy new file
    for f in input_folder.glob("*"):
        if f.is_file():
            f.unlink()

    shutil.copy(input_path, input_folder / input_path.name)

    # Set brand environment variable for the formatter
    env = os.environ.copy()
    env['TALNT_BRAND'] = brand

    # Run the formatter
    result = subprocess.run(
        ['python3', str(format_script)],
        capture_output=True,
        text=True,
        cwd=str(PARENT_DIR),
        timeout=120,  # 2 minute timeout
        env=env
    )

    if result.returncode != 0:
        raise Exception(f"Resume formatting failed: {result.stderr}\n{result.stdout}")

    # Find the output file
    formatted_files = sorted(
        output_folder.glob("*_Formatted.docx"),
        key=lambda f: f.stat().st_mtime,
        reverse=True
    )

    if not formatted_files:
        raise Exception("Could not find formatted resume output")

    return formatted_files[0]


def calculate_totals(hourly_rate, duration_months, commitment_pct):
    """Calculate monthly and total costs."""
    # 173.33 hours per month (40 hrs/week * 52 weeks / 12 months)
    hours_per_month = 173.33
    monthly = hourly_rate * hours_per_month * (commitment_pct / 100)
    total = monthly * duration_months
    return monthly, total


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/format-resume', methods=['POST'])
def format_resume():
    """Format a resume and return the formatted DOCX."""
    try:
        if 'resume' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400

        file = request.files['resume']
        if not file or not file.filename:
            return jsonify({'error': 'No file selected'}), 400

        if not allowed_file(file.filename):
            return jsonify({'error': 'Invalid file type. Please upload PDF or DOCX'}), 400

        # Get brand selection
        brand = request.form.get('brand', 'dc')

        # Save uploaded file
        filename = secure_filename(file.filename)
        upload_path = app.config['UPLOAD_FOLDER'] / filename
        file.save(str(upload_path))

        # Format the resume with selected brand
        formatted_path = format_resume_file(upload_path, brand=brand)

        # Return the formatted file
        return send_file(
            str(formatted_path),
            as_attachment=True,
            download_name=formatted_path.name
        )

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/calculate', methods=['POST'])
def calculate():
    """Calculate totals for live preview."""
    try:
        data = request.json
        hourly_rate = float(data.get('hourly_rate', 0))
        duration = float(data.get('duration', 0))
        commitment = float(data.get('commitment', 100))

        monthly, total = calculate_totals(hourly_rate, duration, commitment)

        return jsonify({
            'monthly': f"${monthly:,.0f}",
            'total': f"${total:,.0f}"
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 400


@app.route('/generate', methods=['POST'])
def generate():
    """Generate the RFQ proposal document."""
    try:
        # Get form data
        staff_name = request.form.get('staff_name', '')
        position = request.form.get('position', '')
        duration = request.form.get('duration', '')
        hourly_rate = request.form.get('hourly_rate', '')
        commitment = request.form.get('commitment', '100')

        # Brand and dates
        brand = request.form.get('brand', 'dc')
        start_date = request.form.get('start_date', '')
        end_date = request.form.get('end_date', '')

        # Project details
        project_experience = request.form.get('project_experience', '')
        project_summary = request.form.get('project_summary', '')

        # Expenses
        expense_type = request.form.get('expense_type', '')
        expense_desc = request.form.get('expense_desc', '') or 'N/A'
        expense_monthly = request.form.get('expense_monthly', 'N/A')
        expense_total = request.form.get('expense_total', 'N/A')

        # Calculate totals
        try:
            hr = float(hourly_rate.replace('$', '').replace(',', ''))
            dur = float(duration.replace(' Months', '').replace(' months', ''))
            comm = float(commitment.replace('%', ''))
            staff_monthly, staff_total = calculate_totals(hr, dur, comm)
        except:
            hr, dur, comm = 0, 0, 100
            staff_monthly, staff_total = 0, 0

        # Parse expense values
        expense_monthly_val = 0
        expense_total_val = 0
        if expense_monthly and expense_monthly != 'N/A':
            try:
                expense_monthly_val = float(expense_monthly.replace('$', '').replace(',', ''))
                expense_total_val = expense_monthly_val * dur
            except:
                pass

        # Handle resume file
        resume_path = None
        formatted_resume_path = None

        if 'resume' in request.files:
            file = request.files['resume']
            if file and file.filename and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                resume_path = app.config['UPLOAD_FOLDER'] / filename
                file.save(str(resume_path))

                # Format the resume with selected brand
                try:
                    formatted_resume_path = format_resume_file(resume_path, brand=brand)
                except Exception as e:
                    print(f"Resume formatting error: {e}")

        # Generate the proposal DOCX
        from rfq.generate_rfq import generate_rfq_proposal

        output_path = SCRIPT_DIR / 'output' / f"RFQ_Proposal_{staff_name.replace(' ', '_')}.docx"
        output_path.parent.mkdir(exist_ok=True)

        # Calculate combined totals
        combined_monthly = staff_monthly + expense_monthly_val
        combined_total = staff_total + expense_total_val

        proposal_data = {
            'staff_name': staff_name,
            'position': position,
            'duration': duration,
            'hourly_rate': hourly_rate,
            'commitment': commitment,
            'staff_monthly': f"${staff_monthly:,.0f}",
            'staff_total': f"${staff_total:,.0f}",
            'project_experience': project_experience,
            'project_summary': project_summary,
            'expense_type': expense_type,
            'expense_desc': expense_desc if expense_desc else 'N/A',
            'expense_monthly': f"${expense_monthly_val:,.0f}" if expense_monthly_val > 0 else 'N/A',
            'expense_total': f"${expense_total_val:,.0f}" if expense_total_val > 0 else 'N/A',
            'combined_monthly': f"${combined_monthly:,.0f}",
            'combined_total': f"${combined_total:,.0f}",
            'formatted_resume_path': str(formatted_resume_path) if formatted_resume_path else None,
            'brand': brand,
            'start_date': start_date,
            'end_date': end_date
        }

        generate_rfq_proposal(proposal_data, str(output_path))

        return send_file(
            str(output_path),
            as_attachment=True,
            download_name=f"RFQ_Proposal_{staff_name.replace(' ', '_')}.docx"
        )

    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    # Get port from environment variable (for cloud deployment) or default to 5050
    port = int(os.environ.get('PORT', 5050))
    debug = os.environ.get('FLASK_DEBUG', 'true').lower() == 'true'

    print("=" * 60)
    print("Talnt Document Generator")
    print("=" * 60)
    print(f"\nOpen http://localhost:{port} in your browser")
    print("Press Ctrl+C to stop\n")

    # Bind to 0.0.0.0 for cloud deployment, localhost for local dev
    host = '0.0.0.0' if os.environ.get('RAILWAY_ENVIRONMENT') else '127.0.0.1'
    app.run(debug=debug, port=port, host=host)
