#!/usr/bin/env python3
"""
Generate RFQ Proposal DOCX documents.
Uses the template structure from existing RFQ proposals.
"""

import subprocess
import json
import tempfile
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent.resolve()
PARENT_DIR = SCRIPT_DIR.parent


def generate_rfq_proposal(data, output_path):
    """
    Generate an RFQ proposal DOCX file.

    Args:
        data: Dictionary containing:
            - staff_name, position, duration, hourly_rate, commitment
            - staff_monthly, staff_total
            - project_experience, project_summary
            - expense_desc, expense_monthly, expense_total
            - combined_monthly, combined_total
            - formatted_resume_path (optional)
        output_path: Path to save the output DOCX
    """
    # Write data to temp JSON file for Node.js script
    with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
        json.dump(data, f)
        data_path = f.name

    try:
        # Run Node.js script to generate DOCX
        node_script = SCRIPT_DIR / 'generate_rfq_docx.js'
        result = subprocess.run(
            ['node', str(node_script), data_path, str(output_path)],
            capture_output=True,
            text=True,
            cwd=str(SCRIPT_DIR)
        )

        if result.returncode != 0:
            raise Exception(f"DOCX generation failed: {result.stderr}\n{result.stdout}")

        return output_path

    finally:
        # Clean up temp file
        Path(data_path).unlink(missing_ok=True)


if __name__ == '__main__':
    # Test
    test_data = {
        'staff_name': 'Test User',
        'position': 'Construction Manager',
        'duration': '12 Months',
        'hourly_rate': '$200/hr',
        'commitment': '100%',
        'staff_monthly': '$34,600',
        'staff_total': '$415,200',
        'project_experience': 'Test experience description...',
        'project_summary': '- Project 1\n- Project 2',
        'expense_desc': 'N/A',
        'expense_monthly': 'N/A',
        'expense_total': 'N/A',
        'combined_monthly': '$34,600',
        'combined_total': '$415,200',
        'formatted_resume_path': None
    }

    output = SCRIPT_DIR / 'output' / 'test_proposal.docx'
    output.parent.mkdir(exist_ok=True)
    generate_rfq_proposal(test_data, str(output))
    print(f"Generated: {output}")
