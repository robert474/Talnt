const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType, BorderStyle, ImageRun, convertInchesToTwip } = require('docx');
const fs = require('fs');
const path = require('path');
const AdmZip = require('adm-zip');

// Read the proposal data
const dataPath = process.argv[2];
const outputPath = process.argv[3];
const data = JSON.parse(fs.readFileSync(dataPath, 'utf8'));

const scriptDir = path.dirname(process.argv[1]);
const parentDir = path.dirname(scriptDir);
const logoPath = path.join(parentDir, 'assets', 'datacenter-logo-black-type-transparent.png');

// Font settings
const fontSize = 20; // 10pt = 20 half-points
const headerFontSize = 24; // 12pt

// Table cell borders
const tableBorders = {
  top: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
  bottom: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
  left: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
  right: { style: BorderStyle.SINGLE, size: 1, color: "000000" }
};

// Helper to create table cell
function cell(text, isHeader = false, width = null) {
  const cellOptions = {
    children: [
      new Paragraph({
        children: [
          new TextRun({
            text: text || '',
            bold: isHeader,
            size: fontSize,
            font: "Arial"
          })
        ]
      })
    ],
    borders: tableBorders
  };

  if (width) {
    cellOptions.width = { size: width, type: WidthType.PERCENTAGE };
  }

  return new TableCell(cellOptions);
}

// Create Staff Table
function createStaffTable() {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      // Header row
      new TableRow({
        children: [
          cell("Staff Name", true, 18),
          cell("Position", true, 18),
          cell("Duration", true, 12),
          cell("Hourly Rate", true, 12),
          cell("Commitment", true, 12),
          cell("Monthly", true, 14),
          cell("Total", true, 14)
        ]
      }),
      // Data row
      new TableRow({
        children: [
          cell(data.staff_name),
          cell(data.position),
          cell(data.duration.toString().includes('Month') ? data.duration : data.duration + ' Months'),
          cell(data.hourly_rate.toString().includes('$') ? data.hourly_rate : '$' + data.hourly_rate + '/hr'),
          cell(data.commitment.toString().includes('%') ? data.commitment : data.commitment + '%'),
          cell(data.staff_monthly),
          cell(data.staff_total)
        ]
      })
    ]
  });
}

// Create Expenses Table
function createExpensesTable() {
  const duration = data.duration.toString().includes('Month') ? data.duration : data.duration + ' Months';

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: [
          cell("Expenses", true, 18),
          cell("Description", true, 18),
          cell("Duration", true, 12),
          cell("Hourly Rate", true, 12),
          cell("Commitment", true, 12),
          cell("Monthly", true, 14),
          cell("Total", true, 14)
        ]
      }),
      new TableRow({
        children: [
          cell(data.expense_type || ""),
          cell(data.expense_desc || "N/A"),
          cell(data.expense_desc === "N/A" ? "" : duration),
          cell(""),
          cell(""),
          cell(data.expense_monthly || "N/A"),
          cell(data.expense_total || "N/A")
        ]
      })
    ]
  });
}

// Create Combined Table
function createCombinedTable() {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: [
          cell("Combined", true, 18),
          cell("Description", true, 18),
          cell("Duration", true, 12),
          cell("Hourly Rate", true, 12),
          cell("Commitment", true, 12),
          cell("Monthly Total", true, 14),
          cell("Project Total", true, 14)
        ]
      }),
      new TableRow({
        children: [
          cell(""),
          cell(""),
          cell(""),
          cell(""),
          cell(""),
          cell(data.combined_monthly),
          cell(data.combined_total)
        ]
      })
    ]
  });
}

// Split project summary into properly formatted paragraphs (not using broken bullet system)
function createProjectSummaryParagraphs() {
  if (!data.project_summary) return [];

  const lines = data.project_summary.split('\n').filter(line => line.trim());
  return lines.map(line => {
    // Remove existing bullet markers
    let cleanLine = line.replace(/^[-•*●]\s*/, '').trim();

    return new Paragraph({
      children: [
        new TextRun({ text: "• ", size: fontSize, font: "Arial" }),
        new TextRun({ text: cleanLine, size: fontSize, font: "Arial" })
      ],
      indent: {
        left: convertInchesToTwip(0.25),
        hanging: convertInchesToTwip(0.25)
      },
      spacing: { after: 100 }
    });
  });
}

// Extract text content from formatted resume DOCX
function extractResumeContent() {
  if (!data.formatted_resume_path || !fs.existsSync(data.formatted_resume_path)) {
    return [
      new Paragraph({
        children: [
          new TextRun({ text: "<Resume not provided>", italics: true, size: fontSize, font: "Arial" })
        ]
      })
    ];
  }

  try {
    // Read the DOCX file (it's a ZIP)
    const zip = new AdmZip(data.formatted_resume_path);
    const documentXml = zip.readAsText('word/document.xml');

    // Parse XML and extract text
    const paragraphs = [];
    let currentText = '';
    let isBold = false;
    let isUnderline = false;

    // Simple regex-based extraction of text from DOCX XML
    const textMatches = documentXml.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || [];
    const paraMatches = documentXml.split(/<\/w:p>/);

    paraMatches.forEach(paraXml => {
      // Check for paragraph properties
      const texts = paraXml.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || [];
      if (texts.length === 0) return;

      const runs = [];

      // Extract runs with formatting
      const runMatches = paraXml.match(/<w:r>[\s\S]*?<\/w:r>/g) || [];
      runMatches.forEach(runXml => {
        const textMatch = runXml.match(/<w:t[^>]*>([^<]*)<\/w:t>/);
        if (textMatch) {
          const text = textMatch[1];
          const bold = runXml.includes('<w:b/>') || runXml.includes('<w:b ');
          const underline = runXml.includes('<w:u ');

          if (text.trim()) {
            runs.push(new TextRun({
              text: text,
              bold: bold,
              underline: underline ? {} : undefined,
              size: fontSize,
              font: "Arial"
            }));
          }
        }
      });

      if (runs.length > 0) {
        paragraphs.push(new Paragraph({
          children: runs,
          spacing: { after: 100 }
        }));
      }
    });

    if (paragraphs.length === 0) {
      // Fallback: simple text extraction
      const allText = textMatches.map(m => m.replace(/<[^>]+>/g, '')).join(' ');
      return [
        new Paragraph({
          children: [
            new TextRun({ text: allText.substring(0, 5000), size: fontSize, font: "Arial" })
          ]
        })
      ];
    }

    return paragraphs;

  } catch (err) {
    console.error('Error extracting resume:', err);
    return [
      new Paragraph({
        children: [
          new TextRun({
            text: "Error embedding resume. See separate file: " + path.basename(data.formatted_resume_path),
            italics: true,
            size: fontSize,
            font: "Arial"
          })
        ]
      })
    ];
  }
}

// Build the document
const children = [
  // Logo
  new Paragraph({
    children: [
      new ImageRun({
        type: "png",
        data: fs.readFileSync(logoPath),
        transformation: { width: 220, height: 72 },
        altText: { title: "Logo", description: "Data Center TALNT Logo", name: "Logo" }
      })
    ],
    spacing: { after: 300 }
  }),

  // TL;DR Summary Header
  new Paragraph({
    children: [
      new TextRun({
        text: `PROJECT TOTAL: ${data.combined_total}`,
        bold: true,
        size: 32, // 16pt
        font: "Arial"
      })
    ],
    alignment: AlignmentType.CENTER,
    spacing: { after: 100 }
  }),

  new Paragraph({
    children: [
      new TextRun({
        text: `${data.staff_name} | ${data.position} | ${data.duration} Months @ $${data.hourly_rate}/hr`,
        size: fontSize,
        font: "Arial"
      })
    ],
    alignment: AlignmentType.CENTER,
    spacing: { after: 400 }
  }),

  // Detailed Project Experience Header
  new Paragraph({
    children: [
      new TextRun({ text: "Detailed Project Experience:", bold: true, underline: {}, size: headerFontSize, font: "Arial" })
    ],
    spacing: { after: 200 }
  }),

  // Project Experience Content
  new Paragraph({
    children: [
      new TextRun({ text: data.project_experience || "", size: fontSize, font: "Arial" })
    ],
    spacing: { after: 300 }
  }),

  // Project Summary Header
  new Paragraph({
    children: [
      new TextRun({ text: "Project Summary:", bold: true, underline: {}, size: headerFontSize, font: "Arial" })
    ],
    spacing: { after: 200 }
  }),

  // Project Summary Bullets
  ...createProjectSummaryParagraphs(),

  // Spacing before tables
  new Paragraph({ children: [], spacing: { after: 300 } }),

  // Staff Table
  createStaffTable(),

  // Spacing
  new Paragraph({ children: [], spacing: { after: 200 } }),

  // Expenses Table
  createExpensesTable(),

  // Spacing
  new Paragraph({ children: [], spacing: { after: 200 } }),

  // Combined Table
  createCombinedTable(),

  // Page break before resume
  new Paragraph({
    children: [],
    pageBreakBefore: true
  }),

  // Resume Header
  new Paragraph({
    children: [
      new TextRun({ text: "Resume:", bold: true, underline: {}, size: headerFontSize, font: "Arial" })
    ],
    spacing: { after: 200 }
  }),

  // Resume Content - now embedded
  ...extractResumeContent()
];

const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: "Arial", size: fontSize }
      }
    }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 }, // Letter size
        margin: { top: 720, right: 720, bottom: 720, left: 720 } // 0.5 inch margins
      }
    },
    children: children
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(outputPath, buffer);
  console.log("RFQ Proposal generated successfully!");
});
