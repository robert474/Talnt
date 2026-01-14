const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType, BorderStyle, ImageRun, convertInchesToTwip, ShadingType, VerticalAlign } = require('docx');
const fs = require('fs');
const path = require('path');
const AdmZip = require('adm-zip');

// Read the proposal data
const dataPath = process.argv[2];
const outputPath = process.argv[3];
const data = JSON.parse(fs.readFileSync(dataPath, 'utf8'));

const scriptDir = path.dirname(process.argv[1]);
const parentDir = path.dirname(scriptDir);

// Select logo based on brand
const brand = data.brand || 'dc';
let logoPath, logoWidth, logoHeight;
if (brand === 'tt') {
  logoPath = path.join(parentDir, 'assets', 'TT-final-side.jpg');
  logoWidth = 200;
  logoHeight = 80;
} else {
  logoPath = path.join(parentDir, 'assets', 'datacenter-logo-black-type-transparent.png');
  logoWidth = 220;
  logoHeight = 72;
}

// Font settings
const fontSize = 20; // 10pt = 20 half-points
const headerFontSize = 24; // 12pt
const smallFontSize = 18; // 9pt for table

// Colors
const primaryColor = "1a1a2e";  // Dark navy
const accentColor = "2e7d32";   // Green for totals
const lightGray = "f5f5f5";
const mediumGray = "e0e0e0";

// Table cell borders - cleaner look
const tableBorders = {
  top: { style: BorderStyle.SINGLE, size: 1, color: mediumGray },
  bottom: { style: BorderStyle.SINGLE, size: 1, color: mediumGray },
  left: { style: BorderStyle.SINGLE, size: 1, color: mediumGray },
  right: { style: BorderStyle.SINGLE, size: 1, color: mediumGray }
};

const noBorders = {
  top: { style: BorderStyle.NONE },
  bottom: { style: BorderStyle.NONE },
  left: { style: BorderStyle.NONE },
  right: { style: BorderStyle.NONE }
};

// Helper to create styled table cell
function styledCell(text, options = {}) {
  const {
    isHeader = false,
    width = null,
    align = AlignmentType.LEFT,
    bgColor = null,
    textColor = "000000",
    bold = false,
    fontSize: cellFontSize = smallFontSize
  } = options;

  const cellOptions = {
    children: [
      new Paragraph({
        children: [
          new TextRun({
            text: text || '',
            bold: isHeader || bold,
            size: cellFontSize,
            font: "Arial",
            color: textColor
          })
        ],
        alignment: align
      })
    ],
    borders: tableBorders,
    verticalAlign: VerticalAlign.CENTER
  };

  if (width) {
    cellOptions.width = { size: width, type: WidthType.PERCENTAGE };
  }

  if (bgColor) {
    cellOptions.shading = { fill: bgColor, type: ShadingType.CLEAR };
  }

  return new TableCell(cellOptions);
}

// Create Summary Grid Table (the main visual grid)
function createSummaryGrid() {
  const duration = data.duration.toString().includes('Month') ? data.duration : data.duration + ' mo';
  const rate = data.hourly_rate.toString().includes('$') ? data.hourly_rate : '$' + data.hourly_rate + '/hr';
  const hasExpenses = data.expense_monthly && data.expense_monthly !== 'N/A' && data.expense_monthly !== '$0';

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      // Header row - dark background
      new TableRow({
        children: [
          styledCell("Type", { isHeader: true, width: 12, bgColor: primaryColor, textColor: "FFFFFF", align: AlignmentType.CENTER }),
          styledCell("Name/Description", { isHeader: true, width: 22, bgColor: primaryColor, textColor: "FFFFFF" }),
          styledCell("Position", { isHeader: true, width: 18, bgColor: primaryColor, textColor: "FFFFFF" }),
          styledCell("Duration", { isHeader: true, width: 10, bgColor: primaryColor, textColor: "FFFFFF", align: AlignmentType.CENTER }),
          styledCell("Rate", { isHeader: true, width: 10, bgColor: primaryColor, textColor: "FFFFFF", align: AlignmentType.CENTER }),
          styledCell("Monthly", { isHeader: true, width: 14, bgColor: primaryColor, textColor: "FFFFFF", align: AlignmentType.RIGHT }),
          styledCell("Total", { isHeader: true, width: 14, bgColor: primaryColor, textColor: "FFFFFF", align: AlignmentType.RIGHT })
        ]
      }),
      // Staff row
      new TableRow({
        children: [
          styledCell("Staff", { width: 12, bold: true, align: AlignmentType.CENTER }),
          styledCell(data.staff_name, { width: 22 }),
          styledCell(data.position, { width: 18 }),
          styledCell(duration, { width: 10, align: AlignmentType.CENTER }),
          styledCell(rate, { width: 10, align: AlignmentType.CENTER }),
          styledCell(data.staff_monthly, { width: 14, align: AlignmentType.RIGHT, textColor: accentColor }),
          styledCell(data.staff_total, { width: 14, align: AlignmentType.RIGHT, textColor: accentColor })
        ]
      }),
      // Expenses row
      new TableRow({
        children: [
          styledCell("Expenses", { width: 12, bold: true, align: AlignmentType.CENTER }),
          styledCell(hasExpenses ? (data.expense_type || data.expense_desc || '-') : '-', { width: 22 }),
          styledCell("-", { width: 18 }),
          styledCell(hasExpenses ? duration : '-', { width: 10, align: AlignmentType.CENTER }),
          styledCell("-", { width: 10, align: AlignmentType.CENTER }),
          styledCell(hasExpenses ? data.expense_monthly : '$0', { width: 14, align: AlignmentType.RIGHT }),
          styledCell(hasExpenses ? data.expense_total : '$0', { width: 14, align: AlignmentType.RIGHT })
        ]
      }),
      // Combined row - highlighted
      new TableRow({
        children: [
          styledCell("COMBINED", { width: 12, bold: true, bgColor: "e8f5e9", align: AlignmentType.CENTER }),
          styledCell("", { width: 22, bgColor: "e8f5e9" }),
          styledCell("", { width: 18, bgColor: "e8f5e9" }),
          styledCell("", { width: 10, bgColor: "e8f5e9" }),
          styledCell("", { width: 10, bgColor: "e8f5e9" }),
          styledCell(data.combined_monthly, { width: 14, align: AlignmentType.RIGHT, bgColor: "e8f5e9", bold: true, textColor: accentColor }),
          styledCell(data.combined_total, { width: 14, align: AlignmentType.RIGHT, bgColor: "e8f5e9", bold: true, textColor: "1b5e20", fontSize: 22 })
        ]
      })
    ]
  });
}

// Split project summary into properly formatted paragraphs
function createProjectSummaryParagraphs() {
  if (!data.project_summary) return [];

  const lines = data.project_summary.split('\n').filter(line => line.trim());
  return lines.map(line => {
    let cleanLine = line.replace(/^[-•*●]\s*/, '').trim();

    return new Paragraph({
      children: [
        new TextRun({ text: "  •  ", size: fontSize, font: "Arial", color: accentColor }),
        new TextRun({ text: cleanLine, size: fontSize, font: "Arial" })
      ],
      indent: {
        left: convertInchesToTwip(0.25)
      },
      spacing: { after: 120 }
    });
  });
}

// Extract text content from formatted resume DOCX
function extractResumeContent() {
  if (!data.formatted_resume_path || !fs.existsSync(data.formatted_resume_path)) {
    return [
      new Paragraph({
        children: [
          new TextRun({ text: "<Resume not provided>", italics: true, size: fontSize, font: "Arial", color: "666666" })
        ]
      })
    ];
  }

  try {
    const zip = new AdmZip(data.formatted_resume_path);
    const documentXml = zip.readAsText('word/document.xml');

    const paragraphs = [];
    const textMatches = documentXml.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || [];
    const paraMatches = documentXml.split(/<\/w:p>/);

    paraMatches.forEach(paraXml => {
      const texts = paraXml.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || [];
      if (texts.length === 0) return;

      const runs = [];
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

// Format dates for display
function formatDate(dateStr) {
  if (!dateStr) return '';
  const date = new Date(dateStr);
  return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
}

const startDateFormatted = formatDate(data.start_date);
const endDateFormatted = formatDate(data.end_date);
const dateRange = startDateFormatted && endDateFormatted ? `${startDateFormatted} - ${endDateFormatted}` : '';

// Build the document
const children = [
  // Logo
  new Paragraph({
    children: [
      new ImageRun({
        type: brand === 'tt' ? "jpg" : "png",
        data: fs.readFileSync(logoPath),
        transformation: { width: logoWidth, height: logoHeight },
        altText: { title: "Logo", description: brand === 'tt' ? "Talnt Team Logo" : "Data Center TALNT Logo", name: "Logo" }
      })
    ],
    spacing: { after: 400 }
  }),

  // Title
  new Paragraph({
    children: [
      new TextRun({
        text: "RFQ PROPOSAL",
        bold: true,
        size: 36,
        font: "Arial",
        color: primaryColor
      })
    ],
    alignment: AlignmentType.CENTER,
    spacing: { after: 100 }
  }),

  // Candidate name and position
  new Paragraph({
    children: [
      new TextRun({
        text: data.staff_name,
        bold: true,
        size: 28,
        font: "Arial"
      }),
      new TextRun({
        text: `  |  ${data.position}`,
        size: 24,
        font: "Arial",
        color: "666666"
      })
    ],
    alignment: AlignmentType.CENTER,
    spacing: { after: 80 }
  }),

  // Project Dates (if provided)
  ...(dateRange ? [new Paragraph({
    children: [
      new TextRun({
        text: dateRange,
        size: fontSize,
        font: "Arial",
        italics: true,
        color: "666666"
      })
    ],
    alignment: AlignmentType.CENTER,
    spacing: { after: 300 }
  })] : [new Paragraph({ children: [], spacing: { after: 200 } })]),

  // Summary Grid Table
  createSummaryGrid(),

  // Spacing after grid
  new Paragraph({ children: [], spacing: { after: 400 } }),

  // Detailed Project Experience Header
  new Paragraph({
    children: [
      new TextRun({ text: "DETAILED PROJECT EXPERIENCE", bold: true, size: headerFontSize, font: "Arial", color: primaryColor })
    ],
    border: {
      bottom: { style: BorderStyle.SINGLE, size: 12, color: primaryColor }
    },
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
      new TextRun({ text: "PROJECT SUMMARY", bold: true, size: headerFontSize, font: "Arial", color: primaryColor })
    ],
    border: {
      bottom: { style: BorderStyle.SINGLE, size: 12, color: primaryColor }
    },
    spacing: { after: 200 }
  }),

  // Project Summary Bullets
  ...createProjectSummaryParagraphs(),

  // Page break before resume
  new Paragraph({
    children: [],
    pageBreakBefore: true
  }),

  // Resume Header
  new Paragraph({
    children: [
      new TextRun({ text: "RESUME", bold: true, size: headerFontSize, font: "Arial", color: primaryColor })
    ],
    border: {
      bottom: { style: BorderStyle.SINGLE, size: 12, color: primaryColor }
    },
    spacing: { after: 300 }
  }),

  // Resume Content
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
        size: { width: 12240, height: 15840 },
        margin: { top: 720, right: 720, bottom: 720, left: 720 }
      }
    },
    children: children
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(outputPath, buffer);
  console.log("RFQ Proposal generated successfully!");
});
