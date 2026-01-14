
const { Document, Packer, Paragraph, TextRun, ImageRun, AlignmentType, convertInchesToTwip } = require('docx');
const fs = require('fs');
const path = require('path');

// Read the parsed resume data
const dataPath = process.argv[2];
const outputPath = process.argv[3];
const data = JSON.parse(fs.readFileSync(dataPath, 'utf8'));

// Keywords to highlight
const highlightKeywords = ['AWS', 'Amazon', 'Google', 'Data Center', 'Microsoft', 'data center'];

function highlightText(text, fontSize = 18) {
  const textRuns = [];
  let remaining = text;
  let foundKeyword = false;
  
  for (const keyword of highlightKeywords) {
    const regex = new RegExp(`(${keyword})`, 'gi');
    const parts = remaining.split(regex);
    
    if (parts.length > 1) {
      foundKeyword = true;
      const runs = [];
      for (let i = 0; i < parts.length; i++) {
        if (parts[i]) {
          const isBold = parts[i].toLowerCase() === keyword.toLowerCase();
          runs.push(new TextRun({ text: parts[i], bold: isBold, size: fontSize, font: "Arial" }));
        }
      }
      return runs;
    }
  }
  
  return [new TextRun({ text: text, size: fontSize, font: "Arial" })];
}

// Arial 9pt = 18 half-points
const fontSize = 18;
const scriptDir = path.dirname(process.argv[1]);

// Select logo based on brand environment variable
const brand = process.env.TALNT_BRAND || 'dc';
let logoPath, logoWidth, logoHeight;
if (brand === 'tt') {
  logoPath = path.join(scriptDir, 'assets', 'TT-final-side.jpg');
  logoWidth = 180;
  logoHeight = 72;
} else {
  logoPath = path.join(scriptDir, 'assets', 'datacenter-logo-black-type-transparent.png');
  logoWidth = 200;
  logoHeight = 65;
}

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
    children: [
      // Logo - uses brand selection
      new Paragraph({
        children: [
          new ImageRun({
            type: brand === 'tt' ? "jpg" : "png",
            data: fs.readFileSync(logoPath),
            transformation: { width: logoWidth, height: logoHeight },
            altText: { title: "Logo", description: brand === 'tt' ? "Talnt Team Logo" : "Data Center TALNT Logo", name: "Logo" }
          })
        ],
        spacing: { after: 200 }
      }),
      
      // Name - CENTERED
      new Paragraph({
        children: [new TextRun({ text: data.name || "Name", size: fontSize, font: "Arial" })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 }
      }),
      
      // Professional Summary HEADER
      new Paragraph({
        children: [new TextRun({ text: "Professional Summary", bold: true, underline: {}, size: fontSize, font: "Arial" })],
        spacing: { after: 120 }
      }),

      // Professional Summary TEXT
      new Paragraph({
        children: highlightText(data.summary || ""),
        spacing: { after: 200 }
      }),

      // Education HEADER
      new Paragraph({
        children: [new TextRun({ text: "Education", bold: true, underline: {}, size: fontSize, font: "Arial" })],
        spacing: { after: 120 }
      }),
      
      // Education entries - all on one line with spaces
      new Paragraph({
        children: (data.education || []).flatMap((edu, idx) => {
          const parts = [];
          if (idx > 0) parts.push(new TextRun({ text: " ", size: fontSize, font: "Arial" }));
          
          if (edu.degree && edu.school) {
            parts.push(new TextRun({ text: edu.degree, bold: true, size: fontSize, font: "Arial" }));
            parts.push(new TextRun({ text: " – ", size: fontSize, font: "Arial" }));
            parts.push(new TextRun({ text: edu.school, size: fontSize, font: "Arial" }));
            if (edu.year) {
              parts.push(new TextRun({ text: ", " + edu.year, size: fontSize, font: "Arial" }));
            }
          } else if (edu.school) {
            parts.push(new TextRun({ text: edu.school, size: fontSize, font: "Arial" }));
          }
          
          return parts;
        }),
        spacing: { after: 200 }
      }),
      
      // Employment History HEADER
      new Paragraph({
        children: [new TextRun({ text: "Employment History", bold: true, underline: {}, size: fontSize, font: "Arial" })],
        spacing: { after: 120 }
      }),
      
      // Experience entries
      ...(data.experience || []).flatMap(exp => {
        const elements = [];
        
        // Company | Title | Location | Dates (skip empty location)
        const headerParts = [
          new TextRun({ text: exp.company || '', bold: true, size: fontSize, font: "Arial" }),
          new TextRun({ text: " | ", size: fontSize, font: "Arial" }),
          new TextRun({ text: exp.title || '', bold: true, size: fontSize, font: "Arial" })
        ];
        // Only add location if it exists
        if (exp.location && exp.location.trim()) {
          headerParts.push(new TextRun({ text: " | ", size: fontSize, font: "Arial" }));
          headerParts.push(new TextRun({ text: exp.location, size: fontSize, font: "Arial" }));
        }
        // Add dates
        headerParts.push(new TextRun({ text: " | ", size: fontSize, font: "Arial" }));
        headerParts.push(new TextRun({ text: exp.dates || '', size: fontSize, font: "Arial" }));

        elements.push(
          new Paragraph({
            children: headerParts,
            spacing: { after: 80 }
          })
        );
        
        // Projects line if exists (italicized)
        if (exp.project_details) {
          elements.push(
            new Paragraph({
              children: [
                new TextRun({ text: "Projects: ", italics: true, size: fontSize, font: "Arial" }),
                ...highlightText(exp.project_details, fontSize).map(run => 
                  new TextRun({ ...run, italics: true, font: "Arial" })
                )
              ],
              spacing: { after: 80 }
            })
          );
        }
        
        // Bullets - using proper Word indentation
        (exp.bullets || []).forEach((bullet, idx) => {
          // Remove any existing bullet characters
          const cleanBullet = bullet.replace(/^[•●\-\*]\s*/, '');

          elements.push(
            new Paragraph({
              children: [
                new TextRun({ text: "• ", size: fontSize, font: "Arial" }),
                ...highlightText(cleanBullet, fontSize)
              ],
              indent: {
                left: convertInchesToTwip(0.25),
                hanging: convertInchesToTwip(0.125)
              },
              spacing: { after: idx === exp.bullets.length - 1 ? 200 : 60 }
            })
          );
        });
        
        return elements;
      }),
      
      // Certifications (if exists)
      ...(data.certifications && data.certifications.length > 0 ? [
        new Paragraph({
          children: [new TextRun({ text: "Certifications", bold: true, underline: {}, size: fontSize, font: "Arial" })],
          spacing: { after: 120 }
        }),
        ...data.certifications.map((cert, idx) =>
          new Paragraph({
            children: [new TextRun({ text: cert, size: fontSize, font: "Arial" })],
            spacing: { after: idx === data.certifications.length - 1 ? 200 : 60 }
          })
        )
      ] : []),
      
      // Technical Tools (if exists)
      ...(data.skills ? [
        new Paragraph({
          children: [new TextRun({ text: "Technical Tools", bold: true, underline: {}, size: fontSize, font: "Arial" })],
          spacing: { after: 120 }
        }),
        new Paragraph({
          children: [new TextRun({ text: data.skills, size: fontSize, font: "Arial" })],
          spacing: { after: 120 }
        })
      ] : [])
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(outputPath, buffer);
  console.log("Resume formatted successfully!");
});
