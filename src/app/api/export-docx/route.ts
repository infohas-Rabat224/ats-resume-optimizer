import { NextRequest, NextResponse } from 'next/server';
import { Document, Packer, Paragraph, TextRun, AlignmentType, LevelFormat } from 'docx';

// 0.95cm = ~539 twips (1 inch = 1440 twips, 1 cm = 567 twips)
const MARGIN_TWIPS = 539;

export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const { content, filename } = body;
    
    if (!content) {
      return NextResponse.json({ error: 'No content provided' }, { status: 400 });
    }
    
    // Parse HTML content and create docx elements
    const elements = parseHtmlToDocx(content);
    
    // Create the document with proper margins
    const doc = new Document({
      styles: {
        default: {
          document: {
            run: { font: "Times New Roman", size: 24 } // 12pt
          }
        }
      },
      numbering: {
        config: [
          {
            reference: "bullet-list",
            levels: [{
              level: 0,
              format: LevelFormat.BULLET,
              text: "•",
              alignment: AlignmentType.LEFT,
              style: {
                paragraph: {
                  indent: { left: 720, hanging: 360 }
                }
              }
            }]
          }
        ]
      },
      sections: [{
        properties: {
          page: {
            margin: {
              top: MARGIN_TWIPS,
              right: MARGIN_TWIPS,
              bottom: MARGIN_TWIPS,
              left: MARGIN_TWIPS
            }
          }
        },
        children: elements
      }]
    });
    
    // Generate buffer
    const buffer = await Packer.toBuffer(doc);
    
    // Return as downloadable file
    return new NextResponse(buffer, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'Content-Disposition': `attachment; filename="${filename || 'Resume.docx'}"`
      }
    });
    
  } catch (error: any) {
    console.error('DOCX generation error:', error);
    return NextResponse.json({ error: error.message }, { status: 500 });
  }
}

function parseHtmlToDocx(html: string): Paragraph[] {
  const paragraphs: Paragraph[] = [];
  
  // Clean HTML
  let cleanHtml = html.replace(/\s+/g, ' ').trim();
  
  // Extract name (h1)
  const h1Match = html.match(/<h1[^>]*>(.*?)<\/h1>/i);
  if (h1Match) {
    const name = cleanText(h1Match[1]);
    paragraphs.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 100 },
      children: [new TextRun({ text: name, bold: true, size: 32, font: "Times New Roman" })]
    }));
  }
  
  // Extract subtitle (h4)
  const h4Match = html.match(/<h4[^>]*>(.*?)<\/h4>/i);
  if (h4Match) {
    const subtitle = cleanText(h4Match[1]);
    paragraphs.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
      children: [new TextRun({ text: subtitle, size: 24, font: "Times New Roman" })]
    }));
  }
  
  // Track seen content to avoid duplicates
  const seenParagraphs = new Set<string>();
  
  // Extract all ul sections and process them
  const ulMatches = html.matchAll(/<ul[^>]*>(.*?)<\/ul>/gis);
  const ulPositions: number[] = [];
  
  for (const ulMatch of ulMatches) {
    ulPositions.push(ulMatch.index || 0);
  }
  
  // Process all <p> tags
  const pMatches = html.matchAll(/<p[^>]*>(.*?)<\/p>/gi);
  
  for (const match of pMatches) {
    const content = match[1];
    const cleanContent = cleanText(content);
    
    // Skip empty or duplicate
    if (!cleanContent || seenParagraphs.has(cleanContent.toLowerCase())) continue;
    seenParagraphs.add(cleanContent.toLowerCase());
    
    // Check if it's a section header (strong tag)
    const strongMatch = content.match(/<strong>(.*?)<\/strong>/i);
    if (strongMatch) {
      const strongText = cleanText(strongMatch[1]);
      const restText = cleanText(content.replace(/<strong>.*?<\/strong>/i, '')).replace(/^\s*[-|]\s*/, '');
      
      // Section header (all uppercase)
      if (strongText === strongText.toUpperCase() && strongText.length > 3) {
        paragraphs.push(new Paragraph({
          spacing: { before: 200, after: 100 },
          children: [new TextRun({ text: strongText, bold: true, size: 24, font: "Times New Roman" })]
        }));
        if (restText) {
          paragraphs.push(new Paragraph({
            spacing: { after: 100 },
            children: [new TextRun({ text: restText, size: 24, font: "Times New Roman" })]
          }));
        }
      } else {
        // Job title / company line
        paragraphs.push(new Paragraph({
          spacing: { after: 100 },
          children: [
            new TextRun({ text: strongText, bold: true, size: 24, font: "Times New Roman" }),
            new TextRun({ text: restText, size: 24, font: "Times New Roman" })
          ]
        }));
      }
    } else if (cleanContent) {
      // Regular paragraph
      paragraphs.push(new Paragraph({
        spacing: { after: 100 },
        children: [new TextRun({ text: cleanContent, size: 24, font: "Times New Roman" })]
      }));
    }
  }
  
  // Process list items - each bullet on its own line
  const liMatches = html.matchAll(/<li[^>]*>(.*?)<\/li>/gis);
  
  for (const liMatch of liMatches) {
    let liContent = cleanText(liMatch[1]);
    
    // Skip if empty
    if (!liContent) continue;
    
    // Remove leading bullet if present (we'll add proper bullet)
    liContent = liContent.replace(/^•\s*/i, '');
    
    // Check for category: value format
    const catMatch = liContent.match(/^(.+?):\s*(.+)$/);
    
    if (catMatch) {
      // Category with skills
      paragraphs.push(new Paragraph({
        numbering: { reference: "bullet-list", level: 0 },
        spacing: { after: 60 },
        children: [
          new TextRun({ text: catMatch[1] + ": ", bold: true, size: 24, font: "Times New Roman" }),
          new TextRun({ text: catMatch[2], size: 24, font: "Times New Roman" })
        ]
      }));
    } else {
      // Regular bullet
      paragraphs.push(new Paragraph({
        numbering: { reference: "bullet-list", level: 0 },
        spacing: { after: 60 },
        children: [new TextRun({ text: liContent, size: 24, font: "Times New Roman" })]
      }));
    }
  }
  
  return paragraphs;
}

function cleanText(html: string): string {
  return html
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/\s+/g, ' ')
    .trim();
}
