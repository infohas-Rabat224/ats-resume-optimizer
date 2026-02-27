import { NextRequest, NextResponse } from 'next/server';

// MODULE 2 — RESUME EXTRACTION ENGINE
export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get('file') as File;
    const apiKey = formData.get('apiKey') as string;
    const model = formData.get('model') as string || 'gemini-2.0-flash';

    if (!file) {
      return NextResponse.json({ success: false, error: 'No file provided' }, { status: 400 });
    }

    const fileName = file.name.toLowerCase();
    const mimeType = file.type;
    
    let extractedText = '';

    // Handle PDF and images with AI vision
    if (fileName.endsWith('.pdf') || fileName.endsWith('.png') || fileName.endsWith('.jpg') || fileName.endsWith('.jpeg')) {
      if (!apiKey) {
        return NextResponse.json({ 
          success: false, 
          error: 'API key required for PDF/image extraction' 
        }, { status: 400 });
      }
      
      const buffer = await file.arrayBuffer();
      const base64 = Buffer.from(buffer).toString('base64');
      
      extractedText = await extractWithAI(base64, mimeType, apiKey, model);
    }
    // Handle text files directly
    else if (fileName.endsWith('.txt') || fileName.endsWith('.md')) {
      extractedText = await file.text();
    }
    // Handle DOCX (as zip-based format - basic extraction)
    else if (fileName.endsWith('.docx')) {
      extractedText = await extractFromDocx(file);
    }
    else {
      return NextResponse.json({ 
        success: false, 
        error: 'Unsupported file format. Use PDF, DOCX, TXT, or images.' 
      }, { status: 400 });
    }

    // Clean and structure the extracted text
    const cleanedText = cleanResumeText(extractedText);
    const sections = detectSections(cleanedText);
    
    return NextResponse.json({ 
      success: true, 
      data: {
        rawText: extractedText,
        cleanedText,
        sections,
        wordCount: cleanedText.split(/\s+/).length,
        characterCount: cleanedText.length
      }
    });
  } catch (error: any) {
    return NextResponse.json({ success: false, error: error.message }, { status: 500 });
  }
}

async function extractWithAI(base64: string, mimeType: string, apiKey: string, model: string): Promise<string> {
  // Use Gemini for vision extraction
  const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      contents: [{
        parts: [
          { inlineData: { mimeType, data: base64 } },
          { text: `Extract ALL text from this resume document. 
          
Return ONLY the extracted text in this exact format:

FULL NAME
[Contact Information]

PROFESSIONAL SUMMARY
[Summary paragraph]

CORE COMPETENCIES
[List of skills]

PROFESSIONAL EXPERIENCE
[Job entries with title, company, dates, and bullet points]

EDUCATION
[Education entries]

LANGUAGES
[Language proficiencies]

CERTIFICATIONS
[Certifications if any]

Extract text exactly as shown. Do not add any commentary. Only extract the text content.` }
        ]
      }],
      generationConfig: { maxOutputTokens: 4096 }
    })
  });

  const data = await response.json();
  return data.candidates?.[0]?.content?.parts?.[0]?.text || '';
}

async function extractFromDocx(file: File): Promise<string> {
  // Basic DOCX extraction - reads XML content
  try {
    const buffer = await file.arrayBuffer();
    const zip = await import('zlib');
    
    // DOCX is a zip file - we'll do a simple text extraction
    // For production, consider using mammoth.js or similar library
    const text = await file.text();
    
    // Try to extract readable text from XML-like content
    const textMatches = text.match(/<w:t[^>]*>([^<]+)<\/w:t>/g) || [];
    const extractedText = textMatches
      .map(match => match.replace(/<[^>]+>/g, ''))
      .join(' ')
      .replace(/\s+/g, ' ')
      .trim();
    
    return extractedText || 'Could not extract text from DOCX. Please paste your resume text directly.';
  } catch (error) {
    return 'DOCX extraction failed. Please paste your resume text directly or use PDF format.';
  }
}

function cleanResumeText(text: string): string {
  return text
    // Remove excessive line breaks
    .replace(/\n{3,}/g, '\n\n')
    // Normalize spacing
    .replace(/[ \t]{2,}/g, ' ')
    // Remove special characters that might interfere
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '')
    // Normalize quotes
    .replace(/['']/g, "'")
    .replace(/[""]/g, '"')
    // Clean up bullet points
    .replace(/[•●○▪▫]/g, '•')
    // Trim whitespace
    .trim();
}

interface DetectedSections {
  header: { name: string; contact: string[] };
  professionalSummary: string;
  coreCompetencies: string[];
  professionalExperience: Array<{
    title: string;
    company: string;
    location: string;
    dates: string;
    bullets: string[];
  }>;
  education: Array<{
    degree: string;
    institution: string;
    location: string;
    dates: string;
  }>;
  languages: string[];
  certifications: string[];
  skills: string[];
}

function detectSections(text: string): DetectedSections {
  const lines = text.split('\n').map(l => l.trim()).filter(l => l);
  
  const sections: DetectedSections = {
    header: { name: '', contact: [] },
    professionalSummary: '',
    coreCompetencies: [],
    professionalExperience: [],
    education: [],
    languages: [],
    certifications: [],
    skills: []
  };

  let currentSection = 'header';
  let currentJob: any = null;
  let currentEducation: any = null;

  const sectionHeaders: Record<string, string[]> = {
    summary: ['professional summary', 'summary', 'profile', 'objective', 'about me'],
    competencies: ['core competencies', 'competencies', 'skills', 'key skills', 'technical skills', 'expertise'],
    experience: ['professional experience', 'experience', 'work history', 'employment', 'career history'],
    education: ['education', 'academic background', 'qualifications', 'academic qualifications'],
    languages: ['languages', 'language skills', 'linguistic skills'],
    certifications: ['certifications', 'certificates', 'professional certifications', 'licenses']
  };

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].toLowerCase();
    
    // Detect section headers
    for (const [section, headers] of Object.entries(sectionHeaders)) {
      if (headers.some(h => line.includes(h))) {
        currentSection = section;
        break;
      }
    }

    // Parse content based on current section
    if (currentSection === 'header') {
      if (!sections.header.name && lines[i].match(/^[A-Z][A-Z\s]+$/)) {
        sections.header.name = lines[i];
      } else if (lines[i].match(/[@\d\+\-\(\)]/) || lines[i].match(/^[a-z]+,\s*[a-z]+/i)) {
        sections.header.contact.push(lines[i]);
      }
    }
    
    else if (currentSection === 'summary') {
      if (line.length > 50 && !line.match(/^[a-z]+$/i)) {
        sections.professionalSummary += (sections.professionalSummary ? ' ' : '') + lines[i];
      }
    }
    
    else if (currentSection === 'competencies') {
      if (line.startsWith('•') || line.startsWith('-') || line.startsWith('*')) {
        sections.coreCompetencies.push(line.replace(/^[•\-\*]\s*/, ''));
      } else if (line.includes(',') && line.length < 200) {
        sections.coreCompetencies.push(...line.split(',').map(s => s.trim()));
      }
    }
    
    else if (currentSection === 'experience') {
      // Detect job title pattern
      if (lines[i].match(/^[A-Z][a-z]+(\s+[A-Z][a-z]+)*\s*\|/) || 
          lines[i].match(/^[A-Z][a-z]+(\s+[A-Z][a-z]+)*\s+at\s+/i)) {
        if (currentJob) sections.professionalExperience.push(currentJob);
        const parts = lines[i].split('|');
        currentJob = {
          title: parts[0]?.trim() || '',
          company: parts[1]?.split(',')[0]?.trim() || '',
          location: parts[1]?.split(',')[1]?.trim() || '',
          dates: parts[2]?.trim() || '',
          bullets: []
        };
      } else if (line.startsWith('•') || line.startsWith('-') || line.startsWith('*')) {
        if (currentJob) {
          currentJob.bullets.push(line.replace(/^[•\-\*]\s*/, ''));
        }
      }
    }
    
    else if (currentSection === 'education') {
      if (lines[i].match(/Bachelor|Master|Diploma|Degree|MBA|PhD|BSc|MSc/i)) {
        if (currentEducation) sections.education.push(currentEducation);
        currentEducation = {
          degree: lines[i],
          institution: '',
          location: '',
          dates: ''
        };
      } else if (currentEducation && !currentEducation.institution) {
        currentEducation.institution = lines[i];
      }
    }
    
    else if (currentSection === 'languages') {
      if (line.match(/[a-z]+\s*[-–]\s*[a-z]+/i)) {
        sections.languages.push(line);
      }
    }
    
    else if (currentSection === 'certifications') {
      if (line.startsWith('•') || line.startsWith('-') || line.match(/[A-Z]{2,}/)) {
        sections.certifications.push(line.replace(/^[•\-\*]\s*/, ''));
      }
    }
  }

  // Add last entries
  if (currentJob) sections.professionalExperience.push(currentJob);
  if (currentEducation) sections.education.push(currentEducation);

  return sections;
}
