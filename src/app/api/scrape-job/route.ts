import { NextRequest, NextResponse } from 'next/server';

// MODULE 1 — JOB DESCRIPTION SCRAPER
export async function POST(request: NextRequest) {
  try {
    const { url, rawText } = await request.json();

    // If raw text provided, normalize and structure it
    if (rawText) {
      const cleaned = cleanJobText(rawText);
      const structured = structureJobDescription(cleaned);
      return NextResponse.json({ success: true, data: structured });
    }

    // If URL provided, scrape the job
    if (url) {
      const scraped = await scrapeJobUrl(url);
      return NextResponse.json({ success: true, data: scraped });
    }

    return NextResponse.json({ success: false, error: 'No URL or text provided' }, { status: 400 });
  } catch (error: any) {
    return NextResponse.json({ success: false, error: error.message }, { status: 500 });
  }
}

// Scrape job from URL
async function scrapeJobUrl(url: string) {
  try {
    const response = await fetch(url, {
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Accept-Language': 'en-US,en;q=0.5',
      }
    });

    if (!response.ok) {
      throw new Error(`Failed to fetch URL: ${response.status}`);
    }

    const html = await response.text();
    const extracted = extractJobFromHtml(html, url);
    return extracted;
  } catch (error: any) {
    // Return structured error with partial data
    return {
      jobTitle: 'Unable to extract',
      company: 'Unknown',
      description: `Failed to scrape URL. Please paste the job description manually. Error: ${error.message}`,
      rawText: '',
      source: url
    };
  }
}

// Extract job data from HTML
function extractJobFromHtml(html: string, url: string) {
  // Remove scripts, styles, navigation
  let cleanHtml = html
    .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '')
    .replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, '')
    .replace(/<nav\b[^<]*(?:(?!<\/nav>)<[^<]*)*<\/nav>/gi, '')
    .replace(/<header\b[^<]*(?:(?!<\/header>)<[^<]*)*<\/header>/gi, '')
    .replace(/<footer\b[^<]*(?:(?!<\/footer>)<[^<]*)*<\/footer>/gi, '')
    .replace(/<!--[\s\S]*?-->/g, '');

  // Extract text content
  const text = cleanHtml
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<\/li>/gi, '\n')
    .replace(/<li[^>]*>/gi, '• ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/\s+/g, ' ')
    .trim();

  // Detect platform and extract accordingly
  const isLinkedIn = url.includes('linkedin.com');
  const isIndeed = url.includes('indeed.com');
  const isGlassdoor = url.includes('glassdoor.com');
  const isMonster = url.includes('monster.com');

  let jobTitle = '';
  let company = '';
  let location = '';

  // Platform-specific extraction patterns
  const titlePatterns = [
    /jobtitle["\s:]+([^"<>\n]+)/i,
    /"title"["\s:]+["']([^"']+)["']/i,
    /position[^:]*:\s*([^\n]+)/i,
    /job\s*title[^:]*:\s*([^\n]+)/i,
    /<h1[^>]*>([^<]+)<\/h1>/i
  ];

  const companyPatterns = [
    /company["\s:]+["']([^"']+)["']/i,
    /"companyName"["\s:]+["']([^"']+)["']/i,
    /company\s*name[^:]*:\s*([^\n]+)/i,
    /employer[^:]*:\s*([^\n]+)/i
  ];

  const locationPatterns = [
    /location["\s:]+["']([^"']+)["']/i,
    /"jobLocation"["\s:]+["']([^"']+)["']/i,
    /location[^:]*:\s*([^\n]+)/i
  ];

  // Extract job title
  for (const pattern of titlePatterns) {
    const match = text.match(pattern);
    if (match && match[1]) {
      jobTitle = match[1].trim();
      break;
    }
  }

  // Extract company
  for (const pattern of companyPatterns) {
    const match = text.match(pattern);
    if (match && match[1]) {
      company = match[1].trim();
      break;
    }
  }

  // Extract location
  for (const pattern of locationPatterns) {
    const match = text.match(pattern);
    if (match && match[1]) {
      location = match[1].trim();
      break;
    }
  }

  // Clean the full text
  const cleanedText = cleanJobText(text);

  return {
    jobTitle: jobTitle || 'Position Title Not Found',
    company: company || 'Company Not Found',
    location: location || '',
    description: cleanedText,
    rawText: cleanedText,
    source: url
  };
}

// Clean job text
function cleanJobText(text: string): string {
  return text
    // Remove cookie notices
    .replace(/cookie\s*policy[^.]*\./gi, '')
    .replace(/we\s*use\s*cookies[^.]*\./gi, '')
    .replace(/accept\s*cookies[^.]*\./gi, '')
    // Remove navigation elements
    .replace(/skip\s*to\s*content/gi, '')
    .replace(/sign\s*in/gi, '')
    .replace(/log\s*in/gi, '')
    .replace(/register/gi, '')
    // Remove duplicate lines
    .split('\n')
    .filter((line, index, self) => self.indexOf(line) === index)
    .join('\n')
    // Normalize whitespace
    .replace(/\n{3,}/g, '\n\n')
    .replace(/\s{2,}/g, ' ')
    .trim();
}

// Structure job description
function structureJobDescription(text: string) {
  // Detect sections
  const sections = {
    responsibilities: extractSection(text, ['responsibilities', 'duties', 'what you will do', 'role description']),
    requirements: extractSection(text, ['requirements', 'qualifications', 'what you need', 'skills required']),
    skills: extractSection(text, ['skills', 'competencies', 'abilities', 'technical skills']),
    benefits: extractSection(text, ['benefits', 'perks', 'what we offer', 'compensation'])
  };

  // Extract keywords
  const keywords = extractKeywords(text);

  return {
    jobTitle: extractJobTitle(text) || 'Position Not Specified',
    company: extractCompany(text) || 'Company Not Specified',
    description: text,
    responsibilities: sections.responsibilities,
    requirements: sections.requirements,
    skills: sections.skills,
    benefits: sections.benefits,
    keywords: keywords,
    rawText: text
  };
}

function extractSection(text: string, headers: string[]): string {
  const lowerText = text.toLowerCase();
  for (const header of headers) {
    const index = lowerText.indexOf(header);
    if (index !== -1) {
      const start = index + header.length;
      const endSection = text.substring(start).search(/\n\s*(?:requirements|qualifications|skills|benefits|responsibilities|about|company|apply)/i);
      const end = endSection !== -1 ? start + endSection : Math.min(start + 500, text.length);
      return text.substring(start, end).trim();
    }
  }
  return '';
}

function extractKeywords(text: string): string[] {
  const words = text.toLowerCase().match(/\b[a-z]{4,}\b/g) || [];
  const frequency: Record<string, number> = {};
  words.forEach(word => {
    frequency[word] = (frequency[word] || 0) + 1;
  });
  return Object.entries(frequency)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 30)
    .map(([word]) => word);
}

function extractJobTitle(text: string): string {
  const patterns = [
    /position[^:]*:\s*([^\n]+)/i,
    /job\s*title[^:]*:\s*([^\n]+)/i,
    /role[^:]*:\s*([^\n]+)/i,
    /title[^:]*:\s*([^\n]+)/i
  ];
  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match) return match[1].trim();
  }
  return '';
}

function extractCompany(text: string): string {
  const patterns = [
    /company[^:]*:\s*([^\n]+)/i,
    /employer[^:]*:\s*([^\n]+)/i,
    /organization[^:]*:\s*([^\n]+)/i
  ];
  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match) return match[1].trim();
  }
  return '';
}
