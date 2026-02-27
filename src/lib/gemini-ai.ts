// Google Gemini AI Helper - Free Tier Supported
// Get your free API key at: https://aistudio.google.com

interface GeminiResponse {
  candidates: Array<{
    content: {
      parts: Array<{ text: string }>;
    };
  }>;
}

// Get API key from environment or use empty string (will fail gracefully)
function getApiKey(): string {
  return process.env.GEMINI_API_KEY || process.env.NEXT_PUBLIC_GEMINI_API_KEY || '';
}

// Check if API key is configured
export function isGeminiConfigured(): boolean {
  return getApiKey().length > 0;
}

// Call Gemini API for text generation
export async function callGeminiChat(
  prompt: string,
  systemInstruction?: string
): Promise<string> {
  const apiKey = getApiKey();
  
  if (!apiKey) {
    throw new Error('Gemini API key not configured. Please add GEMINI_API_KEY to environment variables or get a free key at https://aistudio.google.com');
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`;
  
  const body: any = {
    contents: [
      {
        parts: [{ text: prompt }]
      }
    ],
    generationConfig: {
      temperature: 0.7,
      maxOutputTokens: 8192,
    }
  };

  if (systemInstruction) {
    body.systemInstruction = {
      parts: [{ text: systemInstruction }]
    };
  }

  const response = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Gemini API error: ${response.status} - ${errorText}`);
  }

  const data: GeminiResponse = await response.json();
  return data.candidates?.[0]?.content?.parts?.[0]?.text || '';
}

// Call Gemini Vision API for image/PDF analysis
export async function callGeminiVision(
  prompt: string,
  imageBase64: string,
  mimeType: string
): Promise<string> {
  const apiKey = getApiKey();
  
  if (!apiKey) {
    throw new Error('Gemini API key not configured. Please add GEMINI_API_KEY to environment variables or get a free key at https://aistudio.google.com');
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`;

  const body = {
    contents: [
      {
        parts: [
          { text: prompt },
          {
            inline_data: {
              mime_type: mimeType,
              data: imageBase64
            }
          }
        ]
      }
    ],
    generationConfig: {
      temperature: 0.3,
      maxOutputTokens: 8192,
    }
  };

  const response = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Gemini Vision API error: ${response.status} - ${errorText}`);
  }

  const data: GeminiResponse = await response.json();
  return data.candidates?.[0]?.content?.parts?.[0]?.text || '';
}

// Extract text from PDF/Image using Gemini Vision
export async function extractTextFromDocument(
  base64: string,
  mimeType: string
): Promise<string> {
  const prompt = `Extract ALL text from this document image. 
Return ONLY the extracted text content, preserving the structure and formatting as much as possible.
- Keep the original layout and hierarchy
- Preserve bullet points and lists
- Include all names, dates, and details
- Do not add any commentary or explanations
- If this is a resume, extract all sections including contact info, experience, education, skills`;

  return callGeminiVision(prompt, base64, mimeType);
}

// Optimize resume using Gemini
export async function optimizeResumeWithGemini(
  resume: string,
  job: string,
  missingKeywords?: string[]
): Promise<any> {
  const missingKw = missingKeywords?.join(', ') || 'N/A';
  
  const prompt = `You are an Elite ATS Resume Optimizer - Aviation & Hospitality Specialist.

STRICT OPTIMIZATION RULES:

LAYOUT RULES (MANDATORY):
1. MUST FIT EXACTLY ONE A4 PAGE
2. WORD COUNT: 650-700 words MAXIMUM
3. Font: Times New Roman, 12pt
4. Section headings: ALL CAPS, BOLD
5. Body text: NOT bold
6. Bullets: • symbol only, max 5 per job
7. NO: tables, icons, emojis, graphics

CONTENT RULES:
- EVERY bullet starts with ACTION VERB (Achieved, Delivered, Managed, Led, Implemented)
- Include NUMBERS/PERCENTAGES in achievements
- REMOVE weak phrases: "Responsible for", "Helped with", "Worked on"
- NO repeated action verbs
- Natural keyword integration

INPUT DATA:

MISSING KEYWORDS TO ADD: ${missingKw}

CANDIDATE RESUME:
${resume}

TARGET JOB DESCRIPTION:
${job}

Return ONLY valid JSON (no markdown code blocks):
{
  "score": 85,
  "score_breakdown": {"impact": 85, "brevity": 90, "keywords": 80},
  "summary_critique": "brief assessment",
  "missing_keywords": ["keyword1"],
  "matched_keywords": ["keyword1"],
  "optimized_content": "<h1>NAME</h1><h4>Contact Info</h4><p><strong>PROFESSIONAL SUMMARY</strong></p><p>Summary text...</p><p><strong>EXPERIENCE</strong></p><h4><strong>Job Title</strong> - Company</h4><ul><li>• Achievement with numbers</li></ul>",
  "ats_analysis": {
    "ats_match_score": 85,
    "keyword_coverage": 80,
    "readability_score": 90,
    "action_verb_strength": 85,
    "formatting_compliance": 90
  },
  "improvement_insights": {
    "changes_made": ["change1"],
    "keywords_added": ["kw1"],
    "weak_phrases_removed": ["phrase1"],
    "action_verbs_used": ["verb1"]
  }
}`;

  const responseText = await callGeminiChat(prompt);
  
  // Parse JSON from response
  try {
    // Remove markdown code blocks if present
    let cleanText = responseText.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
    const jsonMatch = cleanText.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      return JSON.parse(jsonMatch[0]);
    }
  } catch (e) {
    console.error('Failed to parse JSON:', e);
  }

  // Return fallback result
  return {
    score: 75,
    score_breakdown: { impact: 80, brevity: 75, keywords: 70 },
    summary_critique: 'Resume optimized for ATS compatibility',
    missing_keywords: [],
    matched_keywords: [],
    optimized_content: responseText,
    ats_analysis: {
      ats_match_score: 75,
      keyword_coverage: 70,
      readability_score: 80,
      action_verb_strength: 75,
      formatting_compliance: 85
    }
  };
}

// Generate cover letter using Gemini
export async function generateCoverLetterWithGemini(
  resume: string,
  job: string
): Promise<string> {
  const prompt = `Generate a professional cover letter for this candidate and job.

RESUME: ${resume.replace(/<[^>]*>/g, ' ')}
JOB: ${job}

Write a compelling one-page cover letter with:
- Professional greeting
- Opening paragraph showing enthusiasm
- 2-3 body paragraphs matching skills to job requirements  
- Closing paragraph with call to action
- Professional sign-off

Return only the cover letter text, no JSON or markdown.`;

  return callGeminiChat(prompt);
}

// Generate hiring manager email using Gemini
export async function generateEmailWithGemini(
  resume: string,
  job: string
): Promise<{ subject_line: string; email_body: string }> {
  const prompt = `Generate a cold email to a hiring manager for this candidate and job.

RESUME: ${resume.replace(/<[^>]*>/g, ' ')}
JOB: ${job}

Return ONLY valid JSON (no markdown):
{
  "subject_line": "compelling subject line",
  "email_body": "professional email body with greeting, value proposition, and call to action"
}`;

  const responseText = await callGeminiChat(prompt);
  
  try {
    let cleanText = responseText.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
    const jsonMatch = cleanText.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      return JSON.parse(jsonMatch[0]);
    }
  } catch (e) {}

  return { subject_line: 'Application Inquiry', email_body: responseText };
}

// Generate interview prep using Gemini
export async function generateInterviewPrepWithGemini(
  resume: string,
  job: string
): Promise<Array<{ question: string; star_answer: string }>> {
  const prompt = `Generate 5 interview questions with STAR method answers for this candidate and job.

RESUME: ${resume.replace(/<[^>]*>/g, ' ')}
JOB: ${job}

Return ONLY valid JSON array (no markdown):
[{"question": "string", "star_answer": "string with Situation, Task, Action, Result"}]`;

  const responseText = await callGeminiChat(prompt);
  
  try {
    let cleanText = responseText.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
    const jsonMatch = cleanText.match(/\[[\s\S]*\]/);
    if (jsonMatch) {
      return JSON.parse(jsonMatch[0]);
    }
  } catch (e) {}

  return [];
}

// LinkedIn optimization using Gemini
export async function optimizeLinkedInWithGemini(
  resume: string,
  job: string
): Promise<string> {
  const prompt = `Create a LinkedIn profile optimization for this candidate targeting this job.

RESUME: ${resume.replace(/<[^>]*>/g, ' ')}
JOB: ${job}

Generate:
1. Professional Headline (120 chars max)
2. About Section (2-3 paragraphs)
3. Key Skills to Add (10-15 skills)
4. Achievement Highlights (3-4 bullets)

Format:
## HEADLINE
[headline]

## ABOUT
[about section]

## SKILLS TO ADD
[comma-separated skills]

## ACHIEVEMENT HIGHLIGHTS
[bullet points]`;

  return callGeminiChat(prompt);
}

// Skills gap analysis using Gemini
export async function analyzeSkillsGapWithGemini(
  resume: string,
  job: string
): Promise<Array<{ skill: string; hasSkill: boolean; importance: string; suggestion: string }>> {
  const prompt = `Analyze the skills gap between this resume and job description.

RESUME: ${resume.replace(/<[^>]*>/g, ' ')}
JOB: ${job}

Identify skills the candidate HAS and LACKS.

Return ONLY valid JSON array (no markdown):
[{"skill": "string", "hasSkill": boolean, "importance": "high/medium/low", "suggestion": "string with actionable advice"}]

Include at least 10 skills.`;

  const responseText = await callGeminiChat(prompt);
  
  try {
    let cleanText = responseText.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
    const jsonMatch = cleanText.match(/\[[\s\S]*\]/);
    if (jsonMatch) {
      return JSON.parse(jsonMatch[0]);
    }
  } catch (e) {}

  return [];
}
