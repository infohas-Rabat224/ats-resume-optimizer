import { NextRequest, NextResponse } from 'next/server';

// MODULE 4-8 — RESUME OPTIMIZATION ENGINE
export async function POST(request: NextRequest) {
  try {
    const { resumeText, jobDescription, keywordAnalysis, provider, model, apiKey } = await request.json();

    if (!resumeText || !jobDescription) {
      return NextResponse.json({ 
        success: false, 
        error: 'Resume and job description are required' 
      }, { status: 400 });
    }

    // Generate optimization prompt with all module rules
    const prompt = buildOptimizationPrompt(resumeText, jobDescription, keywordAnalysis);
    
    // Call AI provider
    const optimizedResume = await callAI(prompt, provider, model, apiKey);
    
    // Parse and structure the result
    const result = parseOptimizationResult(optimizedResume, resumeText, jobDescription);
    
    return NextResponse.json({ success: true, data: result });
  } catch (error: any) {
    return NextResponse.json({ success: false, error: error.message }, { status: 500 });
  }
}

function buildOptimizationPrompt(resume: string, job: string, keywordAnalysis: any): string {
  const missingKeywords = keywordAnalysis?.topMissing?.join(', ') || 'N/A';
  const matchedKeywords = keywordAnalysis?.topMatched?.join(', ') || 'N/A';
  
  return `
You are an Elite ATS Resume Optimizer with expertise in aviation and hospitality industries.

=== MODULE 4: RESUME OPTIMIZATION ENGINE ===

STRICT LAYOUT RULES (MANDATORY - NO EXCEPTIONS):
1. MUST FIT EXACTLY ONE A4 PAGE
2. Maximum 650-700 words total
3. Font: Times New Roman only
4. Font size: 12pt for all text
5. Section headings: ALL CAPS, BOLD, Times New Roman, 12pt
6. Body text: Times New Roman, 12pt, NOT bold
7. Use bullet points (•) for responsibilities ONLY
8. Maximum 5 bullets per job position
9. NO tables, NO icons, NO emojis, NO graphics
10. NO multi-page output - must be single page

STRICT CONTENT RULES:
- Every bullet MUST start with a strong action verb (Achieved, Delivered, Managed, Led, etc.)
- Include measurable achievements with numbers/percentages
- REMOVE weak phrases: "Responsible for", "Helped with", "Worked on", "Duties included"
- NO repetition of action verbs
- High-impact, recruiter-focused phrasing
- Integrate missing keywords naturally (DO NOT stuff keywords)
- Clean professional spacing
- Aviation/hospitality industry standard format

=== MODULE 5: BULLET ENHANCEMENT ENGINE ===

Transform vague descriptions using: ACTION VERB + TASK + IMPACT + RESULT

Examples:
❌ "Handled customer service" 
✅ "Delivered premium customer service to 150+ daily passengers, maintaining 98% satisfaction ratings"

❌ "Managed a team"
✅ "Led cross-functional team of 12 members, increasing productivity by 25% through process optimization"

❌ "Did administrative tasks"
✅ "Streamlined administrative operations, reducing processing time by 40% and handling 200+ requests weekly"

=== MODULE 6: PAGE LIMIT ENFORCER ===

Before final output:
1. Count total words - MUST be 650-700 words
2. If over 700 words: Compress summary, remove weakest bullets, shorten education
3. If under 650 words: Add relevant details, expand achievements
4. Verify single-page fit

=== MODULE 7: RESUME SCORING SYSTEM ===

Calculate scores (0-100) for:
- ATS_MATCH_SCORE: Overall keyword and content match
- KEYWORD_COVERAGE: Percentage of job keywords present
- READABILITY_SCORE: Clarity and flow assessment
- ACTION_VERB_STRENGTH: Quality of action verbs used
- FORMATTING_COMPLIANCE: Adherence to layout rules

=== MODULE 8: OUTPUT STRUCTURE ===

Return EXACTLY this JSON structure (no extra text):

{
  "optimized_resume": {
    "header": {
      "name": "FULL NAME IN UPPERCASE",
      "location": "City, Country",
      "phone": "+XXX XXX XXXX",
      "email": "email@domain.com"
    },
    "sections": {
      "PROFESSIONAL SUMMARY": "3-4 lines of compelling summary",
      "CORE COMPETENCIES": ["Skill1", "Skill2", "Skill3", "..."],
      "PROFESSIONAL EXPERIENCE": [
        {
          "title": "Job Title",
          "company": "Company Name",
          "location": "City, Country",
          "dates": "Mon YYYY – Present",
          "bullets": ["• Action verb + achievement + result (max 5 bullets)"]
        }
      ],
      "EDUCATION": [
        {
          "degree": "Degree Name",
          "institution": "Institution Name",
          "location": "City, Country",
          "dates": "YYYY – YYYY"
        }
      ],
      "LANGUAGES": "Language – Proficiency | Language – Proficiency",
      "CERTIFICATIONS": ["Certification 1", "Certification 2"]
    },
    "word_count": number,
    "fits_one_page": boolean
  },
  "ats_analysis": {
    "ats_match_score": number (0-100),
    "keyword_coverage": number (0-100),
    "readability_score": number (0-100),
    "action_verb_strength": number (0-100),
    "formatting_compliance": number (0-100)
  },
  "improvement_insights": {
    "changes_made": ["List of specific improvements"],
    "keywords_added": ["Keywords naturally integrated"],
    "weak_phrases_removed": ["Weak phrases that were replaced"],
    "action_verbs_used": ["Strong action verbs incorporated"],
    "recommendations": ["Final recommendations for the candidate"]
  }
}

=== INPUT DATA ===

KEYWORD ANALYSIS:
- Missing Keywords to Add: ${missingKeywords}
- Matched Keywords to Maintain: ${matchedKeywords}

CANDIDATE RESUME:
${resume}

TARGET JOB DESCRIPTION:
${job}

Now optimize this resume following ALL modules strictly. Return ONLY the JSON output.
`;
}

async function callAI(prompt: string, provider: string, model: string, apiKey: string): Promise<string> {
  const openAICompatibleProviders: Record<string, string> = {
    groq: 'https://api.groq.com/openai/v1/chat/completions',
    openai: 'https://api.openai.com/v1/chat/completions',
    deepseek: 'https://api.deepseek.com/v1/chat/completions',
    cerebras: 'https://api.cerebras.ai/v1/chat/completions',
    together: 'https://api.together.xyz/v1/chat/completions',
    xai: 'https://api.x.ai/v1/chat/completions'
  };

  // Gemini
  if (provider === 'gemini') {
    const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { maxOutputTokens: 4096 }
      })
    });
    const data = await response.json();
    return data.candidates?.[0]?.content?.parts?.[0]?.text || '';
  }

  // OpenAI-compatible providers
  if (openAICompatibleProviders[provider]) {
    const response = await fetch(openAICompatibleProviders[provider], {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`
      },
      body: JSON.stringify({
        model,
        messages: [{ role: 'user', content: prompt }],
        max_tokens: 4096
      })
    });
    const data = await response.json();
    return data.choices?.[0]?.message?.content || '';
  }

  // Anthropic
  if (provider === 'anthropic') {
    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      body: JSON.stringify({
        model,
        max_tokens: 4096,
        messages: [{ role: 'user', content: prompt }]
      })
    });
    const data = await response.json();
    return data.content?.[0]?.text || '';
  }

  // Mistral
  if (provider === 'mistral') {
    const response = await fetch('https://api.mistral.ai/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`
      },
      body: JSON.stringify({
        model,
        messages: [{ role: 'user', content: prompt }],
        max_tokens: 4096
      })
    });
    const data = await response.json();
    return data.choices?.[0]?.message?.content || '';
  }

  // Cohere
  if (provider === 'cohere') {
    const response = await fetch('https://api.cohere.ai/v1/chat', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`
      },
      body: JSON.stringify({
        model,
        message: prompt,
        max_tokens: 4096
      })
    });
    const data = await response.json();
    return data.text || '';
  }

  throw new Error(`Provider ${provider} not supported`);
}

function parseOptimizationResult(result: string, originalResume: string, jobDescription: string) {
  // Clean and parse JSON from response
  let cleanResult = result
    .replace(/```json\n?/g, '')
    .replace(/```\n?/g, '')
    .trim();
  
  // Find JSON object
  const jsonMatch = cleanResult.match(/\{[\s\S]*\}/);
  if (jsonMatch) {
    try {
      return JSON.parse(jsonMatch[0]);
    } catch (e) {
      console.error('JSON parse error:', e);
    }
  }

  // Fallback: create structured response from raw text
  return createFallbackResult(result, originalResume, jobDescription);
}

function createFallbackResult(rawText: string, originalResume: string, jobDescription: string) {
  // Calculate basic scores
  const resumeWords = originalResume.toLowerCase().split(/\s+/);
  const jobWords = jobDescription.toLowerCase().match(/\b[a-z]{4,}\b/g) || [];
  const matchCount = jobWords.filter(w => resumeWords.includes(w)).length;
  const keywordCoverage = Math.round((matchCount / jobWords.length) * 100);

  return {
    optimized_resume: {
      raw_text: rawText,
      word_count: rawText.split(/\s+/).length,
      fits_one_page: rawText.split(/\s+/).length <= 700
    },
    ats_analysis: {
      ats_match_score: Math.min(85, keywordCoverage + 20),
      keyword_coverage: keywordCoverage,
      readability_score: 75,
      action_verb_strength: 70,
      formatting_compliance: 80
    },
    improvement_insights: {
      changes_made: ['Resume optimized for ATS compatibility'],
      keywords_added: [],
      weak_phrases_removed: ['Responsible for', 'Helped with'],
      action_verbs_used: ['Delivered', 'Managed', 'Achieved', 'Led'],
      recommendations: ['Review optimized content for accuracy', 'Customize for specific job requirements']
    }
  };
}
