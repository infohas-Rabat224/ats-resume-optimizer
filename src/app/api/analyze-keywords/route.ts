import { NextRequest, NextResponse } from 'next/server';

// MODULE 3 — ATS KEYWORD ANALYZER
export async function POST(request: NextRequest) {
  try {
    const { jobDescription, resumeText } = await request.json();

    if (!jobDescription || !resumeText) {
      return NextResponse.json({ 
        success: false, 
        error: 'Both job description and resume text are required' 
      }, { status: 400 });
    }

    const analysis = analyzeKeywords(jobDescription, resumeText);
    return NextResponse.json({ success: true, data: analysis });
  } catch (error: any) {
    return NextResponse.json({ success: false, error: error.message }, { status: 500 });
  }
}

interface KeywordAnalysis {
  keyword: string;
  category: 'hard_skill' | 'soft_skill' | 'technical_tool' | 'industry_term';
  importance: 'high' | 'medium' | 'low';
  frequencyInJob: number;
  frequencyInResume: number;
  matchStatus: 'matched' | 'missing' | 'overused' | 'underused';
  suggestion: string;
}

interface KeywordAnalysisResult {
  totalKeywords: number;
  matchedKeywords: number;
  missingKeywords: number;
  matchPercentage: number;
  categories: {
    hardSkills: KeywordAnalysis[];
    softSkills: KeywordAnalysis[];
    technicalTools: KeywordAnalysis[];
    industryTerms: KeywordAnalysis[];
  };
  topMissing: string[];
  topMatched: string[];
  overusedKeywords: string[];
  keywordDensity: number;
  recommendations: string[];
}

// Categorization dictionaries
const HARD_SKILLS = [
  'project management', 'data analysis', 'financial analysis', 'budgeting', 'forecasting',
  'strategic planning', 'business development', 'sales', 'marketing', 'negotiation',
  'customer service', 'quality assurance', 'risk management', 'compliance', 'auditing',
  'recruitment', 'training', 'coaching', 'mentoring', 'presentation', 'reporting',
  'scheduling', 'coordination', 'administration', 'operations', 'logistics',
  'procurement', 'vendor management', 'contract negotiation', 'account management',
  'revenue management', 'p&l management', 'cost control', 'process improvement',
  'change management', 'crisis management', 'event planning', 'public relations'
];

const SOFT_SKILLS = [
  'leadership', 'communication', 'teamwork', 'collaboration', 'problem solving',
  'critical thinking', 'adaptability', 'flexibility', 'creativity', 'innovation',
  'time management', 'multitasking', 'attention to detail', 'organization', 'planning',
  'interpersonal', 'customer focus', 'empathy', 'patience', 'stress management',
  'conflict resolution', 'decision making', 'analytical', 'proactive', 'self-motivated',
  'reliable', 'punctual', 'professional', 'ethical', 'integrity', 'initiative',
  'delegation', 'motivational', 'persuasive', 'diplomatic'
];

const TECHNICAL_TOOLS = [
  'microsoft office', 'excel', 'word', 'powerpoint', 'outlook', 'teams', 'sharepoint',
  'sap', 'oracle', 'salesforce', 'hubspot', 'jira', 'asana', 'trello', 'monday',
  'tableau', 'power bi', 'sql', 'python', 'java', 'javascript', 'html', 'css',
  'photoshop', 'illustrator', 'figma', 'canva', 'adobe', 'google analytics',
  'crm', 'erp', 'pos', 'ats', 'hris', 'workday', 'taleo', 'successfactors',
  'aviation', 'gds', 'amadeus', 'sabre', 'galileo', 'dcs', 'aal', 'flight operations',
  'reservations', 'ticketing', 'check-in', 'boarding', 'passenger services'
];

const INDUSTRY_TERMS = [
  'airline', 'aviation', 'airport', 'passenger', 'flight', 'aircraft', 'crew',
  'cabin', 'cockpit', 'ground handling', 'ramp', 'terminal', 'gate', 'runway',
  'tower', 'atc', 'icao', 'iata', 'faa', 'easa', 'safety', 'security', 'sms',
  'hospitality', 'hotel', 'resort', 'restaurant', 'tourism', 'travel', 'booking',
  'accommodation', 'concierge', 'front desk', 'housekeeping', 'f&b', 'banquet',
  'events', 'conferences', 'meetings', 'luxury', 'premium', 'vip', 'first class',
  'business class', 'economy', 'loyalty', 'frequent flyer', 'miles', 'rewards',
  'emirates', 'qatar', 'etihad', 'lufthansa', 'british airways', 'delta', 'united'
];

function analyzeKeywords(jobDescription: string, resumeText: string): KeywordAnalysisResult {
  const jobLower = jobDescription.toLowerCase();
  const resumeLower = resumeText.toLowerCase();
  
  // Extract all potential keywords from job description
  const jobWords = jobLower.match(/\b[a-z]{3,}(?:\s+[a-z]{3,})?\b/g) || [];
  const jobWordFreq: Record<string, number> = {};
  
  jobWords.forEach(word => {
    const clean = word.trim();
    if (clean.length > 3) {
      jobWordFreq[clean] = (jobWordFreq[clean] || 0) + 1;
    }
  });

  // Get top 30 keywords by frequency
  const topKeywords = Object.entries(jobWordFreq)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 30);

  // Analyze each keyword
  const keywordAnalysis: KeywordAnalysis[] = topKeywords.map(([keyword, jobFreq]) => {
    const category = categorizeKeyword(keyword);
    const resumeFreq = countInText(keyword, resumeLower);
    const importance = jobFreq >= 5 ? 'high' : jobFreq >= 2 ? 'medium' : 'low';
    const matchStatus = getMatchStatus(jobFreq, resumeFreq);
    const suggestion = generateSuggestion(keyword, matchStatus, category);

    return {
      keyword,
      category,
      importance,
      frequencyInJob: jobFreq,
      frequencyInResume: resumeFreq,
      matchStatus,
      suggestion
    };
  });

  // Group by categories
  const categories = {
    hardSkills: keywordAnalysis.filter(k => k.category === 'hard_skill'),
    softSkills: keywordAnalysis.filter(k => k.category === 'soft_skill'),
    technicalTools: keywordAnalysis.filter(k => k.category === 'technical_tool'),
    industryTerms: keywordAnalysis.filter(k => k.category === 'industry_term')
  };

  // Calculate statistics
  const matched = keywordAnalysis.filter(k => k.matchStatus === 'matched').length;
  const missing = keywordAnalysis.filter(k => k.matchStatus === 'missing').length;
  const overused = keywordAnalysis.filter(k => k.matchStatus === 'overused');
  
  // Calculate keyword density
  const resumeWords = resumeLower.match(/\b[a-z]+\b/g) || [];
  const totalKeywordOccurrences = keywordAnalysis.reduce((sum, k) => sum + k.frequencyInResume, 0);
  const density = (totalKeywordOccurrences / resumeWords.length) * 100;

  return {
    totalKeywords: keywordAnalysis.length,
    matchedKeywords: matched,
    missingKeywords: missing,
    matchPercentage: Math.round((matched / keywordAnalysis.length) * 100),
    categories,
    topMissing: keywordAnalysis.filter(k => k.matchStatus === 'missing').slice(0, 5).map(k => k.keyword),
    topMatched: keywordAnalysis.filter(k => k.matchStatus === 'matched').slice(0, 5).map(k => k.keyword),
    overusedKeywords: overused.map(k => k.keyword),
    keywordDensity: Math.round(density * 100) / 100,
    recommendations: generateRecommendations(keywordAnalysis, matched, missing, density)
  };
}

function categorizeKeyword(keyword: string): KeywordAnalysis['category'] {
  const lower = keyword.toLowerCase();
  
  if (HARD_SKILLS.some(skill => lower.includes(skill) || skill.includes(lower))) {
    return 'hard_skill';
  }
  if (SOFT_SKILLS.some(skill => lower.includes(skill) || skill.includes(lower))) {
    return 'soft_skill';
  }
  if (TECHNICAL_TOOLS.some(tool => lower.includes(tool) || tool.includes(lower))) {
    return 'technical_tool';
  }
  return 'industry_term';
}

function countInText(keyword: string, text: string): number {
  const regex = new RegExp(`\\b${keyword.replace(/\s+/g, '\\s+')}\\b`, 'gi');
  const matches = text.match(regex);
  return matches ? matches.length : 0;
}

function getMatchStatus(jobFreq: number, resumeFreq: number): KeywordAnalysis['matchStatus'] {
  if (resumeFreq === 0) return 'missing';
  if (resumeFreq > jobFreq + 2) return 'overused';
  if (resumeFreq < Math.ceil(jobFreq / 2)) return 'underused';
  return 'matched';
}

function generateSuggestion(keyword: string, status: KeywordAnalysis['matchStatus'], category: KeywordAnalysis['category']): string {
  switch (status) {
    case 'missing':
      return `Add "${keyword}" to your ${category.replace('_', ' ')} section naturally`;
    case 'overused':
      return `Reduce usage of "${keyword}" - it appears too frequently`;
    case 'underused':
      return `Consider using "${keyword}" more prominently in relevant sections`;
    default:
      return `Good coverage of "${keyword}"`;
  }
}

function generateRecommendations(analysis: KeywordAnalysis[], matched: number, missing: number, density: number): string[] {
  const recommendations: string[] = [];
  
  if (missing > 10) {
    recommendations.push('Critical: More than 10 key keywords are missing. Significant revision needed.');
  } else if (missing > 5) {
    recommendations.push('Warning: Several important keywords are missing. Review and add them naturally.');
  } else if (missing > 0) {
    recommendations.push('Good keyword coverage. Consider adding the few missing keywords.');
  } else {
    recommendations.push('Excellent keyword coverage! All major keywords are present.');
  }

  if (density > 5) {
    recommendations.push('Keyword density is high. Ensure natural language flow.');
  } else if (density < 1) {
    recommendations.push('Keyword density is low. Consider more keyword integration.');
  }

  const missingHighImportance = analysis.filter(k => k.matchStatus === 'missing' && k.importance === 'high');
  if (missingHighImportance.length > 0) {
    recommendations.push(`Priority: Add these high-importance keywords: ${missingHighImportance.map(k => k.keyword).join(', ')}`);
  }

  return recommendations;
}
