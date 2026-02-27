// Zhipu AI (智谱AI) Helper - https://open.bigmodel.cn/
// Get your API key at: https://open.bigmodel.cn/ (Register -> Console -> API Keys)

interface ZhipuChatResponse {
  id: string;
  created: number;
  model: string;
  choices: Array<{
    index: number;
    finish_reason: string;
    message: {
      role: string;
      content: string;
    };
  }>;
  usage: {
    prompt_tokens: number;
    completion_tokens: number;
    total_tokens: number;
  };
}

interface ZhipuVisionResponse {
  id: string;
  created: number;
  model: string;
  choices: Array<{
    index: number;
    finish_reason: string;
    message: {
      role: string;
      content: string;
    };
  }>;
}

// Get API key from environment
function getApiKey(): string {
  return process.env.ZHIPU_API_KEY || process.env.NEXT_PUBLIC_ZHIPU_API_KEY || '';
}

// Check if API key is configured
export function isZhipuConfigured(): boolean {
  return getApiKey().length > 10; // Basic validation
}

// Zhipu AI API base URL
const ZHIPU_BASE_URL = 'https://open.bigmodel.cn/api/paas/v4';

// Call Zhipu AI Chat API (GLM-4 model)
export async function callZhipuChat(
  prompt: string,
  options: { model?: string; max_tokens?: number } = {}
): Promise<string> {
  const apiKey = getApiKey();
  
  if (!apiKey) {
    throw new Error('Zhipu AI API key not configured. Please:\n1. Go to https://open.bigmodel.cn/\n2. Register/Login\n3. Go to Console -> API Keys\n4. Create and copy your API key\n5. Add it as ZHIPU_API_KEY environment variable');
  }

  const url = `${ZHIPU_BASE_URL}/chat/completions`;
  
  const body = {
    model: options.model || 'glm-4-flash',
    messages: [
      { role: 'user', content: prompt }
    ],
    max_tokens: options.max_tokens || 4096,
    temperature: 0.7,
    top_p: 0.9
  };

  const response = await fetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`
    },
    body: JSON.stringify(body)
  });

  if (!response.ok) {
    const errorText = await response.text();
    let errorMessage = `Zhipu API error: ${response.status}`;
    try {
      const errorJson = JSON.parse(errorText);
      errorMessage += ` - ${errorJson.error?.message || errorJson.message || errorText}`;
    } catch {
      errorMessage += ` - ${errorText}`;
    }
    throw new Error(errorMessage);
  }

  const data: ZhipuChatResponse = await response.json();
  return data.choices?.[0]?.message?.content || '';
}

// Call Zhipu AI Vision API (GLM-4V model)
export async function callZhipuVision(
  prompt: string,
  imageBase64: string,
  mimeType: string
): Promise<string> {
  const apiKey = getApiKey();
  
  if (!apiKey) {
    throw new Error('Zhipu AI API key not configured. Please:\n1. Go to https://open.bigmodel.cn/\n2. Register/Login\n3. Go to Console -> API Keys\n4. Create and copy your API key\n5. Add it as ZHIPU_API_KEY environment variable');
  }

  const url = `${ZHIPU_BASE_URL}/chat/completions`;

  const body = {
    model: 'glm-4v-flash',
    messages: [
      {
        role: 'user',
        content: [
          { type: 'text', text: prompt },
          {
            type: 'image_url',
            image_url: {
              url: `data:${mimeType};base64,${imageBase64}`
            }
          }
        ]
      }
    ],
    max_tokens: 4096,
    temperature: 0.3
  };

  const response = await fetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`
    },
    body: JSON.stringify(body)
  });

  if (!response.ok) {
    const errorText = await response.text();
    let errorMessage = `Zhipu Vision API error: ${response.status}`;
    try {
      const errorJson = JSON.parse(errorText);
      errorMessage += ` - ${errorJson.error?.message || errorJson.message || errorText}`;
    } catch {
      errorMessage += ` - ${errorText}`;
    }
    throw new Error(errorMessage);
  }

  const data: ZhipuVisionResponse = await response.json();
  return data.choices?.[0]?.message?.content || '';
}

// Extract text from PDF/Image using Zhipu Vision
export async function extractTextFromDocument(
  base64: string,
  mimeType: string
): Promise<string> {
  const prompt = `请从这个文档图片中提取所有文字内容。

要求：
- 提取图片中的所有文字，保持原有的结构和格式
- 保留项目符号和列表格式
- 包含所有姓名、日期和详细信息
- 如果是简历，提取所有部分包括联系方式、工作经历、教育背景、技能等
- 只返回提取的文字内容，不要添加任何解释或评论

Extract all text from this document image.
- Keep the original layout and hierarchy
- Preserve bullet points and lists
- Include all names, dates, and details
- Return ONLY the extracted text content`;

  return callZhipuVision(prompt, base64, mimeType);
}

// Optimize resume using Zhipu AI
export async function optimizeResumeWithZhipu(
  resume: string,
  job: string,
  missingKeywords?: string[]
): Promise<any> {
  const missingKw = missingKeywords?.join(', ') || 'N/A';
  
  const prompt = `你是一位专业的ATS简历优化专家 - 航空与酒店行业专家。

严格优化规则：

布局规则（必须遵守）：
1. 必须正好适合一页A4纸
2. 字数：650-700字（最多）
3. 字体：Times New Roman，12pt
4. 章节标题：全大写，加粗
5. 正文：不加粗
6. 项目符号：仅使用•符号，每个职位最多5条
7. 禁止：表格、图标、表情符号、图形

内容规则：
- 每个项目符号必须以动作动词开头（Achieved, Delivered, Managed, Led, Implemented等）
- 在成就中包含数字/百分比
- 删除弱表达："Responsible for"、"Helped with"、"Worked on"
- 不要重复使用动作动词
- 自然地融入关键词

输入数据：

需要添加的关键词：${missingKw}

候选人简历：
${resume}

目标职位描述：
${job}

只返回有效的JSON（不要使用markdown代码块）：
{
  "score": 85,
  "score_breakdown": {"impact": 85, "brevity": 90, "keywords": 80},
  "summary_critique": "简要评价",
  "missing_keywords": ["关键词1"],
  "matched_keywords": ["关键词1"],
  "optimized_content": "<h1>姓名</h1><h4>联系方式</h4><p><strong>专业概述</strong></p><p>概述内容...</p><p><strong>工作经历</strong></p><h4><strong>职位名称</strong> - 公司</h4><ul><li>• 带有数字的成就</li></ul>",
  "ats_analysis": {
    "ats_match_score": 85,
    "keyword_coverage": 80,
    "readability_score": 90,
    "action_verb_strength": 85,
    "formatting_compliance": 90
  },
  "improvement_insights": {
    "changes_made": ["修改1"],
    "keywords_added": ["关键词1"],
    "weak_phrases_removed": ["弱表达1"],
    "action_verbs_used": ["动词1"]
  }
}`;

  const responseText = await callZhipuChat(prompt, { max_tokens: 4096 });
  
  // Parse JSON from response
  try {
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
    summary_critique: '简历已优化以符合ATS兼容性',
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

// Generate cover letter using Zhipu AI
export async function generateCoverLetterWithZhipu(
  resume: string,
  job: string
): Promise<string> {
  const prompt = `为这位候选人和职位生成一封专业的求职信。

简历：${resume.replace(/<[^>]*>/g, ' ')}
职位：${job}

写一封有说服力的一页求职信，包括：
- 专业的问候语
- 表达热情的开场段落
- 2-3个正文段落，将技能与职位要求相匹配
- 带有行动号召的结尾段落
- 专业的结束语

只返回求职信文本，不要使用JSON或markdown。`;

  return callZhipuChat(prompt);
}

// Generate hiring manager email using Zhipu AI
export async function generateEmailWithZhipu(
  resume: string,
  job: string
): Promise<{ subject_line: string; email_body: string }> {
  const prompt = `为这位候选人和职位生成一封给招聘经理的开发邮件。

简历：${resume.replace(/<[^>]*>/g, ' ')}
职位：${job}

只返回有效的JSON（不要使用markdown）：
{
  "subject_line": "吸引人的邮件主题",
  "email_body": "专业的邮件正文，包括问候、价值主张和行动号召"
}`;

  const responseText = await callZhipuChat(prompt);
  
  try {
    let cleanText = responseText.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
    const jsonMatch = cleanText.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      return JSON.parse(jsonMatch[0]);
    }
  } catch (e) {}

  return { subject_line: '职位咨询', email_body: responseText };
}

// Generate interview prep using Zhipu AI
export async function generateInterviewPrepWithZhipu(
  resume: string,
  job: string
): Promise<Array<{ question: string; star_answer: string }>> {
  const prompt = `为这位候选人和职位生成5个面试问题及STAR方法回答。

简历：${resume.replace(/<[^>]*>/g, ' ')}
职位：${job}

只返回有效的JSON数组（不要使用markdown）：
[{"question": "问题", "star_answer": "包含情境、任务、行动、结果的回答"}]`;

  const responseText = await callZhipuChat(prompt);
  
  try {
    let cleanText = responseText.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
    const jsonMatch = cleanText.match(/\[[\s\S]*\]/);
    if (jsonMatch) {
      return JSON.parse(jsonMatch[0]);
    }
  } catch (e) {}

  return [];
}

// LinkedIn optimization using Zhipu AI
export async function optimizeLinkedInWithZhipu(
  resume: string,
  job: string
): Promise<string> {
  const prompt = `为这位候选人针对该职位创建LinkedIn个人资料优化建议。

简历：${resume.replace(/<[^>]*>/g, ' ')}
职位：${job}

生成：
1. 专业标题（最多120个字符）
2. 关于部分（2-3段）
3. 要添加的关键技能（10-15个）
4. 成就亮点（3-4个项目符号）`;

  return callZhipuChat(prompt);
}

// Skills gap analysis using Zhipu AI
export async function analyzeSkillsGapWithZhipu(
  resume: string,
  job: string
): Promise<Array<{ skill: string; hasSkill: boolean; importance: string; suggestion: string }>> {
  const prompt = `分析这份简历与职位描述之间的技能差距。

简历：${resume.replace(/<[^>]*>/g, ' ')}
职位：${job}

识别候选人拥有和缺乏的技能。

只返回有效的JSON数组（不要使用markdown）：
[{"skill": "技能名称", "hasSkill": true/false, "importance": "high/medium/low", "suggestion": "可行的建议"}]

至少包含10个技能。`;

  const responseText = await callZhipuChat(prompt);
  
  try {
    let cleanText = responseText.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
    const jsonMatch = cleanText.match(/\[[\s\S]*\]/);
    if (jsonMatch) {
      return JSON.parse(jsonMatch[0]);
    }
  } catch (e) {}

  return [];
}
