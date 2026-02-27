import fs from 'fs';
import path from 'path';
import os from 'os';

// Configuration for ZAI API
interface ZAIConfig {
  baseUrl: string;
  apiKey: string;
}

// Get config from environment or write to file
export async function ensureZAIConfig(): Promise<ZAIConfig> {
  // Try to get config from environment variables first
  const envBaseUrl = process.env.ZAI_BASE_URL || process.env.NEXT_PUBLIC_ZAI_BASE_URL;
  const envApiKey = process.env.ZAI_API_KEY || process.env.NEXT_PUBLIC_ZAI_API_KEY;
  
  if (envBaseUrl && envApiKey) {
    // Write config file for SDK to use
    const config = { baseUrl: envBaseUrl, apiKey: envApiKey };
    const configPath = path.join(process.cwd(), '.z-ai-config');
    
    try {
      fs.writeFileSync(configPath, JSON.stringify(config));
    } catch (e) {
      // If write fails, we'll try direct API calls
      console.error('Could not write config file:', e);
    }
    
    return config;
  }
  
  // Default config for development/local environment
  const defaultConfig = {
    baseUrl: 'https://open.bigmodel.cn/api/paas/v4',
    apiKey: 'Z.ai'
  };
  
  // Try to write default config
  const configPath = path.join(process.cwd(), '.z-ai-config');
  try {
    fs.writeFileSync(configPath, JSON.stringify(defaultConfig));
  } catch (e) {
    console.error('Could not write default config:', e);
  }
  
  return defaultConfig;
}

// Direct API call helper (bypasses SDK config issues)
export async function callZAIChatAPI(
  messages: Array<{ role: string; content: string }>,
  options: { max_tokens?: number; thinking?: { type: 'enabled' | 'disabled' } } = {}
): Promise<string> {
  const config = await ensureZAIConfig();
  
  const response = await fetch(`${config.baseUrl}/chat/completions`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${config.apiKey}`,
    },
    body: JSON.stringify({
      messages,
      thinking: options.thinking || { type: 'disabled' },
      ...options
    })
  });
  
  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`API call failed: ${response.status} - ${errorText}`);
  }
  
  const data = await response.json();
  return data.choices?.[0]?.message?.content || '';
}

// Direct Vision API call helper
export async function callZAIVisionAPI(
  messages: Array<{
    role: string;
    content: string | Array<{ type: string; text?: string; image_url?: { url: string } }>;
  }>,
  model: string = 'glm-4.6v'
): Promise<string> {
  const config = await ensureZAIConfig();
  
  const response = await fetch(`${config.baseUrl}/chat/completions`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${config.apiKey}`,
    },
    body: JSON.stringify({
      model,
      messages,
      thinking: { type: 'disabled' }
    })
  });
  
  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Vision API call failed: ${response.status} - ${errorText}`);
  }
  
  const data = await response.json();
  return data.choices?.[0]?.message?.content || '';
}
