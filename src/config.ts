import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';
import { Config } from './types.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

function loadConfig(): Config {
  const configPath = path.join(__dirname, '../config.json');
  
  // Try to load from config.json first
  if (fs.existsSync(configPath)) {
    const configData = fs.readFileSync(configPath, 'utf-8');
    return JSON.parse(configData);
  }
  
  // Fallback to environment variables for single account (backward compatibility)
  if (process.env.CLIENT_ID && process.env.CLIENT_SECRET) {
    return {
      accounts: [{
        name: 'Default',
        tenantId: process.env.TENANT_ID || 'common',
        clientId: process.env.CLIENT_ID,
        clientSecret: process.env.CLIENT_SECRET,
      }],
      server: {
        port: parseInt(process.env.PORT || '3000'),
        redirectUri: process.env.REDIRECT_URI || 'http://localhost:3000/auth/callback',
      }
    };
  }
  
  throw new Error('No configuration found. Please create a config.json file or set environment variables.');
}

export const config = loadConfig();

// Validate configuration
if (!config.accounts || config.accounts.length === 0) {
  throw new Error('No accounts configured. Please add at least one account to config.json');
}

for (const account of config.accounts) {
  if (!account.name || !account.clientId || !account.clientSecret) {
    throw new Error(`Invalid account configuration for ${account.name || 'unnamed account'}`);
  }
}