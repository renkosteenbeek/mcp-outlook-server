import { ConfidentialClientApplication, AuthorizationCodeRequest, AuthorizationUrlRequest } from '@azure/msal-node';
import { config } from './config.js';
import { AccountConfig, AccountTokenCache } from './types.js';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

export class AuthManager {
  private msalClients: Map<string, ConfidentialClientApplication> = new Map();
  private tokenCachePath: string;
  private pendingAuth: Map<string, (code: string) => Promise<void>> = new Map();

  constructor() {
    this.tokenCachePath = path.join(__dirname, '../.tokens.json');
    
    // Initialize MSAL clients for each account
    for (const account of config.accounts) {
      const msalConfig = {
        auth: {
          clientId: account.clientId,
          authority: `https://login.microsoftonline.com/${account.tenantId}`,
          clientSecret: account.clientSecret,
        },
      };
      
      this.msalClients.set(account.name, new ConfidentialClientApplication(msalConfig));
    }
  }

  async getAuthUrl(accountName: string): Promise<string> {
    const client = this.msalClients.get(accountName);
    if (!client) {
      throw new Error(`Account ${accountName} not found`);
    }

    const authCodeUrlParameters: AuthorizationUrlRequest = {
      scopes: [
        'User.Read',
        'Mail.Read',
        'Mail.ReadWrite',
        'Mail.Send',
        'Calendars.Read',
        'Calendars.ReadWrite',
      ],
      redirectUri: config.server.redirectUri,
      state: accountName, // Use state to identify which account is authenticating
    };

    const response = await client.getAuthCodeUrl(authCodeUrlParameters);
    return response;
  }

  async acquireTokenByCode(code: string, accountName: string): Promise<any> {
    const client = this.msalClients.get(accountName);
    if (!client) {
      throw new Error(`Account ${accountName} not found`);
    }

    const tokenRequest: AuthorizationCodeRequest = {
      code: code,
      scopes: [
        'User.Read',
        'Mail.Read',
        'Mail.ReadWrite',
        'Mail.Send',
        'Calendars.Read',
        'Calendars.ReadWrite',
      ],
      redirectUri: config.server.redirectUri,
    };

    const response = await client.acquireTokenByCode(tokenRequest);
    
    // Save token to cache
    this.saveTokenToCache(accountName, response);
    
    return response;
  }

  async getAccessToken(accountName: string): Promise<string | null> {
    const client = this.msalClients.get(accountName);
    if (!client) {
      throw new Error(`Account ${accountName} not found`);
    }

    // Try to get token from cache first
    const cachedToken = this.loadTokenFromCache(accountName);
    
    if (cachedToken && new Date(cachedToken.expiresOn) > new Date()) {
      return cachedToken.accessToken;
    }

    // If we have a refresh token, try to use it
    if (cachedToken?.account) {
      try {
        const silentRequest = {
          account: cachedToken.account,
          scopes: [
            'User.Read',
            'Mail.Read',
            'Mail.ReadWrite',
            'Mail.Send',
            'Calendars.Read',
            'Calendars.ReadWrite',
          ],
        };

        const response = await client.acquireTokenSilent(silentRequest);
        this.saveTokenToCache(accountName, response);
        return response.accessToken;
      } catch (error) {
        console.error(`Failed to refresh token for ${accountName}:`, error);
        return null;
      }
    }

    return null;
  }

  getAllAccountNames(): string[] {
    return config.accounts.map(acc => acc.name);
  }

  getAccountConfig(accountName: string): AccountConfig | undefined {
    return config.accounts.find(acc => acc.name === accountName);
  }

  setPendingAuth(accountName: string, callback: (code: string) => Promise<void>) {
    this.pendingAuth.set(accountName, callback);
  }

  async handleAuthCallback(code: string, state: string) {
    const callback = this.pendingAuth.get(state);
    if (callback) {
      await callback(code);
      this.pendingAuth.delete(state);
    }
  }

  private saveTokenToCache(accountName: string, tokenResponse: any): void {
    try {
      let cache: AccountTokenCache = {};
      
      if (fs.existsSync(this.tokenCachePath)) {
        const data = fs.readFileSync(this.tokenCachePath, 'utf-8');
        cache = JSON.parse(data);
      }
      
      cache[accountName] = tokenResponse;
      
      fs.writeFileSync(this.tokenCachePath, JSON.stringify(cache, null, 2));
    } catch (error) {
      console.error(`Failed to save token for ${accountName}:`, error);
    }
  }

  private loadTokenFromCache(accountName: string): any {
    try {
      if (fs.existsSync(this.tokenCachePath)) {
        const data = fs.readFileSync(this.tokenCachePath, 'utf-8');
        const cache: AccountTokenCache = JSON.parse(data);
        return cache[accountName];
      }
    } catch (error) {
      console.error(`Failed to load token for ${accountName}:`, error);
    }
    return null;
  }

  clearTokenCache(accountName?: string): void {
    try {
      if (accountName) {
        // Clear specific account
        if (fs.existsSync(this.tokenCachePath)) {
          const data = fs.readFileSync(this.tokenCachePath, 'utf-8');
          const cache: AccountTokenCache = JSON.parse(data);
          delete cache[accountName];
          fs.writeFileSync(this.tokenCachePath, JSON.stringify(cache, null, 2));
        }
      } else {
        // Clear all
        if (fs.existsSync(this.tokenCachePath)) {
          fs.unlinkSync(this.tokenCachePath);
        }
      }
    } catch (error) {
      console.error('Failed to clear token cache:', error);
    }
  }
}