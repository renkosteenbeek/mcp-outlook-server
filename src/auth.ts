import { ConfidentialClientApplication, AuthorizationCodeRequest, AuthorizationUrlRequest } from '@azure/msal-node';
import { config } from './config.js';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

export class AuthManager {
  private msalClient: ConfidentialClientApplication;
  private tokenCachePath: string;

  constructor() {
    const msalConfig = {
      auth: {
        clientId: config.clientId,
        authority: `https://login.microsoftonline.com/${config.tenantId}`,
        clientSecret: config.clientSecret,
      },
    };

    this.msalClient = new ConfidentialClientApplication(msalConfig);
    this.tokenCachePath = path.join(__dirname, '../.tokens.json');
  }

  async getAuthUrl(): Promise<string> {
    const authCodeUrlParameters: AuthorizationUrlRequest = {
      scopes: [
        'User.Read',
        'Mail.Read',
        'Mail.ReadWrite',
        'Mail.Send',
        'Calendars.Read',
        'Calendars.ReadWrite',
      ],
      redirectUri: config.redirectUri,
    };

    const response = await this.msalClient.getAuthCodeUrl(authCodeUrlParameters);
    return response;
  }

  async acquireTokenByCode(code: string): Promise<any> {
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
      redirectUri: config.redirectUri,
    };

    const response = await this.msalClient.acquireTokenByCode(tokenRequest);
    
    // Save token to cache
    this.saveTokenToCache(response);
    
    return response;
  }

  async getAccessToken(): Promise<string | null> {
    // Try to get token from cache first
    const cachedToken = this.loadTokenFromCache();
    
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

        const response = await this.msalClient.acquireTokenSilent(silentRequest);
        this.saveTokenToCache(response);
        return response.accessToken;
      } catch (error) {
        console.error('Failed to refresh token:', error);
        return null;
      }
    }

    return null;
  }

  private saveTokenToCache(tokenResponse: any): void {
    try {
      fs.writeFileSync(this.tokenCachePath, JSON.stringify(tokenResponse, null, 2));
    } catch (error) {
      console.error('Failed to save token:', error);
    }
  }

  private loadTokenFromCache(): any {
    try {
      if (fs.existsSync(this.tokenCachePath)) {
        const data = fs.readFileSync(this.tokenCachePath, 'utf-8');
        return JSON.parse(data);
      }
    } catch (error) {
      console.error('Failed to load token:', error);
    }
    return null;
  }

  clearTokenCache(): void {
    try {
      if (fs.existsSync(this.tokenCachePath)) {
        fs.unlinkSync(this.tokenCachePath);
      }
    } catch (error) {
      console.error('Failed to clear token cache:', error);
    }
  }
}