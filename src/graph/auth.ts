import { Client } from '@microsoft/microsoft-graph-client';
import { DeviceCodeCredential } from '@azure/identity';
import { ConfidentialClientApplication } from '@azure/msal-node';
import { readFileSync, writeFileSync, existsSync } from 'fs';
import { join } from 'path';
import { logger } from '../config.js';
import config from '../config.js';

export interface TokenCache {
  accessToken: string;
  refreshToken?: string;
  expiresAt: number;
  scopes: string[];
}

export class GraphAuth {
  private tokenCache: TokenCache | null = null;
  private msalApp: ConfidentialClientApplication | null = null;
  private deviceCodeCredential: DeviceCodeCredential | null = null;

  constructor() {
    this.loadTokenCache();
  }

  /**
   * Get an authenticated Graph client
   */
  async getClient(): Promise<Client> {
    if (config.graph.authMode === 'delegated') {
      return this.getDelegatedClient();
    } else {
      return this.getApplicationClient();
    }
  }

  /**
   * Get delegated client (user authentication)
   */
  private async getDelegatedClient(): Promise<Client> {
    if (!this.deviceCodeCredential) {
      this.deviceCodeCredential = new DeviceCodeCredential({
        tenantId: config.graph.tenantId,
        clientId: config.graph.clientId,
        userPromptCallback: (info) => {
          logger.info(`Please visit ${info.verificationUri} and enter code: ${info.userCode}`);
        }
      });
    }

    try {
      const token = await this.deviceCodeCredential.getToken([
        'https://graph.microsoft.com/Calendars.Read',
        'https://graph.microsoft.com/Calendars.ReadWrite',
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/OnlineMeetings.ReadWrite'
      ]);

      if (!token) {
        throw new Error('Failed to get access token');
      }

      // Update token cache
      this.tokenCache = {
        accessToken: token.token,
        expiresAt: token.expiresOnTimestamp,
        scopes: [
          'https://graph.microsoft.com/Calendars.Read',
          'https://graph.microsoft.com/Calendars.ReadWrite',
          'https://graph.microsoft.com/User.Read',
          'https://graph.microsoft.com/OnlineMeetings.ReadWrite'
        ]
      };
      this.saveTokenCache();

      return Client.init({
        authProvider: (done) => {
          done(null, token.token);
        }
      });
    } catch (error) {
      logger.error('Failed to get delegated token:', error);
      throw new Error(`Authentication failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  }

  /**
   * Get application client (service principal authentication)
   */
  private async getApplicationClient(): Promise<Client> {
    if (!config.graph.clientSecret) {
      throw new Error('Client secret is required for application authentication');
    }

    if (!this.msalApp) {
      this.msalApp = new ConfidentialClientApplication({
        auth: {
          clientId: config.graph.clientId,
          clientSecret: config.graph.clientSecret,
          authority: `https://login.microsoftonline.com/${config.graph.tenantId}`
        }
      });
    }

    try {
      const result = await this.msalApp.acquireTokenByClientCredential({
        scopes: ['https://graph.microsoft.com/.default']
      });

      if (!result) {
        throw new Error('Failed to acquire application token');
      }

      return Client.init({
        authProvider: (done) => {
          done(null, result.accessToken);
        }
      });
    } catch (error) {
      logger.error('Failed to get application token:', error);
      throw new Error(`Application authentication failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  }

  /**
   * Check if current token is valid
   */
  private isTokenValid(): boolean {
    if (!this.tokenCache) return false;
    
    const now = Date.now();
    const buffer = 5 * 60 * 1000; // 5 minutes buffer
    
    return this.tokenCache.expiresAt > (now + buffer);
  }

  /**
   * Load token cache from disk
   */
  private loadTokenCache(): void {
    try {
      const cachePath = join(process.cwd(), config.auth.tokenCachePath);
      if (existsSync(cachePath)) {
        const data = readFileSync(cachePath, 'utf8');
        this.tokenCache = JSON.parse(data);
        
        // Check if token is still valid
        if (!this.isTokenValid()) {
          this.tokenCache = null;
          this.saveTokenCache();
        }
      }
    } catch (error) {
      logger.warn('Failed to load token cache:', error);
      this.tokenCache = null;
    }
  }

  /**
   * Save token cache to disk
   */
  private saveTokenCache(): void {
    try {
      const cachePath = join(process.cwd(), config.auth.tokenCachePath);
      if (this.tokenCache) {
        writeFileSync(cachePath, JSON.stringify(this.tokenCache, null, 2));
      } else if (existsSync(cachePath)) {
        // Remove cache file if no token
        writeFileSync(cachePath, '');
      }
    } catch (error) {
      logger.warn('Failed to save token cache:', error);
    }
  }

  /**
   * Clear token cache
   */
  clearTokenCache(): void {
    this.tokenCache = null;
    this.saveTokenCache();
  }

  /**
   * Get current scopes
   */
  getScopes(): string[] {
    return this.tokenCache?.scopes || [];
  }
}
