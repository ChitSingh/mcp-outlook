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
    logger.info('getDelegatedClient called - checking token cache...');
    
    // Check if we have a valid cached token first
    if (this.isTokenValid() && this.tokenCache) {
      logger.info('Using cached token for delegated authentication');
      logger.info(`Token expires at: ${new Date(this.tokenCache.expiresAt).toISOString()}`);
      return Client.init({
        authProvider: (done) => {
          done(null, this.tokenCache!.accessToken);
        }
      });
    }

    // Clear invalid token cache
    if (this.tokenCache && !this.isTokenValid()) {
      logger.info('Clearing expired token cache');
      logger.info(`Expired token was valid until: ${new Date(this.tokenCache.expiresAt).toISOString()}`);
      this.tokenCache = null;
      this.saveTokenCache();
    }

    logger.info('No valid cached token found, creating new device code credential...');
    
    if (!this.deviceCodeCredential) {
      logger.info(`Creating DeviceCodeCredential with tenantId: ${config.graph.tenantId}, clientId: ${config.graph.clientId}`);
      this.deviceCodeCredential = new DeviceCodeCredential({
        tenantId: config.graph.tenantId,
        clientId: config.graph.clientId,
        userPromptCallback: (info) => {
          logger.info(`Please visit ${info.verificationUri} and enter code: ${info.userCode}`);
        }
      });
    }

    try {
      logger.info('Requesting new delegated token from Microsoft Graph...');
      
      // Add timeout to prevent hanging
      const tokenPromise = this.deviceCodeCredential.getToken([
        'https://graph.microsoft.com/Calendars.Read',
        'https://graph.microsoft.com/Calendars.ReadWrite',
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/OnlineMeetings.ReadWrite'
      ]);
      
      // Set a 2-minute timeout for device code authentication
      const timeoutPromise = new Promise((_, reject) => {
        setTimeout(() => reject(new Error('Device code authentication timed out after 2 minutes')), 2 * 60 * 1000);
      });
      
      const token = await Promise.race([tokenPromise, timeoutPromise]) as any;

      if (!token) {
        throw new Error('Failed to get access token');
      }

      logger.info(`Token received! Expires at: ${new Date(token.expiresOnTimestamp).toISOString()}`);
      logger.info(`Raw expiresOnTimestamp value: ${token.expiresOnTimestamp}`);
      logger.info(`Current time: ${new Date().toISOString()}`);
      logger.info(`Time difference: ${token.expiresOnTimestamp - Date.now()}ms`);
      
      // Debug: Check what scopes are actually in the token
      if (token.scopes) {
        logger.info(`Token scopes received: ${JSON.stringify(token.scopes)}`);
      } else {
        logger.warn('No scopes found in token response');
      }
      
      // Debug: Log the entire token response to see what's available
      logger.info('Full token response keys:', Object.keys(token));
      if (token.scopes) {
        logger.info('Token.scopes type:', typeof token.scopes);
        logger.info('Token.scopes length:', token.scopes.length);
      }
      
      // Update token cache
      this.tokenCache = {
        accessToken: token.token,
        expiresAt: token.expiresOnTimestamp,
        scopes: token.scopes || []
      };
      
      // If no scopes in token response, use the requested scopes
      if (!this.tokenCache.scopes || this.tokenCache.scopes.length === 0) {
        this.tokenCache.scopes = [
          'https://graph.microsoft.com/Calendars.Read',
          'https://graph.microsoft.com/Calendars.ReadWrite',
          'https://graph.microsoft.com/User.Read',
          'https://graph.microsoft.com/OnlineMeetings.ReadWrite'
        ];
        logger.info('Using requested scopes as fallback since token response had no scopes');
      }
      
      logger.info('Saving token to cache...');
      this.saveTokenCache();
      
      logger.info('Successfully obtained and cached new delegated token');

      return Client.init({
        authProvider: (done) => {
          logger.info(`AuthProvider called - providing token: ${token.token.substring(0, 20)}...`);
          done(null, token.token);
        }
      });
    } catch (error) {
      logger.error('Failed to get delegated token. Full error details:', error);
      if (error instanceof Error) {
        logger.error('Error name:', error.name);
        logger.error('Error message:', error.message);
        logger.error('Error stack:', error.stack);
      }
      throw new Error(`Authentication failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  }

  /**
   * Get application client (service principal authentication)
   */
  private async getApplicationClient(): Promise<Client> {
    logger.info('getApplicationClient called - attempting application authentication...');
    
    if (!config.graph.clientSecret) {
      const error = 'Client secret is required for application authentication';
      logger.error(error);
      throw new Error(error);
    }

    logger.info(`Using client ID: ${config.graph.clientId}, tenant ID: ${config.graph.tenantId}`);

    if (!this.msalApp) {
      logger.info('Creating new ConfidentialClientApplication...');
      this.msalApp = new ConfidentialClientApplication({
        auth: {
          clientId: config.graph.clientId,
          clientSecret: config.graph.clientSecret,
          authority: `https://login.microsoftonline.com/${config.graph.tenantId}`
        }
      });
    }

    try {
      logger.info('Requesting application token with scope: https://graph.microsoft.com/.default');
      const result = await this.msalApp.acquireTokenByClientCredential({
        scopes: ['https://graph.microsoft.com/.default']
      });

      if (!result) {
        throw new Error('Failed to acquire application token - no result returned');
      }

      if (!result.expiresOn) {
        throw new Error('Application token acquired but expiresOn is null');
      }

      logger.info(`Application token acquired successfully! Expires at: ${new Date(result.expiresOn).toISOString()}`);
      
      // Cache the application token
      this.tokenCache = {
        accessToken: result.accessToken,
        expiresAt: result.expiresOn.getTime(),
        scopes: ['https://graph.microsoft.com/.default']
      };
      this.saveTokenCache();

      return Client.init({
        authProvider: (done) => {
          done(null, result.accessToken);
        }
      });
    } catch (error) {
      logger.error('Failed to get application token. Full error details:', error);
      logger.error('Error type:', typeof error);
      logger.error('Error constructor:', error?.constructor?.name);
      
      if (error instanceof Error) {
        logger.error('Error name:', error.name);
        logger.error('Error message:', error.message);
        logger.error('Error stack:', error.stack);
      } else {
        logger.error('Raw error object:', JSON.stringify(error, null, 2));
        logger.error('Error toString():', String(error));
      }
      
      throw new Error(`Application authentication failed: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  /**
   * Check if current token is valid
   */
  private isTokenValid(): boolean {
    if (!this.tokenCache) return false;
    
    const now = Date.now();
    const buffer = 10 * 60 * 1000; // 10 minutes buffer to allow for refresh
    
    return this.tokenCache.expiresAt > (now + buffer);
  }

  /**
   * Check if token needs refresh (within buffer time)
   */
  private shouldRefreshToken(): boolean {
    if (!this.tokenCache) return false;
    
    const now = Date.now();
    const buffer = 5 * 60 * 1000; // 5 minutes buffer
    
    return this.tokenCache.expiresAt <= (now + buffer);
  }

  /**
   * Load token cache from disk
   */
  private loadTokenCache(): void {
    try {
      const cachePath = join(process.cwd(), config.auth.tokenCachePath);
      logger.info(`Attempting to load token cache from: ${cachePath}`);
      
      if (existsSync(cachePath)) {
        const data = readFileSync(cachePath, 'utf8');
        if (data.trim()) {
          this.tokenCache = JSON.parse(data);
          logger.info('Token cache loaded from disk successfully');
          if (this.tokenCache) {
            logger.info(`Cached token expires at: ${new Date(this.tokenCache.expiresAt).toISOString()}`);
            
            // Check if token is still valid
            if (!this.isTokenValid()) {
              logger.info('Cached token is expired, clearing cache');
              this.tokenCache = null;
              this.saveTokenCache();
            } else {
              logger.info('Cached token is valid and will be used');
            }
          }
        } else {
          logger.info('Token cache file exists but is empty');
          this.tokenCache = null;
        }
      } else {
        logger.info('No token cache file found on disk');
        this.tokenCache = null;
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
      logger.info(`Attempting to save token cache to: ${cachePath}`);
      
      if (this.tokenCache) {
        const data = JSON.stringify(this.tokenCache, null, 2);
        writeFileSync(cachePath, data);
        logger.info('Token cache saved to disk successfully');
        logger.info(`Saved token expires at: ${new Date(this.tokenCache.expiresAt).toISOString()}`);
      } else if (existsSync(cachePath)) {
        // Remove cache file if no token
        writeFileSync(cachePath, '');
        logger.info('Token cache file cleared (no token to save)');
      } else {
        logger.info('No token cache file to clear');
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
