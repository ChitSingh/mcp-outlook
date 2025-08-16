import { Client } from '@microsoft/microsoft-graph-client';
import { logger } from '../config.js';

export class GraphAuthWithToken {
  private accessToken: string | null = null;

  constructor() {
    // Check if we have a pre-generated token
    this.accessToken = process.env.GRAPH_ACCESS_TOKEN || null;
    
    if (this.accessToken) {
      logger.info('Using pre-generated access token from environment');
    } else {
      logger.warn('No GRAPH_ACCESS_TOKEN found in environment');
    }
  }

  /**
   * Get an authenticated Graph client
   */
  async getClient(): Promise<Client> {
    if (!this.accessToken) {
      throw new Error('No access token available. Set GRAPH_ACCESS_TOKEN environment variable or run authentication first.');
    }

    return Client.init({
      authProvider: (done) => {
        done(null, this.accessToken!);
      }
    });
  }

  /**
   * Check if we have a valid token
   */
  hasToken(): boolean {
    return !!this.accessToken;
  }

  /**
   * Get token info for debugging
   */
  getTokenInfo(): { hasToken: boolean; tokenPreview: string } {
    return {
      hasToken: !!this.accessToken,
      tokenPreview: this.accessToken ? `${this.accessToken.substring(0, 20)}...` : 'None'
    };
  }
}
