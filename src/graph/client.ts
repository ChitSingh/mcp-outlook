import { Client } from '@microsoft/microsoft-graph-client';
import { GraphAuth } from './auth.js';
import { logger } from '../config.js';
import config from '../config.js';

// Remove the local logger creation since we're using the file-based one from config
// const logger = pino({ level: config.logging.level });

export class GraphClientFactory {
  private auth: GraphAuth;

  constructor(auth?: GraphAuth) {
    this.auth = auth || new GraphAuth();
  }

  /**
   * Get an authenticated Graph client
   */
  async getClient(): Promise<Client> {
    try {
      return await this.auth.getClient();
    } catch (error) {
      logger.error('Failed to get Graph client:', error);
      throw error;
    }
  }

  /**
   * Get current authentication scopes
   */
  getScopes(): string[] {
    return this.auth.getScopes();
  }

  /**
   * Clear authentication cache
   */
  clearAuth(): void {
    this.auth.clearTokenCache();
  }

  /**
   * Execute a Graph API call with retry logic
   */
  async executeWithRetry<T>(
    operation: () => Promise<T>,
    maxRetries: number = 3
  ): Promise<T> {
    let lastError: Error | null = null;
    
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        return await operation();
      } catch (error) {
        lastError = error instanceof Error ? error : new Error(String(error));
        
        if (attempt === maxRetries) {
          break;
        }

        // Check if we should retry
        if (this.shouldRetry(lastError)) {
          const delay = this.calculateRetryDelay(attempt, lastError);
          logger.warn(`Graph API call failed, retrying in ${delay}ms (attempt ${attempt}/${maxRetries})`);
          
          await this.sleep(delay);
          continue;
        } else {
          // Don't retry for client errors
          break;
        }
      }
    }

    throw lastError || new Error('Graph API call failed after retries');
  }

  /**
   * Determine if an error should trigger a retry
   */
  private shouldRetry(error: Error): boolean {
    const retryableErrors = [
      '429', // Too Many Requests
      '500', // Internal Server Error
      '502', // Bad Gateway
      '503', // Service Unavailable
      '504'  // Gateway Timeout
    ];

    return retryableErrors.some(code => error.message.includes(code));
  }

  /**
   * Calculate retry delay with exponential backoff
   */
  private calculateRetryDelay(attempt: number, error: Error): number {
    let baseDelay = 1000; // 1 second base
    
            // If we have a Retry-After header, use it
        if (error.message.includes('429')) {
          const retryAfterMatch = error.message.match(/Retry-After:\s*(\d+)/i);
          if (retryAfterMatch && retryAfterMatch[1]) {
            baseDelay = parseInt(retryAfterMatch[1]) * 1000;
          }
        }
    
    // Exponential backoff: 2^attempt * baseDelay
    const exponentialDelay = Math.pow(2, attempt - 1) * baseDelay;
    
    // Cap at 30 seconds
    return Math.min(exponentialDelay, 30000);
  }

  /**
   * Sleep utility
   */
  private sleep(ms: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
}
