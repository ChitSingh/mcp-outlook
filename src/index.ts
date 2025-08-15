import { MCPServer } from './mcp.js';
import { logger } from './config.js';
import config from './config.js';

// Remove the local logger creation since we're using the file-based one from config
// const logger = pino({ level: config.logging.level });

async function main() {
  try {
    logger.info('Starting MCP Outlook Scheduler...');
    logger.info(`Configuration: ${JSON.stringify({
      tenantId: config.graph.tenantId,
      authMode: config.graph.authMode,
      organizer: config.organizer.email,
      timezone: config.server.defaultTimezone
    }, null, 2)}`);

    const server = new MCPServer();
    await server.start();

    logger.info('MCP server is running. Press Ctrl+C to stop.');
    
    // Keep the process alive
    process.on('SIGINT', () => {
      logger.info('Received SIGINT, shutting down gracefully...');
      process.exit(0);
    });

    process.on('SIGTERM', () => {
      logger.info('Received SIGTERM, shutting down gracefully...');
      process.exit(0);
    });

  } catch (error) {
    logger.error('Failed to start MCP server:', error);
    process.exit(1);
  }
}

// Handle uncaught errors
process.on('uncaughtException', (error) => {
  logger.error('Uncaught exception:', error);
  process.exit(1);
});

process.on('unhandledRejection', (reason, promise) => {
  logger.error('Unhandled rejection at:', promise, 'reason:', reason);
  process.exit(1);
});

main().catch((error) => {
  logger.error('Main function failed:', error);
  process.exit(1);
});
