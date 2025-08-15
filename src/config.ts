import { z } from 'zod';
import dotenv from 'dotenv';
import pino from 'pino';

// Load environment variables
dotenv.config();

// Configuration schema
const ConfigSchema = z.object({
  graph: z.object({
    tenantId: z.string().min(1),
    clientId: z.string().min(1),
    clientSecret: z.string().optional(),
    authMode: z.enum(['delegated', 'app']),
  }),
  organizer: z.object({
    email: z.string().email(),
  }),
  server: z.object({
    port: z.coerce.number().int().positive().default(7337),
    defaultTimezone: z.string().min(1).default('Asia/Dubai'),
  }),
  auth: z.object({
    tokenCachePath: z.string().min(1).default('.tokens.json'),
  }),
  logging: z.object({
    level: z.enum(['fatal', 'error', 'warn', 'info', 'debug', 'trace']).default('info'),
  }),
});

// Parse and validate configuration
const config = ConfigSchema.parse({
  graph: {
    tenantId: process.env.GRAPH_TENANT_ID,
    clientId: process.env.GRAPH_CLIENT_ID,
    clientSecret: process.env.GRAPH_CLIENT_SECRET,
    authMode: process.env.GRAPH_AUTH_MODE,
  },
  organizer: {
    email: process.env.ORGANIZER_EMAIL,
  },
  server: {
    port: process.env.PORT,
    defaultTimezone: process.env.DEFAULT_TIMEZONE,
  },
  auth: {
    tokenCachePath: process.env.TOKEN_CACHE_PATH,
  },
  logging: {
    level: process.env.LOG_LEVEL,
  },
});

// Create logger that writes to file instead of stdout/stderr to avoid MCP communication issues
export const logger = pino({
  level: config.logging.level
}, pino.destination('./logs/mcp-server.log'));

export default config;
