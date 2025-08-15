import { z } from 'zod';

// Test configuration schema
const TestConfigSchema = z.object({
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
    defaultTimezone: z.string().min(1).default('UTC'),
  }),
  auth: z.object({
    tokenCachePath: z.string().min(1).default('.tokens.json'),
  }),
  logging: z.object({
    level: z.enum(['fatal', 'error', 'warn', 'info', 'debug', 'trace']).default('info'),
  }),
});

// Test configuration with mock values
export const testConfig = TestConfigSchema.parse({
  graph: {
    tenantId: 'test-tenant-id',
    clientId: 'test-client-id',
    clientSecret: 'test-client-secret',
    authMode: 'delegated',
  },
  organizer: {
    email: 'test@example.com',
  },
  server: {
    port: 7337,
    defaultTimezone: 'UTC',
  },
  auth: {
    tokenCachePath: '.tokens.json',
  },
  logging: {
    level: 'info',
  },
});
