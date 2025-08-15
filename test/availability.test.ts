import { describe, it, expect, vi, beforeEach } from 'vitest';
import { AvailabilityService } from '../src/scheduling/availability.js';
import { GetAvailabilityInput } from '../src/types.js';

// Mock the config module
vi.mock('../src/config.js', () => ({
  default: {
    graph: {
      tenantId: 'test-tenant-id',
      clientId: 'test-client-id',
      authMode: 'delegated'
    },
    organizer: {
      email: 'test@example.com'
    },
    server: {
      defaultTimezone: 'Asia/Dubai'
    },
    logging: {
      level: 'info'
    }
  },
  logger: {
    info: vi.fn(),
    error: vi.fn(),
    warn: vi.fn(),
    debug: vi.fn(),
    trace: vi.fn(),
    fatal: vi.fn()
  }
}));

// Mock the GraphClientFactory
vi.mock('../src/graph/client.js', () => ({
  GraphClientFactory: vi.fn().mockImplementation(() => ({
    getClient: vi.fn().mockResolvedValue({
      api: vi.fn().mockReturnValue({
        post: vi.fn().mockResolvedValue({
          value: [{
            scheduleItems: [
              {
                start: { dateTime: '2025-01-15T10:00:00Z' },
                end: { dateTime: '2025-01-15T11:00:00Z' },
                status: 'busy',
                subject: 'Test Meeting'
              }
            ]
          }]
        }),
        query: vi.fn().mockReturnValue({
          get: vi.fn().mockResolvedValue({
            value: [
              {
                start: { dateTime: '2025-01-15T10:00:00Z' },
                end: { dateTime: '2025-01-15T11:00:00Z' },
                subject: 'Test Meeting',
                showAs: 'busy'
              }
            ]
          })
        }),
        get: vi.fn().mockResolvedValue({
          workingHours: {
            startTime: '09:00',
            endTime: '17:00',
            daysOfWeek: [1, 2, 3, 4, 5]
          }
        })
      }),
      executeWithRetry: vi.fn().mockImplementation((fn) => fn())
    })
  }))
}));

describe('AvailabilityService', () => {
  let service: AvailabilityService;

  beforeEach(() => {
    service = new AvailabilityService();
  });

  describe('getAvailability', () => {
    it('should get availability for multiple participants', async () => {
      const input: GetAvailabilityInput = {
        participants: ['user1@example.com', 'user2@example.com'],
        windowStart: '2025-01-15T09:00:00Z',
        windowEnd: '2025-01-15T17:00:00Z',
        granularityMinutes: 30,
        workHoursOnly: true,
        timeZone: 'UTC'
      };

      const result = await service.getAvailability(input);

      expect(result.timeZone).toBe('UTC');
      expect(result.users).toHaveLength(2);
      expect(result.users[0].email).toBe('user1@example.com');
      expect(result.users[1].email).toBe('user2@example.com');
    });

    it('should handle failed availability requests gracefully', async () => {
      // Mock a failure for one user
      const mockService = new AvailabilityService();
      const mockGraphClient = {
        getClient: vi.fn().mockRejectedValue(new Error('API Error')),
        executeWithRetry: vi.fn().mockRejectedValue(new Error('API Error'))
      };

      // Replace the graphClient property
      (mockService as any).graphClient = mockGraphClient;

      const input: GetAvailabilityInput = {
        participants: ['user1@example.com', 'user2@example.com'],
        windowStart: '2025-01-15T09:00:00Z',
        windowEnd: '2025-01-15T17:00:00Z',
        granularityMinutes: 30,
        workHoursOnly: true,
        timeZone: 'UTC'
      };

      const result = await mockService.getAvailability(input);

      expect(result.users).toHaveLength(2);
      // Failed users should have empty availability
      expect(result.users[0].busy).toHaveLength(0);
      expect(result.users[0].free).toHaveLength(0);
      expect(result.users[1].busy).toHaveLength(0);
      expect(result.users[1].free).toHaveLength(0);
    });

    it('should use default timezone when none specified', async () => {
      const input: GetAvailabilityInput = {
        participants: ['user1@example.com'],
        windowStart: '2025-01-15T09:00:00Z',
        windowEnd: '2025-01-15T17:00:00Z',
        granularityMinutes: 30,
        workHoursOnly: true
        // timeZone not specified
      };

      const result = await service.getAvailability(input);

      expect(result.timeZone).toBe('Asia/Dubai'); // Default from config
    });

    it('should process free/busy data correctly', async () => {
      // Create a mock service and replace its GraphClientFactory
      const mockService = new AvailabilityService();
      
      // Create a more specific mock that handles the exact API calls
      const mockClient = {
        api: vi.fn().mockImplementation((path: string) => {
          if (path.includes('/calendar/getSchedule')) {
            return {
              post: vi.fn().mockResolvedValue({
                value: [{
                  scheduleItems: [
                    {
                      start: { dateTime: '2025-01-15T10:00:00Z' },
                      end: { dateTime: '2025-01-15T11:00:00Z' },
                      status: 'busy',
                      subject: 'Test Meeting'
                    }
                  ]
                }]
              })
            };
          } else if (path.includes('/calendarView')) {
            return {
              query: vi.fn().mockReturnValue({
                get: vi.fn().mockResolvedValue({
                  value: [
                    {
                      start: { dateTime: '2025-01-15T10:00:00Z' },
                      end: { dateTime: '2025-01-15T11:00:00Z' },
                      subject: 'Test Meeting',
                      showAs: 'busy'
                    }
                  ]
                })
              })
            };
          } else if (path.includes('/mailboxSettings')) {
            return {
              get: vi.fn().mockResolvedValue({
                workingHours: {
                  startTime: '09:00',
                  endTime: '17:00',
                  daysOfWeek: [1, 2, 3, 4, 5]
                }
              })
            };
          }
          return {
            post: vi.fn().mockRejectedValue(new Error('API not found')),
            query: vi.fn().mockReturnValue({
              get: vi.fn().mockRejectedValue(new Error('API not found'))
            }),
            get: vi.fn().mockRejectedValue(new Error('API not found'))
          };
        })
      };

      const mockGraphClient = {
        getClient: vi.fn().mockResolvedValue(mockClient),
        executeWithRetry: vi.fn().mockImplementation((fn) => fn())
      };

      // Replace the graphClient property
      (mockService as any).graphClient = mockGraphClient;
      
      const input: GetAvailabilityInput = {
        participants: ['user1@example.com'],
        windowStart: '2025-01-15T09:00:00Z',
        windowEnd: '2025-01-15T17:00:00Z',
        granularityMinutes: 30,
        workHoursOnly: false,
        timeZone: 'UTC'
      };

      const result = await mockService.getAvailability(input);

      expect(result.users[0].busy).toHaveLength(1);
      expect(result.users[0].busy[0].subject).toBe('Test Meeting');
      expect(result.users[0].free).toHaveLength(2); // Before and after the busy period
    });

    it('should respect granularity settings', async () => {
      const input: GetAvailabilityInput = {
        participants: ['user1@example.com'],
        windowStart: '2025-01-15T09:00:00Z',
        windowEnd: '2025-01-15T17:00:00Z',
        granularityMinutes: 60, // 1 hour granularity
        workHoursOnly: false,
        timeZone: 'UTC'
      };

      const result = await service.getAvailability(input);

      // With 1 hour granularity, times should be rounded
      result.users[0].free.forEach(slot => {
        const start = new Date(slot.start);
        const end = new Date(slot.end);
        expect(start.getMinutes()).toBe(0);
        expect(end.getMinutes()).toBe(0);
      });
    });
  });

  describe('working hours handling', () => {
    it('should get working hours from mailbox settings', async () => {
      const input: GetAvailabilityInput = {
        participants: ['user1@example.com'],
        windowStart: '2025-01-15T09:00:00Z',
        windowEnd: '2025-01-15T17:00:00Z',
        granularityMinutes: 30,
        workHoursOnly: true,
        timeZone: 'UTC'
      };

      const result = await service.getAvailability(input);

      expect(result.users[0].workingHours).toBeDefined();
      if (result.users[0].workingHours) {
        expect(result.users[0].workingHours.start).toBe('09:00');
        expect(result.users[0].workingHours.end).toBe('17:00');
        expect(result.users[0].workingHours.days).toEqual([1, 2, 3, 4, 5]);
      }
    });

    it('should use default working hours when mailbox settings unavailable', async () => {
      // Mock failure to get mailbox settings
      const mockService = new AvailabilityService();
      const mockGraphClient = {
        getClient: vi.fn().mockResolvedValue({
          api: vi.fn().mockReturnValue({
            post: vi.fn().mockResolvedValue({
              value: [{
                scheduleItems: []
              }]
            }),
            query: vi.fn().mockReturnValue({
              get: vi.fn().mockResolvedValue({
                value: []
              })
            }),
            get: vi.fn().mockRejectedValue(new Error('Mailbox settings not available'))
          }),
          executeWithRetry: vi.fn().mockImplementation((fn) => fn())
        })
      };

      (mockService as any).graphClient = mockGraphClient;

      const input: GetAvailabilityInput = {
        participants: ['user1@example.com'],
        windowStart: '2025-01-15T09:00:00Z',
        windowEnd: '2025-01-15T17:00:00Z',
        granularityMinutes: 30,
        workHoursOnly: true,
        timeZone: 'UTC'
      };

      const result = await mockService.getAvailability(input);

      // Should still have working hours (defaults)
      expect(result.users[0].workingHours).toBeDefined();
    });
  });
});
