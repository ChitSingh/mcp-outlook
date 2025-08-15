import { describe, it, expect, vi } from 'vitest';
import { IntersectionService } from '../src/scheduling/intersect.js';
import { ProposeMeetingTimesInput } from '../src/types.js';

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
      defaultTimezone: 'UTC'
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

describe('IntersectionService', () => {
  const service = new IntersectionService();

  describe('findIntersectingSlots', () => {
    it('should find intersecting slots for simple case', () => {
      const input: ProposeMeetingTimesInput = {
        participants: ['user1@example.com', 'user2@example.com'],
        durationMinutes: 30,
        windowStart: '2025-01-15T09:00:00Z',
        windowEnd: '2025-01-15T17:00:00Z',
        maxCandidates: 5,
        bufferBeforeMinutes: 0,
        bufferAfterMinutes: 0,
        workHoursOnly: false
      };

      const userAvailabilities = [
        {
          email: 'user1@example.com',
          free: [
            { start: '2025-01-15T09:00:00Z', end: '2025-01-15T12:00:00Z' },
            { start: '2025-01-15T14:00:00Z', end: '2025-01-15T17:00:00Z' }
          ],
          busy: [
            { start: '2025-01-15T12:00:00Z', end: '2025-01-15T14:00:00Z', subject: 'Lunch' }
          ]
        },
        {
          email: 'user2@example.com',
          free: [
            { start: '2025-01-15T10:00:00Z', end: '2025-01-15T11:00:00Z' },
            { start: '2025-01-15T15:00:00Z', end: '2025-01-15T16:00:00Z' }
          ],
          busy: [
            { start: '2025-01-15T11:00:00Z', end: '2025-01-15T15:00:00Z', subject: 'Meeting' }
          ]
        }
      ];

      const result = service.findIntersectingSlots(input, userAvailabilities);

      expect(result).toHaveLength(2);
      expect(result[0].start).toBe('2025-01-15T10:00:00Z');
      expect(result[0].end).toBe('2025-01-15T11:00:00Z');
      expect(result[1].start).toBe('2025-01-15T15:00:00Z');
      expect(result[1].end).toBe('2025-01-15T16:00:00Z');
    });

    it('should respect meeting duration constraints', () => {
      const input: ProposeMeetingTimesInput = {
        participants: ['user1@example.com', 'user2@example.com'],
        durationMinutes: 60,
        windowStart: '2025-01-15T09:00:00Z',
        windowEnd: '2025-01-15T17:00:00Z',
        maxCandidates: 5,
        bufferBeforeMinutes: 0,
        bufferAfterMinutes: 0,
        workHoursOnly: false
      };

      const userAvailabilities = [
        {
          email: 'user1@example.com',
          free: [
            { start: '2025-01-15T09:00:00Z', end: '2025-01-15T12:00:00Z' }
          ],
          busy: []
        },
        {
          email: 'user2@example.com',
          free: [
            { start: '2025-01-15T10:00:00Z', end: '2025-01-15T11:00:00Z' }
          ],
          busy: []
        }
      ];

      const result = service.findIntersectingSlots(input, userAvailabilities);

      // Should not return slots shorter than 60 minutes
      expect(result).toHaveLength(0);
    });

    it('should apply buffers correctly', () => {
      const input: ProposeMeetingTimesInput = {
        participants: ['user1@example.com', 'user2@example.com'],
        durationMinutes: 30,
        windowStart: '2025-01-15T09:00:00Z',
        windowEnd: '2025-01-15T17:00:00Z',
        maxCandidates: 5,
        bufferBeforeMinutes: 15,
        bufferAfterMinutes: 15,
        workHoursOnly: false
      };

      const userAvailabilities = [
        {
          email: 'user1@example.com',
          free: [
            { start: '2025-01-15T10:00:00Z', end: '2025-01-15T11:00:00Z' }
          ],
          busy: []
        },
        {
          email: 'user2@example.com',
          free: [
            { start: '2025-01-15T10:00:00Z', end: '2025-01-15T11:00:00Z' }
          ],
          busy: []
        }
      ];

      const result = service.findIntersectingSlots(input, userAvailabilities);

      // With 15 min buffers, the effective slot is 30 minutes (60 - 15 - 15)
      expect(result).toHaveLength(1);
      expect(result[0].start).toBe('2025-01-15T10:15:00Z');
      expect(result[0].end).toBe('2025-01-15T10:45:00Z');
    });

    it('should handle no intersections gracefully', () => {
      const input: ProposeMeetingTimesInput = {
        participants: ['user1@example.com', 'user2@example.com'],
        durationMinutes: 30,
        windowStart: '2025-01-15T09:00:00Z',
        windowEnd: '2025-01-15T17:00:00Z',
        maxCandidates: 5,
        bufferBeforeMinutes: 0,
        bufferAfterMinutes: 0,
        workHoursOnly: false
      };

      const userAvailabilities = [
        {
          email: 'user1@example.com',
          free: [
            { start: '2025-01-15T09:00:00Z', end: '2025-01-15T12:00:00Z' }
          ],
          busy: []
        },
        {
          email: 'user2@example.com',
          free: [
            { start: '2025-01-15T14:00:00Z', end: '2025-01-15T17:00:00Z' }
          ],
          busy: []
        }
      ];

      const result = service.findIntersectingSlots(input, userAvailabilities);

      expect(result).toHaveLength(0);
    });

    it('should respect maxCandidates limit', () => {
      const input: ProposeMeetingTimesInput = {
        participants: ['user1@example.com', 'user2@example.com'],
        durationMinutes: 30,
        windowStart: '2025-01-15T09:00:00Z',
        windowEnd: '2025-01-15T17:00:00Z',
        maxCandidates: 2,
        bufferBeforeMinutes: 0,
        bufferAfterMinutes: 0,
        workHoursOnly: false
      };

      const userAvailabilities = [
        {
          email: 'user1@example.com',
          free: [
            { start: '2025-01-15T09:00:00Z', end: '2025-01-15T10:00:00Z' },
            { start: '2025-01-15T11:00:00Z', end: '2025-01-15T12:00:00Z' },
            { start: '2025-01-15T13:00:00Z', end: '2025-01-15T14:00:00Z' }
          ],
          busy: []
        },
        {
          email: 'user2@example.com',
          free: [
            { start: '2025-01-15T09:00:00Z', end: '2025-01-15T10:00:00Z' },
            { start: '2025-01-15T11:00:00Z', end: '2025-01-15T12:00:00Z' },
            { start: '2025-01-15T13:00:00Z', end: '2025-01-15T14:00:00Z' }
          ],
          busy: []
        }
      ];

      const result = service.findIntersectingSlots(input, userAvailabilities);

      expect(result).toHaveLength(2);
    });
  });

  describe('checkMinimumAttendance', () => {
    it('should pass when minimum attendance is met', () => {
      const slot = {
        start: '2025-01-15T10:00:00Z',
        end: '2025-01-15T11:00:00Z',
        attendeeAvailability: {
          'user1@example.com': 'free' as const,
          'user2@example.com': 'free' as const,
          'user3@example.com': 'tentative' as const
        },
        confidence: 0.8
      };

      const result = service.checkMinimumAttendance(slot, 2);
      expect(result).toBe(true);
    });

    it('should fail when minimum attendance is not met', () => {
      const slot = {
        start: '2025-01-15T10:00:00Z',
        end: '2025-01-15T11:00:00Z',
        attendeeAvailability: {
          'user1@example.com': 'free' as const,
          'user2@example.com': 'busy' as const,
          'user3@example.com': 'busy' as const
        },
        confidence: 0.3
      };

      const result = service.checkMinimumAttendance(slot, 2);
      expect(result).toBe(false);
    });

    it('should pass when no minimum is specified', () => {
      const slot = {
        start: '2025-01-15T10:00:00Z',
        end: '2025-01-15T11:00:00Z',
        attendeeAvailability: {
          'user1@example.com': 'busy' as const
        },
        confidence: 0.0
      };

      const result = service.checkMinimumAttendance(slot, undefined);
      expect(result).toBe(true);
    });
  });
});
