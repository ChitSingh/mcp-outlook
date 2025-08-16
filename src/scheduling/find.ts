import { Client } from '@microsoft/microsoft-graph-client';
import { GraphClientFactory } from '../graph/client.js';
import { ProposeMeetingTimesInput, ProposeMeetingTimesOutput } from '../types.js';
import { AvailabilityService } from './availability.js';
import { IntersectionService } from './intersect.js';
import { logger } from '../config.js';
import config from '../config.js';

// Remove the local logger creation since we're using the file-based one from config
// const logger = pino({ level: config.logging.level });

export class FindMeetingTimesService {
  private graphClient: GraphClientFactory;
  private availabilityService: AvailabilityService;
  private intersectionService: IntersectionService;

  constructor(graphClient: GraphClientFactory) {
    this.graphClient = graphClient;
    this.availabilityService = new AvailabilityService(graphClient);
    this.intersectionService = new IntersectionService();
  }

  /**
   * Find meeting times using Microsoft Graph findMeetingTimes or fallback to local intersection
   */
  async findMeetingTimes(input: ProposeMeetingTimesInput): Promise<ProposeMeetingTimesOutput> {
    // Enhanced input validation with helpful error messages
    this.validateFindMeetingTimesInput(input);
    
    try {
      // Try Microsoft Graph findMeetingTimes first
      const graphResult = await this.tryGraphFindMeetingTimes(input);
      if (graphResult && graphResult.candidates.length > 0) {
        logger.info('Using Microsoft Graph findMeetingTimes result');
        return {
          source: 'graph_findMeetingTimes',
          candidates: graphResult.candidates
        };
      }
    } catch (error) {
      logger.warn('Graph findMeetingTimes failed, falling back to local intersection:', error);
    }

    // Fallback to local intersection logic
    logger.info('Using local intersection fallback');
    return await this.findMeetingTimesLocal(input);
  }

  /**
   * Try to use Microsoft Graph findMeetingTimes API
   */
  private async tryGraphFindMeetingTimes(input: ProposeMeetingTimesInput): Promise<ProposeMeetingTimesOutput | null> {
    const client = await this.graphClient.getClient();
    
    // Prepare attendees
    const attendees = input.participants.map(email => ({
      emailAddress: { address: email },
      type: 'required'
    }));

    // Set minimum required attendees if specified
    if (input.minRequiredAttendees && input.minRequiredAttendees < input.participants.length) {
      const requiredCount = input.minRequiredAttendees;
      attendees.forEach((attendee, index) => {
        attendee.type = index < requiredCount ? 'required' : 'optional';
      });
    }

    try {
      const result = await this.graphClient.executeWithRetry(() =>
        client.api('/me/findMeetingTimes')
          .post({
            attendees,
            timeConstraint: {
              activityDomain: 'work',
              timeSlots: [{
                start: {
                  dateTime: input.windowStart,
                  timeZone: input.timeZone || config.server.defaultTimezone
                },
                end: {
                  dateTime: input.windowEnd,
                  timeZone: input.timeZone || config.server.defaultTimezone
                }
              }]
            },
            meetingDuration: `PT${input.durationMinutes}M`,
            maxCandidates: input.maxCandidates,
            isOrganizerOptional: false,
            returnSuggestionReasons: true,
            minimumAttendeePercentage: input.minRequiredAttendees 
              ? Math.round((input.minRequiredAttendees / input.participants.length) * 100)
              : undefined
          })
      );

      if (result.meetingTimeSuggestions && result.meetingTimeSuggestions.length > 0) {
        const candidates = result.meetingTimeSuggestions.map((suggestion: any) => {
          const attendeeAvailability: Record<string, 'free' | 'tentative' | 'busy'> = {};
          
          // Process attendee availability
          if (suggestion.attendeeAvailability) {
            suggestion.attendeeAvailability.forEach((availability: any) => {
              const email = availability.attendee.emailAddress.address;
              let status: 'free' | 'tentative' | 'busy' = 'free';
              
              if (availability.availability === 'busy') {
                status = 'busy';
              } else if (availability.availability === 'tentative') {
                status = 'tentative';
              }
              
              attendeeAvailability[email] = status;
            });
          }

          // Calculate confidence based on availability
          const availableCount = Object.values(attendeeAvailability).filter(
            status => status === 'free' || status === 'tentative'
          ).length;
          const confidence = input.participants.length > 0 ? availableCount / input.participants.length : 0;

          return {
            start: suggestion.meetingTimeSlot.start.dateTime,
            end: suggestion.meetingTimeSlot.end.dateTime,
            attendeeAvailability,
            confidence
          };
        });

        return {
          source: 'graph_findMeetingTimes' as const,
          candidates
        };
      }
    } catch (error) {
      logger.error('Graph findMeetingTimes API call failed:', error);
      throw error;
    }

    return null;
  }

  /**
   * Fallback to local intersection logic
   */
  private async findMeetingTimesLocal(input: ProposeMeetingTimesInput): Promise<ProposeMeetingTimesOutput> {
    // Get availability for all participants
    const availability = await this.availabilityService.getAvailability({
      participants: input.participants,
      windowStart: input.windowStart,
      windowEnd: input.windowEnd,
      granularityMinutes: 30,
      workHoursOnly: input.workHoursOnly,
      timeZone: input.timeZone
    });

    // Find intersecting slots
    const candidates = this.intersectionService.findIntersectingSlots(input, availability.users.map(user => ({
      email: user.email,
      free: user.free,
      busy: user.busy
    })));

    // Filter by minimum attendance if specified
    const filteredCandidates = this.intersectionService.filterByMinimumAttendance(
      candidates,
      input.minRequiredAttendees
    );

    return {
      source: 'local_intersection',
      candidates: filteredCandidates
    };
  }

  /**
   * Check if Graph findMeetingTimes is available for the current tenant
   */
  async isGraphFindMeetingTimesAvailable(): Promise<boolean> {
    try {
      const client = await this.graphClient.getClient();
      await this.graphClient.executeWithRetry(() =>
        client.api('/me/findMeetingTimes')
          .post({
            attendees: [{ emailAddress: { address: 'test@example.com' }, type: 'required' }],
            timeConstraint: {
              activityDomain: 'work',
              timeSlots: [{
                start: { dateTime: new Date().toISOString(), timeZone: 'UTC' },
                end: { dateTime: new Date(Date.now() + 3600000).toISOString(), timeZone: 'UTC' }
              }]
            },
            meetingDuration: 'PT30M',
            maxCandidates: 1
          })
      );
      return true;
    } catch (error) {
      logger.debug('Graph findMeetingTimes not available:', error);
      return false;
    }
  }

  /**
   * Validate find meeting times input with helpful error messages
   */
  private validateFindMeetingTimesInput(input: ProposeMeetingTimesInput): void {
    const errors: string[] = [];

    // Check for participants
    if (!input.participants || input.participants.length === 0) {
      errors.push("At least one participant is required");
    }

    // Check for valid email formats
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    for (const email of input.participants) {
      if (!emailRegex.test(email)) {
        errors.push(`Invalid email format: ${email}`);
      }
    }

    // Check duration
    if (input.durationMinutes < 1 || input.durationMinutes > 1440) {
      errors.push("durationMinutes must be between 1 and 1440 minutes");
    }

    // Check for proper date format
    if (!input.windowStart.endsWith('.000Z') && !input.windowStart.includes('+') && !input.windowStart.includes('-')) {
      errors.push("windowStart must be in UTC format (.000Z) or include timezone offset (+/-HH:MM). Example: 2025-08-16T09:00:00+04:00 for Dubai time");
    }

    if (!input.windowEnd.endsWith('.000Z') && !input.windowEnd.includes('+') && !input.windowEnd.includes('-')) {
      errors.push("windowEnd must be in UTC format (.000Z) or include timezone offset (+/-HH:MM). Example: 2025-08-16T17:00:00+04:00 for Dubai time");
    }

    // Check for valid date range
    const startTime = new Date(input.windowStart);
    const endTime = new Date(input.windowEnd);
    if (startTime >= endTime) {
      errors.push("windowEnd must be after windowStart");
    }

    // Check for reasonable time window (not longer than 30 days)
    const durationMs = endTime.getTime() - startTime.getTime();
    const durationDays = durationMs / (1000 * 60 * 60 * 24);
    if (durationDays > 30) {
      errors.push("Time window cannot exceed 30 days");
    }

    // Check maxCandidates
    if (input.maxCandidates && (input.maxCandidates < 1 || input.maxCandidates > 20)) {
      errors.push("maxCandidates must be between 1 and 20");
    }

    // Check buffer times
    if (input.bufferBeforeMinutes && (input.bufferBeforeMinutes < 0 || input.bufferBeforeMinutes > 120)) {
      errors.push("bufferBeforeMinutes must be between 0 and 120");
    }

    if (input.bufferAfterMinutes && (input.bufferAfterMinutes < 0 || input.bufferAfterMinutes > 120)) {
      errors.push("bufferAfterMinutes must be between 0 and 120");
    }

    // Check minRequiredAttendees
    if (input.minRequiredAttendees && (input.minRequiredAttendees < 1 || input.minRequiredAttendees > input.participants.length)) {
      errors.push("minRequiredAttendees must be between 1 and the total number of participants");
    }

    if (errors.length > 0) {
      const errorMessage = `Find meeting times validation failed:\n${errors.map(err => `- ${err}`).join('\n')}\n\nCorrect format:\n` +
        `{\n` +
        `  "participants": ["email1@domain.com", "email2@domain.com"],\n` +
        `  "durationMinutes": 60,\n` +
        `  "windowStart": "2025-08-16T09:00:00+04:00",\n` +
        `  "windowEnd": "2025-08-16T17:00:00+04:00",\n` +
        `  "maxCandidates": 5,\n` +
        `  "bufferBeforeMinutes": 10,\n` +
        `  "bufferAfterMinutes": 10,\n` +
        `  "workHoursOnly": true,\n` +
        `  "minRequiredAttendees": 2,\n` +
        `  "timeZone": "Asia/Dubai"\n` +
        `}\n\nNote: Your email address is automatically set as the organizer from configuration.`;
      
      throw new Error(errorMessage);
    }
  }
}
