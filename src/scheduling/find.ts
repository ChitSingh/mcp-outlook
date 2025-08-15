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
}
