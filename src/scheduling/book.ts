import { Client } from '@microsoft/microsoft-graph-client';
import { GraphClientFactory } from '../graph/client.js';
import { BookMeetingInput, BookMeetingOutput, CancelMeetingInput, CancelMeetingOutput } from '../types.js';
import { logger } from '../config.js';
import config from '../config.js';

export class BookingService {
  private graphClient: GraphClientFactory;

  constructor(graphClient: GraphClientFactory) {
    this.graphClient = graphClient;
  }

  /**
   * Book a meeting by creating a calendar event
   */
  async bookMeeting(input: BookMeetingInput): Promise<BookMeetingOutput> {
    const organizer = input.organizer || config.organizer.email;
    
    // Enhanced input validation with helpful error messages
    this.validateMeetingInput(input);
    
    logger.info(`Booking meeting: ${input.subject} from ${input.start} to ${input.end} with ${input.participants.length} participants`);

    const client = await this.graphClient.getClient();

    // Check for conflicts if not allowed
    if (!input.allowConflicts) {
      await this.checkForConflicts(client, input.participants, input.start, input.end);
    }

    // Create the event
    const event = await this.createEvent(client, input, organizer);

    // Send invitations
    await this.sendInvitations(client, event.id, input.participants, input.required, input.optional);

    return {
      eventId: event.id,
      iCalUid: event.iCalUId,
      webLink: event.webLink,
      organizer
    };
  }

  /**
   * Cancel a meeting
   */
  async cancelMeeting(input: CancelMeetingInput): Promise<CancelMeetingOutput> {
    const organizer = input.organizer || config.organizer.email;
    
    logger.info(`Cancelling meeting: ${input.eventId}`);

    const client = await this.graphClient.getClient();

    try {
      await this.graphClient.executeWithRetry(() =>
        client.api(`/users/${organizer}/events/${input.eventId}`)
          .delete()
      );

      return {
        cancelled: true,
        eventId: input.eventId
      };
    } catch (error) {
      logger.error('Failed to cancel meeting:', error);
      throw new Error(`Failed to cancel meeting: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  }

  /**
   * Create the calendar event
   */
  private async createEvent(
    client: Client,
    input: BookMeetingInput,
    organizer: string
  ): Promise<any> {
    const eventData: any = {
      subject: input.subject,
      start: {
        dateTime: input.start,
        timeZone: config.server.defaultTimezone
      },
      end: {
        dateTime: input.end,
        timeZone: config.server.defaultTimezone
      },
      attendees: [],
      isOnlineMeeting: input.onlineMeeting,
      onlineMeetingProvider: input.onlineMeeting ? 'teamsForBusiness' : undefined,
      location: input.location ? {
        displayName: input.location
      } : undefined,
      body: input.bodyHtml ? {
        contentType: 'HTML',
        content: input.bodyHtml
      } : undefined,
      reminderMinutesBeforeStart: input.remindersMinutesBeforeStart
    };

    // Set attendees based on input structure
    if (input.required || input.optional) {
      // Use required/optional structure if provided
      if (input.required) {
        eventData.attendees.push(...input.required.map(email => ({
          emailAddress: { address: email },
          type: 'required'
        })));
      }
      
      if (input.optional) {
        eventData.attendees.push(...input.optional.map(email => ({
          emailAddress: { address: email },
          type: 'optional'
        })));
      }
    } else {
      // Fall back to participants if no required/optional structure
      eventData.attendees = input.participants.map(email => ({
        emailAddress: { address: email },
        type: 'required'
      }));
    }

    try {
      const event = await this.graphClient.executeWithRetry(() =>
        client.api(`/users/${organizer}/events`)
          .post(eventData)
      );

      logger.info(`Created event with ID: ${event.id}`);
      return event;
    } catch (error) {
      logger.error('Failed to create event:', error);
      // Log detailed error information for debugging
      if (error instanceof Error) {
        logger.error(`Error name: ${error.name}`);
        logger.error(`Error message: ${error.message}`);
        logger.error(`Error stack: ${error.stack}`);
        // Check if it's a GraphError with additional properties
        if ('statusCode' in error) {
          logger.error(`Graph error status code: ${(error as any).statusCode}`);
        }
        if ('code' in error) {
          logger.error(`Graph error code: ${(error as any).code}`);
        }
        if ('body' in error) {
          logger.error(`Graph error body:`, (error as any).body);
        }
      } else {
        logger.error(`Unknown error type: ${typeof error}`);
        logger.error(`Raw error:`, error);
      }
      throw new Error(`Failed to create event: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  }

  /**
   * Check for scheduling conflicts
   */
  private async checkForConflicts(
    client: Client,
    participants: string[],
    start: string,
    end: string
  ): Promise<void> {
    for (const email of participants) {
      try {
        const calendarView = await this.graphClient.executeWithRetry(() =>
          client.api(`/users/${email}/calendarView`)
            .query({
              startDateTime: start,
              endDateTime: end,
              $select: 'subject,start,end'
            })
            .get()
        );

        if (calendarView.value && calendarView.value.length > 0) {
          const conflictSubjects = calendarView.value.map((c: any) => c.subject || 'Untitled');
          throw new Error(`Scheduling conflict for ${email}: ${conflictSubjects.join(', ')}`);
        }
      } catch (error) {
        if (error instanceof Error && error.message.includes('Scheduling conflict')) {
          throw error;
        }
        logger.warn(`Could not check conflicts for ${email}:`, error);
        // Log the full error for debugging
        if (error instanceof Error) {
          logger.warn(`Error details for ${email}: ${error.message}`);
          logger.warn(`Error stack for ${email}: ${error.stack}`);
          // Check if it's a GraphError with additional properties
          if ('statusCode' in error) {
            logger.warn(`Graph error status code for ${email}: ${(error as any).statusCode}`);
          }
          if ('code' in error) {
            logger.warn(`Graph error code for ${email}: ${(error as any).code}`);
          }
          if ('body' in error) {
            logger.warn(`Graph error body for ${email}:`, (error as any).body);
          }
        } else {
          logger.warn(`Unknown error type for ${email}:`, error);
        }
      }
    }
  }

  /**
   * Send meeting invitations
   */
  private async sendInvitations(
    client: Client,
    eventId: string,
    participants: string[],
    required?: string[],
    optional?: string[]
  ): Promise<void> {
    // For Microsoft Graph, invitations are automatically sent when creating the event
    // with attendees. We just need to ensure the event is properly created.
    logger.info(`Invitations will be sent automatically for event ${eventId}`);
  }

  /**
   * Update an existing meeting
   */
  async updateMeeting(
    eventId: string,
    updates: Partial<BookMeetingInput>,
    organizer?: string
  ): Promise<BookMeetingOutput> {
    const organizerEmail = organizer || config.organizer.email;
    
    logger.info(`Updating meeting: ${eventId}`);

    const client = await this.graphClient.getClient();

    try {
      const event = await this.graphClient.executeWithRetry(() =>
        client.api(`/users/${organizerEmail}/events/${eventId}`)
          .patch(updates)
      );

      return {
        eventId: event.id,
        iCalUid: event.iCalUId,
        webLink: event.webLink,
        organizer: organizerEmail
      };
    } catch (error) {
      logger.error('Failed to update meeting:', error);
      throw new Error(`Failed to update meeting: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  }

  /**
   * Get meeting details
   */
  async getMeeting(eventId: string, organizer?: string): Promise<any> {
    const organizerEmail = organizer || config.organizer.email;
    
    const client = await this.graphClient.getClient();

    try {
      const event = await this.graphClient.executeWithRetry(() =>
        client.api(`/users/${organizerEmail}/events/${eventId}`)
          .get()
      );

      return event;
    } catch (error) {
      logger.error('Failed to get meeting:', error);
      throw new Error(`Failed to get meeting: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  }

  /**
   * Validate meeting input with helpful error messages
   */
  private validateMeetingInput(input: BookMeetingInput): void {
    const errors: string[] = [];

    // Check for missing bodyHtml
    if (!input.bodyHtml || input.bodyHtml.trim() === '') {
      errors.push("Missing required field 'bodyHtml'. Please provide a meeting description.");
    }

    // Check for proper date format
    if (!input.start.endsWith('.000Z') && !input.start.includes('+') && !input.start.includes('-')) {
      errors.push("Start time must be in UTC format (.000Z) or include timezone offset (+/-HH:MM). Example: 2025-08-16T14:30:00+04:00 for Dubai time");
    }

    if (!input.end.endsWith('.000Z') && !input.end.includes('+') && !input.end.includes('-')) {
      errors.push("End time must be in UTC format (.000Z) or include timezone offset (+/-HH:MM). Example: 2025-08-16T15:00:00+04:00 for Dubai time");
    }

    // Check for valid date range
    const startTime = new Date(input.start);
    const endTime = new Date(input.end);
    if (startTime >= endTime) {
      errors.push("End time must be after start time");
    }

    // Check for reasonable meeting duration (not longer than 24 hours)
    const durationMs = endTime.getTime() - startTime.getTime();
    const durationHours = durationMs / (1000 * 60 * 60);
    if (durationHours > 24) {
      errors.push("Meeting duration cannot exceed 24 hours");
    }

    if (errors.length > 0) {
      const errorMessage = `Meeting validation failed:\n${errors.map(err => `- ${err}`).join('\n')}\n\nCorrect format:\n` +
        `{\n` +
        `  "start": "2025-08-16T14:30:00+04:00",\n` +
        `  "end": "2025-08-16T15:00:00+04:00",\n` +
        `  "subject": "Meeting Subject",\n` +
        `  "participants": ["email@domain.com"],\n` +
        `  "bodyHtml": "<p>Meeting description</p>",\n` +
        `  "onlineMeeting": true\n` +
        `}\n\nNote: Your email address is automatically set as the organizer from configuration.`;
      
      throw new Error(errorMessage);
    }
  }
}
