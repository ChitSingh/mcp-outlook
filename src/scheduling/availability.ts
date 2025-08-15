import { Client } from '@microsoft/microsoft-graph-client';
import { GraphClientFactory } from '../graph/client.js';
import { TimeInterval, WorkingHours, parseISOToUTC, roundToInterval, clipToWorkingHours, normalizeISOString } from '../utils/time.js';
import { GetAvailabilityInput, GetAvailabilityOutput, WorkingHours as WorkingHoursType } from '../types.js';
import { logger } from '../config.js';
import config from '../config.js';

export interface UserAvailability {
  email: string;
  workingHours: WorkingHours | null;
  busy: Array<{ start: string; end: string; subject?: string }>;
  free: Array<{ start: string; end: string }>;
}

export class AvailabilityService {
  private graphClient: GraphClientFactory;

  constructor() {
    this.graphClient = new GraphClientFactory();
  }

  /**
   * Get availability for multiple users
   */
  async getAvailability(input: GetAvailabilityInput): Promise<GetAvailabilityOutput> {
    const timezone = input.timeZone || config.server.defaultTimezone;
    const windowStart = parseISOToUTC(input.windowStart);
    const windowEnd = parseISOToUTC(input.windowEnd);
    
    logger.info(`Getting availability for ${input.participants.length} participants from ${windowStart.toISOString()} to ${windowEnd.toISOString()}`);

    const userAvailabilities: UserAvailability[] = [];

    for (const email of input.participants) {
      try {
        const availability = await this.getUserAvailability(
          email,
          windowStart,
          windowEnd,
          input.granularityMinutes,
          input.workHoursOnly,
          timezone
        );
        userAvailabilities.push(availability);
      } catch (error) {
        logger.error(`Failed to get availability for ${email}:`, error);
        // Add empty availability for failed users
        userAvailabilities.push({
          email,
          workingHours: null,
          busy: [],
          free: []
        });
      }
    }

    return {
      timeZone: timezone,
      users: userAvailabilities
    };
  }

  /**
   * Get availability for a single user
   */
  private async getUserAvailability(
    email: string,
    windowStart: Date,
    windowEnd: Date,
    granularityMinutes: number,
    workHoursOnly: boolean,
    timezone: string
  ): Promise<UserAvailability> {
    const client = await this.graphClient.getClient();
    
    // Get working hours
    const workingHours = await this.getWorkingHours(client, email, timezone);
    
    // Get free/busy information
    const freeBusy = await this.getFreeBusy(client, email, windowStart, windowEnd);
    
    // Process and normalize the data
    const processed = this.processFreeBusy(
      freeBusy,
      windowStart,
      windowEnd,
      granularityMinutes,
      workHoursOnly,
      workingHours,
      timezone
    );

    return {
      email,
      workingHours: workingHours ? this.convertToWorkingHoursType(workingHours) : null,
      busy: processed.busy,
      free: processed.free
    };
  }

  /**
   * Get working hours for a user
   */
  private async getWorkingHours(
    client: Client,
    email: string,
    timezone: string
  ): Promise<WorkingHours | null> {
    try {
      // Try to get mailbox settings
      const mailboxSettings = await this.graphClient.executeWithRetry(() =>
        client.api(`/users/${email}/mailboxSettings`).get()
      );

      if (mailboxSettings.workingHours) {
        const wh = mailboxSettings.workingHours;
        return {
          start: wh.startTime || '09:00',
          end: wh.endTime || '17:00',
          days: wh.daysOfWeek || [1, 2, 3, 4, 5] // Monday to Friday
        };
      }
    } catch (error) {
      logger.debug(`Could not get working hours for ${email}, using defaults`);
    }

    // Default working hours
    return {
      start: '09:00',
      end: '17:00',
      days: [1, 2, 3, 4, 5] // Monday to Friday
    };
  }

  /**
   * Get free/busy information for a user
   */
  private async getFreeBusy(
    client: Client,
    email: string,
    windowStart: Date,
    windowEnd: Date
  ): Promise<Array<{ start: string; end: string; availability: string; subject?: string }>> {
    try {
      const freeBusy = await this.graphClient.executeWithRetry(() =>
        client.api(`/users/${email}/calendar/getSchedule`)
          .post({
            schedules: [email],
            startTime: {
              dateTime: windowStart.toISOString(),
              timeZone: 'UTC'
            },
            endTime: {
              dateTime: windowEnd.toISOString(),
              timeZone: 'UTC'
            },
            availabilityViewInterval: 30
          })
      );

      if (freeBusy.value && freeBusy.value[0] && freeBusy.value[0].scheduleItems) {
        return freeBusy.value[0].scheduleItems.map((item: any) => ({
          start: item.start.dateTime,
          end: item.end.dateTime,
          availability: item.status || 'busy',
          subject: item.subject
        }));
      }
    } catch (error) {
      logger.warn(`Failed to get schedule for ${email}, trying free/busy:`, error);
      
      // Fallback to free/busy
      try {
        const freeBusy = await this.graphClient.executeWithRetry(() =>
          client.api(`/users/${email}/calendarView`)
            .query({
              startDateTime: windowStart.toISOString(),
              endDateTime: windowEnd.toISOString(),
              $select: 'start,end,subject,showAs'
            })
            .get()
        );

        return freeBusy.value.map((item: any) => ({
          start: item.start.dateTime,
          end: item.end.dateTime,
          availability: item.showAs || 'busy',
          subject: item.subject
        }));
      } catch (fallbackError) {
        logger.error(`Failed to get free/busy for ${email}:`, fallbackError);
        throw fallbackError;
      }
    }

    return [];
  }

  /**
   * Process free/busy data and generate free slots
   */
  private processFreeBusy(
    freeBusy: Array<{ start: string; end: string; availability: string; subject?: string }>,
    windowStart: Date,
    windowEnd: Date,
    granularityMinutes: number,
    workHoursOnly: boolean,
    workingHours: WorkingHours | null,
    timezone: string
  ): { busy: Array<{ start: string; end: string; subject?: string }>; free: Array<{ start: string; end: string }> } {
    // Sort by start time
    const sorted = freeBusy.sort((a, b) => new Date(a.start).getTime() - new Date(b.start).getTime());
    
    const busy: Array<{ start: string; end: string; subject?: string }> = [];
    const free: Array<{ start: string; end: string }> = [];
    
    let currentTime = new Date(windowStart);
    
    for (const item of sorted) {
      const itemStart = new Date(item.start);
      const itemEnd = new Date(item.end);
      
      // Add free time before this busy period
      if (currentTime < itemStart) {
        const freeStart = roundToInterval(currentTime, granularityMinutes);
        const freeEnd = roundToInterval(itemStart, granularityMinutes);
        
        if (freeStart < freeEnd) {
          const freeInterval = { start: freeStart, end: freeEnd };
          if (!workHoursOnly || !workingHours || clipToWorkingHours(freeInterval, workingHours, timezone)) {
            free.push({
              start: normalizeISOString(freeStart),
              end: normalizeISOString(freeEnd)
            });
          }
        }
      }
      
      // Add busy period
      if (item.availability === 'busy' || item.availability === 'tentative') {
        busy.push({
          start: normalizeISOString(itemStart),
          end: normalizeISOString(itemEnd),
          ...(item.subject && { subject: item.subject })
        });
      }
      
      currentTime = itemEnd;
    }
    
    // Add free time after last busy period
    if (currentTime < windowEnd) {
      const freeStart = roundToInterval(currentTime, granularityMinutes);
      const freeEnd = roundToInterval(windowEnd, granularityMinutes);
      
      if (freeStart < freeEnd) {
        const freeInterval = { start: freeStart, end: freeEnd };
        if (!workHoursOnly || !workingHours || clipToWorkingHours(freeInterval, workingHours, timezone)) {
          free.push({
            start: normalizeISOString(freeStart),
            end: normalizeISOString(freeEnd)
          });
        }
      }
    }
    
    return { busy, free };
  }

  /**
   * Convert internal WorkingHours to exported type
   */
  private convertToWorkingHoursType(wh: WorkingHours): WorkingHoursType {
    return {
      start: wh.start,
      end: wh.end,
      days: wh.days
    };
  }
}
