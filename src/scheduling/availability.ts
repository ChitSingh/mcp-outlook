import { Client } from '@microsoft/microsoft-graph-client';
import { GraphClientFactory } from '../graph/client.js';
import { TimeInterval, WorkingHours, parseISOToUTC, roundToInterval, clipToWorkingHours, normalizeISOString, convertUTCToTimezone, parseGraphAPITime } from '../utils/time.js';
import { GetAvailabilityInput, GetAvailabilityOutput, WorkingHours as WorkingHoursType } from '../types.js';
import { logger } from '../config.js';
import config from '../config.js';

export interface UserAvailability {
  email: string;
  workingHours: WorkingHours | null;
  busy: Array<{ start: string; end: string }>;
  free: Array<{ start: string; end: string }>;
}

export class AvailabilityService {
  private graphClient: GraphClientFactory;

  constructor(graphClient: GraphClientFactory) {
    this.graphClient = graphClient;
  }

  /**
   * Get availability for multiple users
   */
  async getAvailability(input: GetAvailabilityInput): Promise<GetAvailabilityOutput> {
    // Enhanced input validation with helpful error messages
    this.validateAvailabilityInput(input);
    
    const timezone = input.timeZone || config.server.defaultTimezone;
    
    // Convert input times to UTC for the API call
    // The input is in Dubai timezone (+04:00), but API expects UTC
    const windowStart = this.convertInputTimeToUTC(input.windowStart, timezone);
    const windowEnd = this.convertInputTimeToUTC(input.windowEnd, timezone);
    
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
    const freeBusy = await this.getFreeBusy(client, email, windowStart, windowEnd, timezone);
    
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
    windowEnd: Date,
    timezone: string
  ): Promise<Array<{ start: string; end: string; availability: string }>> {
    try {
      logger.debug(`Getting schedule for ${email} from ${windowStart.toISOString()} to ${windowEnd.toISOString()}`);
      
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
        const items = freeBusy.value[0].scheduleItems;
        logger.debug(`Found ${items.length} schedule items for ${email}`);
        
        return items.map((item: any) => {
          const startTime = this.formatTimeForOutput(item.start.dateTime, timezone);
          const endTime = this.formatTimeForOutput(item.end.dateTime, timezone);
          
          logger.debug(`Schedule item for ${email}: ${startTime} to ${endTime} (${item.status})`);
          
          return {
            start: startTime,
            end: endTime,
            availability: item.status || 'busy'
          };
        });
      }
      
      logger.debug(`No schedule items found for ${email}`);
      return [];
    } catch (error) {
      logger.warn(`Failed to get schedule for ${email}, trying free/busy:`, error);
      
      // Fallback to free/busy
      try {
        logger.debug(`Falling back to calendarView for ${email}`);
        
        const freeBusy = await this.graphClient.executeWithRetry(() =>
          client.api(`/users/${email}/calendarView`)
            .query({
              startDateTime: windowStart.toISOString(),
              endDateTime: windowEnd.toISOString(),
              $select: 'start,end,subject,showAs'
            })
            .get()
        );

        if (freeBusy.value && freeBusy.value.length > 0) {
          logger.debug(`Found ${freeBusy.value.length} calendar events for ${email}`);
          
          return freeBusy.value.map((item: any) => {
            const startTime = this.formatTimeForOutput(item.start.dateTime, timezone);
            const endTime = this.formatTimeForOutput(item.end.dateTime, timezone);
            
            logger.debug(`Calendar event for ${email}: ${startTime} to ${endTime} (${item.showAs})`);
            
            return {
              start: startTime,
              end: endTime,
              availability: item.showAs || 'busy'
            };
          });
        }
        
        logger.debug(`No calendar events found for ${email}`);
        return [];
      } catch (fallbackError) {
        logger.error(`Failed to get free/busy for ${email}:`, fallbackError);
        throw fallbackError;
      }
    }
  }

  /**
   * Process free/busy data and generate free slots
   */
  private processFreeBusy(
    freeBusy: Array<{ start: string; end: string; availability: string }>,
    windowStart: Date,
    windowEnd: Date,
    granularityMinutes: number,
    workHoursOnly: boolean,
    workingHours: WorkingHours | null,
    timezone: string
  ): { busy: Array<{ start: string; end: string }>; free: Array<{ start: string; end: string }> } {
    logger.debug(`Processing free/busy data for ${freeBusy.length} items`);
    logger.debug(`Window: ${windowStart.toISOString()} to ${windowEnd.toISOString()}`);
    
    // Sort by start time - use the raw string values to avoid timezone interpretation issues
    const sorted = freeBusy.sort((a, b) => {
      // Extract time components for comparison without timezone interpretation
      const aMatch = a.start.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})/);
      const bMatch = b.start.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})/);
      
      if (aMatch && bMatch) {
        // Compare as strings to avoid timezone issues
        return a.start.localeCompare(b.start);
      }
      
      // Fallback to date comparison
      return new Date(a.start).getTime() - new Date(b.start).getTime();
    });
    
    const busy: Array<{ start: string; end: string }> = [];
    const free: Array<{ start: string; end: string }> = [];
    
    // Convert window boundaries to timezone-aware strings for consistent comparison
    const windowStartStr = convertUTCToTimezone(windowStart, timezone);
    const windowEndStr = convertUTCToTimezone(windowEnd, timezone);
    
    logger.debug(`Timezone-adjusted window: ${windowStartStr} to ${windowEndStr}`);
    
    // Parse the window boundaries as timezone-aware dates
    const windowStartDate = new Date(windowStartStr);
    const windowEndDate = new Date(windowEndStr);
    
    // Filter busy periods to only include those within the working window
    const relevantBusyPeriods = sorted.filter(item => {
      const itemStart = new Date(item.start);
      const itemEnd = new Date(item.end);
      
      // Check if the busy period overlaps with the working window
      const overlaps = itemStart < windowEndDate && itemEnd > windowStartDate;
      logger.debug(`Busy period ${item.start} to ${item.end} overlaps with window: ${overlaps}`);
      
      if (overlaps && (item.availability === 'busy' || item.availability === 'tentative')) {
        // Add to busy periods, but clip to window boundaries
        const clippedStart = itemStart < windowStartDate ? windowStartDate : itemStart;
        const clippedEnd = itemEnd > windowEndDate ? windowEndDate : itemEnd;
        
        busy.push({
          start: convertUTCToTimezone(clippedStart, timezone),
          end: convertUTCToTimezone(clippedEnd, timezone)
        });
        logger.debug(`Added clipped busy period: ${convertUTCToTimezone(clippedStart, timezone)} to ${convertUTCToTimezone(clippedEnd, timezone)}`);
      }
      
      return overlaps;
    });
    
    logger.debug(`Found ${relevantBusyPeriods.length} relevant busy periods within working window`);
    
    // Generate free time slots
    if (relevantBusyPeriods.length === 0) {
      // No busy periods in window - entire window is free
      const freeSlot = {
        start: windowStartStr,
        end: windowEndStr
      };
      free.push(freeSlot);
      logger.debug(`No busy periods in window, entire window is free: ${freeSlot.start} to ${freeSlot.end}`);
    } else {
      // Sort relevant busy periods by start time
      const sortedRelevant = relevantBusyPeriods.sort((a, b) => {
        const aStart = new Date(a.start);
        const bStart = new Date(b.start);
        return aStart.getTime() - bStart.getTime();
      });
      
      // Add free time before first busy period
      const firstItem = sortedRelevant[0];
      if (firstItem) {
        const firstBusyStart = new Date(firstItem.start);
        if (firstBusyStart > windowStartDate) {
          const freeStart = roundToInterval(windowStartDate, granularityMinutes);
          const freeEnd = roundToInterval(firstBusyStart, granularityMinutes);
          
          if (freeStart < freeEnd) {
            const freeSlot = {
              start: convertUTCToTimezone(freeStart, timezone),
              end: convertUTCToTimezone(freeEnd, timezone)
            };
            free.push(freeSlot);
            logger.debug(`Added free slot before first busy period: ${freeSlot.start} to ${freeSlot.end}`);
          }
        }
      }
      
      // Add free time between busy periods
      for (let i = 0; i < sortedRelevant.length - 1; i++) {
        const currentItem = sortedRelevant[i];
        const nextItem = sortedRelevant[i + 1];
        
        if (currentItem && nextItem) {
          const currentBusyEnd = new Date(currentItem.end);
          const nextBusyStart = new Date(nextItem.start);
          
          if (currentBusyEnd < nextBusyStart) {
            const freeStart = roundToInterval(currentBusyEnd, granularityMinutes);
            const freeEnd = roundToInterval(nextBusyStart, granularityMinutes);
            
            if (freeStart < freeEnd) {
              const freeSlot = {
                start: convertUTCToTimezone(freeStart, timezone),
                end: convertUTCToTimezone(freeEnd, timezone)
              };
              free.push(freeSlot);
              logger.debug(`Added free slot between busy periods: ${freeSlot.start} to ${freeSlot.end}`);
            }
          }
        }
      }
      
      // Add free time after last busy period
      const lastItem = sortedRelevant[sortedRelevant.length - 1];
      if (lastItem) {
        const lastBusyEnd = new Date(lastItem.end);
        if (lastBusyEnd < windowEndDate) {
          const freeStart = roundToInterval(lastBusyEnd, granularityMinutes);
          const freeEnd = roundToInterval(windowEndDate, granularityMinutes);
          
          if (freeStart < freeEnd) {
            const freeSlot = {
              start: convertUTCToTimezone(freeStart, timezone),
              end: convertUTCToTimezone(freeEnd, timezone)
            };
            free.push(freeSlot);
            logger.debug(`Added free slot after last busy period: ${freeSlot.start} to ${freeSlot.end}`);
          }
        }
      }
    }
    
    logger.debug(`Generated ${busy.length} busy periods and ${free.length} free periods`);
    
    return { busy, free };
  }

  /**
   * Debug method to validate availability results and identify timezone issues
   */
  async debugAvailability(input: GetAvailabilityInput): Promise<any> {
    const timezone = input.timeZone || config.server.defaultTimezone;
    const windowStart = this.convertInputTimeToUTC(input.windowStart, timezone);
    const windowEnd = this.convertInputTimeToUTC(input.windowEnd, timezone);
    
    logger.info(`Debugging availability for ${input.participants.length} participants`);
    logger.info(`Input window: ${input.windowStart} to ${input.windowEnd}`);
    logger.info(`UTC window: ${windowStart.toISOString()} to ${windowEnd.toISOString()}`);
    logger.info(`Target timezone: ${timezone}`);
    
    const debugResults = [];
    
    for (const email of input.participants) {
      try {
        const client = await this.graphClient.getClient();
        
        // Get raw API response
        const rawSchedule = await this.graphClient.executeWithRetry(() =>
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
        
        const rawEvents = await this.graphClient.executeWithRetry(() =>
          client.api(`/users/${email}/calendarView`)
            .query({
              startDateTime: windowStart.toISOString(),
              endDateTime: windowEnd.toISOString(),
              $select: 'start,end,subject,showAs'
            })
            .get()
        );
        
        // Get processed availability
        const availability = await this.getUserAvailability(
          email,
          windowStart,
          windowEnd,
          input.granularityMinutes,
          input.workHoursOnly,
          timezone
        );
        
        debugResults.push({
          email,
          rawSchedule: rawSchedule.value?.[0]?.scheduleItems || [],
          rawEvents: rawEvents.value || [],
          processedAvailability: availability,
          timezone
        });
        
      } catch (error) {
        logger.error(`Failed to debug availability for ${email}:`, error);
        debugResults.push({
          email,
          error: error instanceof Error ? error.message : String(error)
        });
      }
    }
    
    return {
      input,
      timezone,
      debugResults
    };
  }

  /**
   * Convert input time from local timezone to UTC for API calls
   */
  private convertInputTimeToUTC(dateTimeString: string, timezone: string): Date {
    // If the input already has timezone offset, parse it directly
    if (dateTimeString.includes('+') || dateTimeString.includes('-')) {
      return new Date(dateTimeString);
    }
    
    // If the input ends with 'Z', it's already UTC
    if (dateTimeString.endsWith('Z')) {
      return new Date(dateTimeString);
    }
    
    // If neither, assume it's in the specified timezone and convert to UTC
    // For Dubai timezone (+04:00), we need to subtract 4 hours to get UTC
    if (timezone === 'Asia/Dubai') {
      const localDate = new Date(dateTimeString);
      // Subtract 4 hours to convert from Dubai time to UTC
      const utcDate = new Date(localDate.getTime() - (4 * 60 * 60 * 1000));
      return utcDate;
    }
    
    // For other timezones, try to calculate the offset
    try {
      const localDate = new Date(dateTimeString);
      const utcTime = localDate.getTime();
      
      // Get the timezone offset in minutes
      const timezoneOffset = new Date().toLocaleString('en-US', { timeZone: timezone });
      const localTime = new Date(timezoneOffset).getTime();
      const offsetMs = localTime - utcTime;
      
      // Apply the offset to convert to UTC
      const utcDate = new Date(localDate.getTime() - offsetMs);
      return utcDate;
    } catch (error) {
      // Fallback: assume input is already UTC
      return new Date(dateTimeString);
    }
  }

  /**
   * Format time for output, handling both UTC and local timezone inputs
   */
  private formatTimeForOutput(dateTimeString: string, timezone: string): string {
    logger.debug(`Formatting time: "${dateTimeString}" for timezone: "${timezone}"`);
    
    // Check if the input already has timezone offset
    // Look for timezone offset pattern: +HH:MM or -HH:MM at the end of the string
    // This excludes date separators like "2025-08-18"
    const timezoneOffsetPattern = /[+-]\d{2}:\d{2}$/;
    if (timezoneOffsetPattern.test(dateTimeString)) {
      logger.debug(`Time already has timezone offset: ${dateTimeString}`);
      return dateTimeString;
    }
    
    // Check if the input ends with 'Z' (UTC)
    if (dateTimeString.endsWith('Z')) {
      logger.debug(`Time is UTC, converting to timezone: ${timezone}`);
      // Convert from UTC to target timezone
      const date = new Date(dateTimeString);
      const result = convertUTCToTimezone(date, timezone);
      logger.debug(`UTC conversion result: ${result}`);
      return result;
    }
    
    // CRITICAL FIX: Microsoft Graph API often returns times that claim to be UTC
    // but are actually in local timezone. We need to detect this pattern and
    // use parseGraphAPITime to handle it correctly.
    
    // Check if this looks like a Microsoft Graph API time (has microseconds)
    if (dateTimeString.includes('.0000000')) {
      logger.debug(`Detected Microsoft Graph API time format, using parseGraphAPITime`);
      const result = parseGraphAPITime(dateTimeString, timezone);
      logger.debug(`parseGraphAPITime result: ${result}`);
      return result;
    }
    
    // For other cases, try to determine if it's actually UTC or local time
    // Microsoft Graph API sometimes returns local times without timezone indicators
    try {
      // Extract the hour from the time string
      const timeMatch = dateTimeString.match(/T(\d{2}):(\d{2}):(\d{2})/);
      if (timeMatch) {
        const hour = timeMatch[1];
        const minute = timeMatch[2];
        const second = timeMatch[3];
        
        logger.debug(`Extracted time components: hour=${hour}, minute=${minute}, second=${second}`);
        
        if (hour && minute && second) {
          const localHour = parseInt(hour);
          
          // For Dubai timezone (+04:00), if the hour is 0-12, it's likely already in local time (not UTC)
          // This is because if it were UTC, these times would be 4 hours earlier in Dubai
          if (timezone === 'Asia/Dubai' && localHour >= 0 && localHour <= 12) {
            logger.debug(`Time appears to be already in Dubai timezone (hour: ${localHour}), using parseGraphAPITime`);
            const result = parseGraphAPITime(dateTimeString, timezone);
            logger.debug(`parseGraphAPITime result: ${result}`);
            return result;
          } else {
            logger.debug(`Time appears to be UTC (hour: ${localHour}), not using parseGraphAPITime`);
          }
        }
      }
      
      // If we reach here, treat as UTC time
      logger.debug(`Treating as UTC time, converting to timezone: ${timezone}`);
      const utcDate = new Date(dateTimeString);
      const utcResult = convertUTCToTimezone(utcDate, timezone);
      logger.debug(`UTC conversion result: ${utcResult}`);
      return utcResult;
    } catch (error) {
      logger.debug(`Error in timezone conversion, falling back to parseGraphAPITime`);
      const result = parseGraphAPITime(dateTimeString, timezone);
      logger.debug(`parseGraphAPITime fallback result: ${result}`);
      return result;
    }
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

  /**
   * Validate availability input with helpful error messages
   */
  private validateAvailabilityInput(input: GetAvailabilityInput): void {
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

    // Check for proper date format
    if (!input.windowStart.endsWith('.000Z') && !input.windowStart.includes('+') && !input.windowStart.includes('-')) {
      errors.push("windowStart must be in UTC format (.000Z) or include timezone offset (+/-HH:MM). Example: 2025-08-16T08:00:00+04:00 for Dubai time");
    }

    if (!input.windowEnd.endsWith('.000Z') && !input.windowEnd.includes('+') && !input.windowEnd.includes('-')) {
      errors.push("windowEnd must be in UTC format (.000Z) or include timezone offset (+/-HH:MM). Example: 2025-08-16T18:00:00+04:00 for Dubai time");
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

    // Check granularity
    if (input.granularityMinutes < 1 || input.granularityMinutes > 1440) {
      errors.push("granularityMinutes must be between 1 and 1440");
    }

    if (errors.length > 0) {
      const errorMessage = `Availability validation failed:\n${errors.map(err => `- ${err}`).join('\n')}\n\nCorrect format:\n` +
        `{\n` +
        `  "participants": ["email1@domain.com", "email2@domain.com"],\n` +
        `  "windowStart": "2025-08-16T08:00:00+04:00",\n` +
        `  "windowEnd": "2025-08-16T18:00:00+04:00",\n` +
        `  "granularityMinutes": 30,\n` +
        `  "workHoursOnly": true,\n` +
        `  "timeZone": "Asia/Dubai"\n` +
        `}`;
      
      throw new Error(errorMessage);
    }
  }
}
