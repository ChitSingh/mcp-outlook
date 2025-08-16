import { z } from 'zod';

// Time interval type
export interface TimeInterval {
  start: Date;
  end: Date;
}

// Working hours type
export interface WorkingHours {
  start: string; // HH:MM format
  end: string; // HH:MM format
  days: number[]; // 0=Sunday, 6=Saturday
}

/**
 * Convert a date to a specific timezone and format as ISO string
 */
export function toTimezoneISO(date: Date, timezone: string): string {
  try {
    return date.toLocaleString('en-CA', {
      timeZone: timezone,
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
      hour12: false
    }).replace(',', 'T') + 'Z';
  } catch {
    // Fallback to UTC if timezone is invalid
    return date.toISOString();
  }
}

/**
 * Parse ISO string and convert to UTC Date
 */
export function parseISOToUTC(isoString: string): Date {
  return new Date(isoString);
}

/**
 * Round a date to the nearest interval (e.g., 30 minutes)
 */
export function roundToInterval(date: Date, intervalMinutes: number): Date {
  const minutes = date.getMinutes();
  const roundedMinutes = Math.round(minutes / intervalMinutes) * intervalMinutes;
  
  const rounded = new Date(date);
  rounded.setMinutes(roundedMinutes, 0, 0);
  
  return rounded;
}

/**
 * Add buffer time before and after an interval
 */
export function addBuffers(
  interval: TimeInterval,
  bufferBeforeMinutes: number,
  bufferAfterMinutes: number
): TimeInterval {
  const start = new Date(interval.start);
  const end = new Date(interval.end);
  
  start.setMinutes(start.getMinutes() - bufferBeforeMinutes);
  end.setMinutes(end.getMinutes() + bufferAfterMinutes);
  
  return { start, end };
}

/**
 * Check if a date falls within working hours
 */
export function isWithinWorkingHours(
  date: Date,
  workingHours: WorkingHours,
  timezone: string
): boolean {
  try {
    const localDate = new Date(date.toLocaleString('en-US', { timeZone: timezone }));
    const dayOfWeek = localDate.getDay();
    
    if (!workingHours.days.includes(dayOfWeek)) {
      return false;
    }
    
    const timeString = localDate.toTimeString().slice(0, 5); // HH:MM
    return timeString >= workingHours.start && timeString <= workingHours.end;
  } catch {
    // If timezone conversion fails, assume it's within working hours
    return true;
  }
}

/**
 * Clip an interval to working hours
 */
export function clipToWorkingHours(
  interval: TimeInterval,
  workingHours: WorkingHours,
  timezone: string
): TimeInterval | null {
  if (!isWithinWorkingHours(interval.start, workingHours, timezone) ||
      !isWithinWorkingHours(interval.end, workingHours, timezone)) {
    return null;
  }
  
  return interval;
}

/**
 * Calculate the duration between two dates in minutes
 */
export function getDurationMinutes(start: Date, end: Date): number {
  return Math.round((end.getTime() - start.getTime()) / (1000 * 60));
}

/**
 * Check if two time intervals overlap
 */
export function intervalsOverlap(a: TimeInterval, b: TimeInterval): boolean {
  return a.start < b.end && b.start < a.end;
}

/**
 * Find the intersection of two time intervals
 */
export function intersectIntervals(a: TimeInterval, b: TimeInterval): TimeInterval | null {
  if (!intervalsOverlap(a, b)) {
    return null;
  }
  
  const start = new Date(Math.max(a.start.getTime(), b.start.getTime()));
  const end = new Date(Math.min(a.end.getTime(), b.end.getTime()));
  
  return { start, end };
}

/**
 * Validate timezone string
 */
export function isValidTimezone(timezone: string): boolean {
  try {
    Intl.DateTimeFormat(undefined, { timeZone: timezone });
    return true;
  } catch {
    return false;
  }
}

/**
 * Get current time in a specific timezone
 */
export function getCurrentTimeInTimezone(timezone: string): Date {
  try {
    const now = new Date();
    const localTime = now.toLocaleString('en-US', { timeZone: timezone });
    return new Date(localTime);
  } catch {
    return new Date();
  }
}

/**
 * Normalize ISO time string to remove milliseconds for consistent formatting
 */
export function normalizeISOString(date: Date): string {
  return date.toISOString().replace(/\.\d{3}Z$/, 'Z');
}

/**
 * Parse Microsoft Graph API time strings that are already in local timezone
 * but get misinterpreted as UTC by JavaScript's Date constructor.
 * 
 * This function handles the specific case where the API returns:
 * - "2025-08-18T04:00:00.0000000" (which represents 8:00 AM Dubai time)
 * - But JavaScript treats it as 4:00 AM UTC
 * 
 * We extract the raw components and format them with the correct timezone offset.
 */
export function parseGraphAPITime(dateTimeString: string, timezone: string): string {
  // Remove any milliseconds and timezone indicators
  const cleanString = dateTimeString.replace(/\.\d{6,}Z?$/, '');
  
  // Extract year, month, day, hour, minute, second
  const match = cleanString.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})/);
  if (!match) {
    // Fallback to regular parsing if format doesn't match
    const date = new Date(dateTimeString);
    return convertUTCToTimezone(date, timezone);
  }
  
  const [, year, month, day, hour, minute, second] = match;
  
  // Ensure all components are defined
  if (!year || !month || !day || !hour || !minute || !second) {
    // Fallback to regular parsing if any component is missing
    const date = new Date(dateTimeString);
    return convertUTCToTimezone(date, timezone);
  }
  
  // For Dubai timezone, we know it's +04:00
  if (timezone === 'Asia/Dubai') {
    return `${year}-${month}-${day}T${hour}:${minute}:${second}+04:00`;
  }
  
  // For other timezones, calculate the offset more reliably
  try {
    // Create a date object representing the local time in the target timezone
    const localDate = new Date(`${year}-${month}-${day}T${hour}:${minute}:${second}`);
    
    // Get the timezone offset using Intl.DateTimeFormat
    const formatter = new Intl.DateTimeFormat('en-US', {
      timeZone: timezone,
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
      hour12: false
    });
    
    // Format the date in the target timezone
    const localTimeString = formatter.formatToParts(localDate);
    const localTimeMap = new Map(localTimeString.map(part => [part.type, part.value]));
    
    // Get the timezone offset in minutes for the current date
    const utcDate = new Date(Date.UTC(
      parseInt(year),
      parseInt(month) - 1,
      parseInt(day),
      parseInt(hour),
      parseInt(minute),
      parseInt(second)
    ));
    
    // Calculate the offset by comparing UTC time with local time in the target timezone
    const utcTime = utcDate.getTime();
    const localTimeInTimezone = new Date(`${year}-${month}-${day}T${hour}:${minute}:${second}`).getTime();
    
    // The offset is the difference between local time and UTC time
    // If local time is ahead of UTC, offset is positive
    const offsetMs = localTimeInTimezone - utcTime;
    const offsetHours = Math.floor(Math.abs(offsetMs) / (1000 * 60 * 60));
    const offsetMinutes = Math.floor((Math.abs(offsetMs) / (1000 * 60)) % 60);
    const offsetSign = offsetMs >= 0 ? '+' : '-';
    
    return `${year}-${month}-${day}T${hour}:${minute}:${second}${offsetSign}${String(offsetHours).padStart(2, '0')}:${String(offsetMinutes).padStart(2, '0')}`;
  } catch (error) {
    // Fallback: assume the time is already in the correct timezone
    return `${year}-${month}-${day}T${hour}:${minute}:${second}`;
  }
}

/**
 * Convert UTC time to a specific timezone and format as ISO string with timezone offset
 */
export function convertUTCToTimezone(utcDate: Date, timezone: string): string {
  try {
    // For Dubai timezone, we know it's +04:00
    if (timezone === 'Asia/Dubai') {
      const localDate = new Date(utcDate.getTime() + (4 * 60 * 60 * 1000)); // Add 4 hours
      
      const year = localDate.getUTCFullYear();
      const month = String(localDate.getUTCMonth() + 1).padStart(2, '0');
      const day = String(localDate.getUTCDate()).padStart(2, '0');
      const hour = String(localDate.getUTCHours()).padStart(2, '0');
      const minute = String(localDate.getUTCMinutes()).padStart(2, '0');
      const second = String(localDate.getUTCSeconds()).padStart(2, '0');
      
      return `${year}-${month}-${day}T${hour}:${minute}:${second}+04:00`;
    }
    
    // For other timezones, try to calculate the offset
    const localDate = new Date(utcDate.toLocaleString('en-US', { timeZone: timezone }));
    
    // Get the timezone offset in minutes
    const utcTime = utcDate.getTime();
    const localTime = localDate.getTime();
    const offsetMs = localTime - utcTime;
    
    // Calculate offset hours and minutes
    const offsetHours = Math.floor(Math.abs(offsetMs) / (1000 * 60 * 60));
    const offsetMinutes = Math.floor((Math.abs(offsetMs) / (1000 * 60)) % 60);
    const offsetSign = offsetMs >= 0 ? '+' : '-';
    
    // Format the local date
    const year = localDate.getFullYear();
    const month = String(localDate.getMonth() + 1).padStart(2, '0');
    const day = String(localDate.getDate()).padStart(2, '0');
    const hour = String(localDate.getHours()).padStart(2, '0');
    const minute = String(localDate.getMinutes()).padStart(2, '0');
    const second = String(localDate.getSeconds()).padStart(2, '0');
    
    return `${year}-${month}-${day}T${hour}:${minute}:${second}${offsetSign}${String(offsetHours).padStart(2, '0')}:${String(offsetMinutes).padStart(2, '0')}`;
  } catch (error) {
    // Fallback to UTC if timezone conversion fails
    return utcDate.toISOString();
  }
}
