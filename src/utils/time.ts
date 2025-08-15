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
