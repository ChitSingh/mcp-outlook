import { TimeInterval, addBuffers, getDurationMinutes, intersectIntervals, normalizeISOString } from '../utils/time.js';
import { ProposeMeetingTimesInput, ProposeMeetingTimesOutput } from '../types.js';
import { logger } from '../config.js';
import config from '../config.js';

export interface CandidateSlot {
  start: string;
  end: string;
  attendeeAvailability: Record<string, 'free' | 'tentative' | 'busy'>;
  confidence: number;
}

export class IntersectionService {
  /**
   * Find intersecting free time slots across multiple calendars
   */
  findIntersectingSlots(
    input: ProposeMeetingTimesInput,
    userAvailabilities: Array<{
      email: string;
      free: Array<{ start: string; end: string }>;
      busy: Array<{ start: string; end: string; subject?: string | undefined }>;
    }>
  ): CandidateSlot[] {
    const { durationMinutes, maxCandidates, bufferBeforeMinutes, bufferAfterMinutes } = input;
    
    logger.info(`Finding intersecting slots for ${durationMinutes} minute meeting with ${userAvailabilities.length} participants`);

    // Handle single user case - no intersection needed
    if (userAvailabilities.length === 1) {
      const singleUser = userAvailabilities[0];
      if (!singleUser) return [];
      
      logger.info(`Single user case: ${singleUser.email} with ${singleUser.free.length} free slots`);
      
      const candidates: CandidateSlot[] = [];
      
      for (const slot of singleUser.free) {
        const slotStart = new Date(slot.start);
        const slotEnd = new Date(slot.end);
        const slotDuration = getDurationMinutes(slotStart, slotEnd);
        
        // Check if slot can accommodate the meeting duration
        if (slotDuration >= durationMinutes) {
          // Apply buffers if specified
          let meetingStart: Date;
          let meetingEnd: Date;
          
          if (bufferBeforeMinutes === 0 && bufferAfterMinutes === 0) {
            // No buffers - use the full slot
            meetingStart = new Date(slotStart);
            meetingEnd = new Date(slotEnd);
          } else {
            // Apply buffers to create the effective meeting slot
            const totalBufferTime = bufferBeforeMinutes + bufferAfterMinutes;
            const effectiveDuration = slotDuration - totalBufferTime;
            
            if (effectiveDuration < durationMinutes) {
              continue; // Not enough time after buffers
            }
            
            // Calculate the centered meeting slot
            meetingStart = new Date(slotStart);
            meetingStart.setMinutes(meetingStart.getMinutes() + bufferBeforeMinutes);
            
            meetingEnd = new Date(meetingStart);
            meetingEnd.setMinutes(meetingEnd.getMinutes() + durationMinutes);
          }
          
          candidates.push({
            start: normalizeISOString(meetingStart),
            end: normalizeISOString(meetingEnd),
            attendeeAvailability: { [singleUser.email]: 'free' },
            confidence: 1.0 // Full confidence for single user
          });
        }
      }
      
      // Sort by start time and return top candidates
      candidates.sort((a, b) => new Date(a.start).getTime() - new Date(b.start).getTime());
      return candidates.slice(0, maxCandidates);
    }

    // Convert free slots to TimeInterval objects for multi-user intersection
    const freeIntervals = userAvailabilities.map(user => ({
      email: user.email,
      intervals: user.free.map(slot => ({
        start: new Date(slot.start),
        end: new Date(slot.end)
      }))
    }));

    // Find all possible intersections
    const candidates: CandidateSlot[] = [];
    
    // Start with the first user's free slots
    if (freeIntervals.length === 0 || freeIntervals[0]?.intervals.length === 0) {
      return [];
    }

    const firstUser = freeIntervals[0];
    if (!firstUser) {
      return [];
    }
    
    for (const interval of firstUser.intervals) {
      // Find intersection with other users first
      let commonInterval = interval;
      const attendeeAvailability: Record<string, 'free' | 'tentative' | 'busy'> = {
        [firstUser.email]: 'free'
      };

      for (let i = 1; i < freeIntervals.length; i++) {
        const user = freeIntervals[i];
        if (!user) continue;
        
        const userIntersection = this.findBestIntersection(commonInterval, user.intervals);
        
        if (!userIntersection) {
          // No intersection with this user
          attendeeAvailability[user.email] = 'busy';
          continue;
        }

        commonInterval = userIntersection;
        attendeeAvailability[user.email] = 'free';
      }

      // For multi-user case, we need at least one intersection
      if (freeIntervals.length > 1 && commonInterval === interval) {
        // No intersection found with other users, skip this slot
        continue;
      }

      // Check if the intersection can accommodate the meeting duration
      const intersectionDuration = getDurationMinutes(commonInterval.start, commonInterval.end);
      if (intersectionDuration >= durationMinutes) {
        let slotStart: Date;
        let slotEnd: Date;
        
        if (bufferBeforeMinutes === 0 && bufferAfterMinutes === 0) {
          // No buffers - use the full intersection
          slotStart = new Date(commonInterval.start);
          slotEnd = new Date(commonInterval.end);
        } else {
          // Apply buffers to create the effective meeting slot
          // The meeting slot should be centered within the intersection
          const totalBufferTime = bufferBeforeMinutes + bufferAfterMinutes;
          const effectiveDuration = intersectionDuration - totalBufferTime;
          
          if (effectiveDuration < durationMinutes) {
            continue; // Not enough time after buffers
          }
          
          // Calculate the centered meeting slot
          slotStart = new Date(commonInterval.start);
          slotStart.setMinutes(slotStart.getMinutes() + bufferBeforeMinutes);
          
          slotEnd = new Date(slotStart);
          slotEnd.setMinutes(slotEnd.getMinutes() + durationMinutes);
        }
        
        const confidence = this.calculateConfidence(attendeeAvailability, userAvailabilities.length);
        
        candidates.push({
          start: normalizeISOString(slotStart),
          end: normalizeISOString(slotEnd),
          attendeeAvailability,
          confidence
        });
      }
    }

    // Sort candidates by confidence and start time
    candidates.sort((a, b) => {
      if (Math.abs(a.confidence - b.confidence) > 0.1) {
        return b.confidence - a.confidence; // Higher confidence first
      }
      return new Date(a.start).getTime() - new Date(b.start).getTime(); // Earlier start first
    });

    // Return top candidates
    return candidates.slice(0, maxCandidates);
  }

  /**
   * Find the best intersection between a target interval and user's free intervals
   */
  private findBestIntersection(
    target: TimeInterval,
    userIntervals: TimeInterval[]
  ): TimeInterval | null {
    let bestIntersection: TimeInterval | null = null;
    let bestOverlap = 0;

    for (const userInterval of userIntervals) {
      const intersection = intersectIntervals(target, userInterval);
      if (intersection) {
        const overlap = getDurationMinutes(intersection.start, intersection.end);
        if (overlap > bestOverlap) {
          bestOverlap = overlap;
          bestIntersection = intersection;
        }
      }
    }

    return bestIntersection;
  }

  /**
   * Calculate confidence score for a candidate slot
   */
  private calculateConfidence(
    attendeeAvailability: Record<string, 'free' | 'tentative' | 'busy'>,
    totalAttendees: number
  ): number {
    let freeCount = 0;
    let tentativeCount = 0;
    let busyCount = 0;

    for (const status of Object.values(attendeeAvailability)) {
      switch (status) {
        case 'free':
          freeCount++;
          break;
        case 'tentative':
          tentativeCount++;
          break;
        case 'busy':
          busyCount++;
          break;
      }
    }

    // Calculate confidence based on availability
    const freeWeight = 1.0;
    const tentativeWeight = 0.5;
    const busyWeight = 0.0;

    const totalScore = (freeCount * freeWeight) + (tentativeCount * tentativeWeight) + (busyCount * busyWeight);
    const maxScore = totalAttendees;

    return maxScore > 0 ? totalScore / maxScore : 0;
  }

  /**
   * Check if a slot meets minimum attendance requirements
   */
  checkMinimumAttendance(
    slot: CandidateSlot,
    minRequiredAttendees: number | undefined
  ): boolean {
    if (!minRequiredAttendees) {
      return true; // No minimum requirement
    }

    const availableAttendees = Object.values(slot.attendeeAvailability).filter(
      status => status === 'free' || status === 'tentative'
    ).length;

    return availableAttendees >= minRequiredAttendees;
  }

  /**
   * Filter candidates by minimum attendance
   */
  filterByMinimumAttendance(
    candidates: CandidateSlot[],
    minRequiredAttendees: number | undefined
  ): CandidateSlot[] {
    if (!minRequiredAttendees) {
      return candidates;
    }

    return candidates.filter(candidate => 
      this.checkMinimumAttendance(candidate, minRequiredAttendees)
    );
  }
}
