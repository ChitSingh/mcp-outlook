import { z } from 'zod';

// Base schemas
export const EmailSchema = z.string().email();
export const ISODateTimeSchema = z.string().refine(
  (val) => {
    // Check if it's a valid ISO 8601 datetime string
    const isoRegex = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d{3})?(Z|[+-]\d{2}:\d{2})$/;
    return isoRegex.test(val);
  },
  {
    message: "Date must be in ISO 8601 format with timezone. Examples: 2025-08-16T17:30:00.000Z or 2025-08-16T17:30:00+04:00"
  }
);
export const TimeZoneSchema = z.string().min(1);

// Working hours schema
export const WorkingHoursSchema = z.object({
  start: z.string().regex(/^([0-1]?[0-9]|2[0-3]):[0-5][0-9]$/), // HH:MM format
  end: z.string().regex(/^([0-1]?[0-9]|2[0-3]):[0-5][0-9]$/), // HH:MM format
  days: z.array(z.number().min(0).max(6)) // 0=Sunday, 6=Saturday
});

// Health check
export const HealthCheckInputSchema = z.object({});
export const HealthCheckOutputSchema = z.object({
  ok: z.boolean(),
  graphScopes: z.array(z.string()),
  organizer: z.string()
});

// Get availability
export const GetAvailabilityInputSchema = z.object({
  participants: z.array(EmailSchema),
  windowStart: ISODateTimeSchema,
  windowEnd: ISODateTimeSchema,
  granularityMinutes: z.number().min(1).max(1440).optional().default(30),
  workHoursOnly: z.boolean().optional().default(true),
  timeZone: TimeZoneSchema.optional()
});

export const GetAvailabilityOutputSchema = z.object({
  timeZone: z.string(),
  users: z.array(z.object({
    email: z.string(),
    workingHours: WorkingHoursSchema.nullable(),
    busy: z.array(z.object({
      start: ISODateTimeSchema,
      end: ISODateTimeSchema
    })),
    free: z.array(z.object({
      start: ISODateTimeSchema,
      end: ISODateTimeSchema
    }))
  }))
});

// Propose meeting times
export const ProposeMeetingTimesInputSchema = z.object({
  participants: z.array(EmailSchema),
  durationMinutes: z.number().min(1).max(1440),
  windowStart: ISODateTimeSchema,
  windowEnd: ISODateTimeSchema,
  maxCandidates: z.number().min(1).max(20).optional().default(5),
  bufferBeforeMinutes: z.number().min(0).max(120).optional().default(0),
  bufferAfterMinutes: z.number().min(0).max(120).optional().default(0),
  workHoursOnly: z.boolean().optional().default(true),
  minRequiredAttendees: z.number().min(1).optional(),
  organizer: z.string().optional(),
  timeZone: TimeZoneSchema.optional()
});

export const ProposeMeetingTimesOutputSchema = z.object({
  source: z.enum(['graph_findMeetingTimes', 'local_intersection']),
  candidates: z.array(z.object({
    start: ISODateTimeSchema,
    end: ISODateTimeSchema,
    attendeeAvailability: z.record(z.string(), z.enum(['free', 'tentative', 'busy'])),
    confidence: z.number().min(0).max(1)
  }))
});

// Book meeting
export const BookMeetingInputSchema = z.object({
  subject: z.string().min(1),
  participants: z.array(EmailSchema),
  required: z.array(EmailSchema).optional(),
  optional: z.array(EmailSchema).optional(),
  start: ISODateTimeSchema,
  end: ISODateTimeSchema,
  organizer: z.string().optional(),
  bodyHtml: z.string().min(1, "Meeting description (bodyHtml) is required"),
  location: z.string().optional(),
  onlineMeeting: z.boolean().optional().default(true),
  allowConflicts: z.boolean().optional().default(false),
  remindersMinutesBeforeStart: z.number().min(0).max(1440).optional().default(10)
});

export const BookMeetingOutputSchema = z.object({
  eventId: z.string(),
  iCalUid: z.string(),
  webLink: z.string(),
  organizer: z.string()
});

// Cancel meeting
export const CancelMeetingInputSchema = z.object({
  eventId: z.string(),
  organizer: z.string().optional(),
  comment: z.string().optional()
});

export const CancelMeetingOutputSchema = z.object({
  cancelled: z.literal(true),
  eventId: z.string()
});

// Error schema
export const ErrorSchema = z.object({
  error: z.object({
    code: z.string(),
    message: z.string(),
    details: z.unknown().optional()
  })
});

// Type exports
export type HealthCheckInput = z.infer<typeof HealthCheckInputSchema>;
export type HealthCheckOutput = z.infer<typeof HealthCheckOutputSchema>;
export type GetAvailabilityInput = z.infer<typeof GetAvailabilityInputSchema>;
export type GetAvailabilityOutput = z.infer<typeof GetAvailabilityOutputSchema>;
export type ProposeMeetingTimesInput = z.infer<typeof ProposeMeetingTimesInputSchema>;
export type ProposeMeetingTimesOutput = z.infer<typeof ProposeMeetingTimesOutputSchema>;
export type BookMeetingInput = z.infer<typeof BookMeetingInputSchema>;
export type BookMeetingOutput = z.infer<typeof BookMeetingOutputSchema>;
export type CancelMeetingInput = z.infer<typeof CancelMeetingInputSchema>;
export type CancelMeetingOutput = z.infer<typeof CancelMeetingOutputSchema>;
export type WorkingHours = z.infer<typeof WorkingHoursSchema>;
export type ErrorOutput = z.infer<typeof ErrorSchema>;
