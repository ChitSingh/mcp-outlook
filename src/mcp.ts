import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { 
  CallToolRequestSchema, 
  ListToolsRequestSchema,
  Tool 
} from '@modelcontextprotocol/sdk/types.js';
import { AvailabilityService } from './scheduling/availability.js';
import { FindMeetingTimesService } from './scheduling/find.js';
import { BookingService } from './scheduling/book.js';
import { GraphClientFactory } from './graph/client.js';
import { GraphAuth } from './graph/auth.js';
import {
  HealthCheckInputSchema,
  HealthCheckOutputSchema,
  GetAvailabilityInputSchema,
  GetAvailabilityOutputSchema,
  ProposeMeetingTimesInputSchema,
  ProposeMeetingTimesOutputSchema,
  BookMeetingInputSchema,
  BookMeetingOutputSchema,
  CancelMeetingInputSchema,
  CancelMeetingOutputSchema,
  ErrorSchema
} from './types.js';
import { logger } from './config.js';
import config from './config.js';

export class MCPServer {
  private server: Server;
  private availabilityService: AvailabilityService;
  private findMeetingTimesService: FindMeetingTimesService;
  private bookingService: BookingService;
  private graphClient: GraphClientFactory;

  constructor() {
    this.server = new Server(
      {
        name: 'mcp-outlook-scheduler',
        version: '1.0.0'
      }
    );

    // Create a single shared GraphAuth instance
    const sharedAuth = new GraphAuth();
    
    // Create a single shared GraphClientFactory instance using the shared auth
    this.graphClient = new GraphClientFactory(sharedAuth);
    
    // Pass the shared GraphClientFactory to all services
    this.availabilityService = new AvailabilityService(this.graphClient);
    this.findMeetingTimesService = new FindMeetingTimesService(this.graphClient);
    this.bookingService = new BookingService(this.graphClient);

    this.setupToolHandlers();
  }

  /**
   * Setup all MCP tool handlers
   */
  private setupToolHandlers(): void {
    // Health check tool
    this.server.setRequestHandler(ListToolsRequestSchema, async () => {
      return {
        tools: [
          {
            name: 'health_check',
            description: `Check the health and status of the MCP server.

REQUIRED FORMAT:
{
  // No parameters required - just call the tool
}

RESPONSE INCLUDES:
- Server status (ok: true/false)
- Microsoft Graph API scopes available
- Organizer email configured
- Authentication status

Use this tool to verify the server is working before using other tools.`,
            inputSchema: {
              type: 'object',
              properties: {},
              required: []
            }
          },
          {
            name: 'get_availability',
            description: `Get free/busy availability for multiple participants.

REQUIRED FORMAT:
{
  "participants": ["email1@domain.com", "email2@domain.com"],
  "windowStart": "2025-08-16T08:00:00+04:00",     // Start time in Dubai timezone (+04:00)
  "windowEnd": "2025-08-16T18:00:00+04:00",       // End time in Dubai timezone (+04:00)
  "granularityMinutes": 30,                        // Time slots in minutes (1-1440)
  "workHoursOnly": true,                          // Only show working hours
  "timeZone": "Asia/Dubai"                        // Timezone for working hours
}

IMPORTANT: 
- Use Dubai timezone (+04:00) for local times
- Use ISO 8601 format for dates
- Participants must be valid email addresses
- Example: 9:00 AM Dubai = "2025-08-16T09:00:00+04:00"`,
            inputSchema: {
              type: 'object',
              properties: {
                participants: {
                  type: 'array',
                  items: { type: 'string', format: 'email' },
                  description: 'List of participant email addresses (required)'
                },
                windowStart: {
                  type: 'string',
                  format: 'date-time',
                  description: 'Start of availability window in Dubai timezone format (YYYY-MM-DDTHH:MM:SS+04:00) - required'
                },
                windowEnd: {
                  type: 'string',
                  format: 'date-time',
                  description: 'End of availability window in Dubai timezone format (YYYY-MM-DDTHH:MM:SS+04:00) - required'
                },
                granularityMinutes: {
                  type: 'number',
                  minimum: 1,
                  maximum: 1440,
                  default: 30,
                  description: 'Time granularity in minutes (1-1440)'
                },
                workHoursOnly: {
                  type: 'boolean',
                  default: true,
                  description: 'Only show availability during working hours'
                },
                timeZone: {
                  type: 'string',
                  description: 'Timezone for working hours (IANA format, e.g., Asia/Dubai)'
                }
              },
              required: ['participants', 'windowStart', 'windowEnd']
            }
          },
          {
            name: 'debug_availability',
            description: `Debug availability data and identify timezone issues for multiple participants.

REQUIRED FORMAT:
{
  "participants": ["email1@domain.com", "email2@domain.com"],
  "windowStart": "2025-08-16T08:00:00+04:00",     // Start time in Dubai timezone (+04:00)
  "windowEnd": "2025-08-16T18:00:00+04:00",       // End time in Dubai timezone (+04:00)
  "granularityMinutes": 30,                        // Time slots in minutes (1-1440)
  "workHoursOnly": true,                          // Only show working hours
  "timeZone": "Asia/Dubai"                        // Timezone for working hours
}

RESPONSE INCLUDES:
- Raw Microsoft Graph API responses
- Processed availability data
- Timezone conversion details
- Detailed logging for debugging

Use this tool to diagnose timezone and availability issues.`,
            inputSchema: {
              type: 'object',
              properties: {
                participants: {
                  type: 'array',
                  items: { type: 'string', format: 'email' },
                  description: 'List of participant email addresses (required)'
                },
                windowStart: {
                  type: 'string',
                  format: 'date-time',
                  description: 'Start of availability window in Dubai timezone format (YYYY-MM-DDTHH:MM:SS+04:00) - required'
                },
                windowEnd: {
                  type: 'string',
                  format: 'date-time',
                  description: 'End of availability window in Dubai timezone format (YYYY-MM-DDTHH:MM:SS+04:00) - required'
                },
                granularityMinutes: {
                  type: 'number',
                  minimum: 1,
                  maximum: 1440,
                  default: 30,
                  description: 'Time slots in minutes (1-1440)'
                },
                workHoursOnly: {
                  type: 'boolean',
                  default: true,
                  description: 'Only show working hours'
                },
                timeZone: {
                  type: 'string',
                  description: 'Timezone for working hours (IANA format, e.g., Asia/Dubai)'
                }
              },
              required: ['participants', 'windowStart', 'windowEnd']
            }
          },
          {
            name: 'propose_meeting_times',
            description: `Find available meeting time slots for participants.

REQUIRED FORMAT:
{
  "participants": ["email1@domain.com", "email2@domain.com"],
  "durationMinutes": 60,                           // Meeting duration in minutes (1-1440)
  "windowStart": "2025-08-16T09:00:00+04:00",     // Start of search window in Dubai time (+04:00)
  "windowEnd": "2025-08-16T17:00:00+04:00",       // End of search window in Dubai time (+04:00)
  "maxCandidates": 5,                              // Max time slots to return (1-20)
  "bufferBeforeMinutes": 10,                       // Buffer before meeting (0-120)
  "bufferAfterMinutes": 10,                        // Buffer after meeting (0-120)
  "workHoursOnly": true,                           // Only consider working hours
  "minRequiredAttendees": 2,                       // Minimum attendees required
  "timeZone": "Asia/Dubai"                         // Timezone for working hours
}

IMPORTANT: 
- Use Dubai timezone (+04:00) for local times
- Use ISO 8601 format for dates
- Duration must be between 1-1440 minutes
- Participants must be valid email addresses
- Example: 9:00 AM Dubai = "2025-08-16T09:00:00+04:00"
- Your email address is automatically set as the organizer from configuration`,
            inputSchema: {
              type: 'object',
              properties: {
                participants: {
                  type: 'array',
                  items: { type: 'string', format: 'email' },
                  description: 'List of participant email addresses (required)'
                },
                durationMinutes: {
                  type: 'number',
                  minimum: 1,
                  maximum: 1440,
                  description: 'Meeting duration in minutes (1-1440) - required'
                },
                windowStart: {
                  type: 'string',
                  format: 'date-time',
                  description: 'Start of search window in Dubai timezone format (YYYY-MM-DDTHH:MM:SS+04:00) - required'
                },
                windowEnd: {
                  type: 'string',
                  format: 'date-time',
                  description: 'End of search window in Dubai timezone format (YYYY-MM-DDTHH:MM:SS+04:00) - required'
                },
                maxCandidates: {
                  type: 'number',
                  minimum: 1,
                  maximum: 20,
                  default: 5,
                  description: 'Maximum number of time slot candidates (1-20)'
                },
                bufferBeforeMinutes: {
                  type: 'number',
                  minimum: 0,
                  maximum: 120,
                  default: 0,
                  description: 'Buffer time before meeting in minutes (0-120)'
                },
                bufferAfterMinutes: {
                  type: 'number',
                  minimum: 0,
                  maximum: 120,
                  default: 0,
                  description: 'Buffer time after meeting in minutes (0-120)'
                },
                workHoursOnly: {
                  type: 'boolean',
                  default: true,
                  description: 'Only consider working hours'
                },
                minRequiredAttendees: {
                  type: 'number',
                  minimum: 1,
                  description: 'Minimum number of required attendees'
                },
                organizer: {
                  type: 'string',
                  format: 'email',
                  description: 'Meeting organizer email (optional - will use your configured email if not provided)'
                },
                timeZone: {
                  type: 'string',
                  description: 'Timezone for working hours (IANA format, e.g., Asia/Dubai)'
                }
              },
              required: ['participants', 'durationMinutes', 'windowStart', 'windowEnd']
            }
          },
          {
            name: 'book_meeting',
            description: `Book a meeting by creating a calendar event in Outlook.

REQUIRED FORMAT:
{
  "start": "2025-08-16T14:30:00+04:00",     // Start time in Dubai timezone (+04:00)
  "end": "2025-08-16T15:00:00+04:00",       // End time in Dubai timezone (+04:00)
  "subject": "Meeting Subject",                // Meeting title
  "participants": ["email1@domain.com"],      // Array of participant emails
  "bodyHtml": "<p>Meeting description</p>",   // REQUIRED: HTML description
  "onlineMeeting": true,                      // Whether it's online
  "remindersMinutesBeforeStart": 15           // Reminder time
}

IMPORTANT: 
- Use Dubai timezone (+04:00) for local times
- bodyHtml is required for meeting description
- Use ISO 8601 format for dates
- Example: 2:30 PM Dubai = "2025-08-16T14:30:00+04:00"
- Your email address is automatically set as the organizer from configuration`,
            inputSchema: {
              type: 'object',
              properties: {
                subject: {
                  type: 'string',
                  minLength: 1,
                  description: 'Meeting subject/title (required)'
                },
                participants: {
                  type: 'array',
                  items: { type: 'string', format: 'email' },
                  description: 'List of participant email addresses (required)'
                },
                required: {
                  type: 'array',
                  items: { type: 'string', format: 'email' },
                  description: 'Required attendees (subset of participants)'
                },
                optional: {
                  type: 'array',
                  items: { type: 'string', format: 'email' },
                  description: 'Optional attendees (subset of participants)'
                },
                start: {
                  type: 'string',
                  format: 'date-time',
                  description: 'Meeting start time in Dubai timezone format (YYYY-MM-DDTHH:MM:SS+04:00) - required'
                },
                end: {
                  type: 'string',
                  format: 'date-time',
                  description: 'Meeting end time in Dubai timezone format (YYYY-MM-DDTHH:MM:SS+04:00) - required'
                },
                organizer: {
                  type: 'string',
                  format: 'email',
                  description: 'Meeting organizer email (optional - will use your configured email if not provided)'
                },
                bodyHtml: {
                  type: 'string',
                  description: 'REQUIRED: Meeting body content in HTML format (e.g., "<p>Meeting description</p>")'
                },
                location: {
                  type: 'string',
                  description: 'Meeting location'
                },
                onlineMeeting: {
                  type: 'boolean',
                  default: true,
                  description: 'Create Teams online meeting'
                },
                allowConflicts: {
                  type: 'boolean',
                  default: false,
                  description: 'Allow scheduling conflicts'
                },
                remindersMinutesBeforeStart: {
                  type: 'number',
                  minimum: 0,
                  maximum: 1440,
                  default: 10,
                  description: 'Reminder time in minutes before start'
                }
              },
              required: ['subject', 'participants', 'start', 'end', 'bodyHtml']
            }
          },
          {
            name: 'cancel_meeting',
            description: `Cancel an existing meeting.

REQUIRED FORMAT:
{
  "eventId": "AAMkAGI2TG93AAA=",                  // Calendar event ID to cancel (required)
  "comment": "Meeting cancelled due to conflict"   // Optional cancellation comment
}

IMPORTANT: 
- eventId is required (get this from book_meeting response)
- Your email address is automatically set as the organizer from configuration
- comment is optional but recommended for audit trail
- No time formatting needed for this tool`,
            inputSchema: {
              type: 'object',
              properties: {
                eventId: {
                  type: 'string',
                  description: 'Calendar event ID to cancel (required) - get this from book_meeting response'
                },
                organizer: {
                  type: 'string',
                  format: 'email',
                  description: 'Meeting organizer email (optional - will use your configured email if not provided)'
                },
                comment: {
                  type: 'string',
                  description: 'Optional cancellation comment for audit trail'
                }
              },
              required: ['eventId']
            }
          }
        ]
      };
    });

    // Tool execution handler
        this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      const { name, arguments: args } = request.params;
      
      logger.info(`Tool called: ${name}`, { arguments: args });
      
      try {
        switch (name) {
          case 'health_check':
            return await this.handleHealthCheck();

          case 'get_availability':
            return await this.handleGetAvailability(args);

          case 'debug_availability':
            return await this.handleDebugAvailability(args);

          case 'propose_meeting_times':
            return await this.handleProposeMeetingTimes(args);

          case 'book_meeting':
            return await this.handleBookMeeting(args);

          case 'cancel_meeting':
            return await this.handleCancelMeeting(args);

          default:
            throw new Error(`Unknown tool: ${name}`);
        }
      } catch (error) {
        logger.error(`Tool execution failed for ${name}:`, error);
        
        const errorOutput: any = {
          error: {
            code: 'TOOL_EXECUTION_FAILED',
            message: error instanceof Error ? error.message : 'Unknown error occurred',
            details: error instanceof Error ? error.stack : undefined
          }
        };

        return errorOutput;
      }
    });
  }

  /**
   * Handle health check tool
   */
  private async handleHealthCheck(): Promise<any> {
    try {
      const scopes = this.graphClient.getScopes();
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              ok: true,
              graphScopes: scopes,
              organizer: config.organizer.email
            })
          }
        ]
      };
    } catch (error) {
      logger.error('Health check failed:', error);
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              ok: false,
              graphScopes: [],
              organizer: config.organizer.email,
              error: error instanceof Error ? error.message : 'Unknown error'
            })
          }
        ]
      };
    }
  }

  /**
   * Handle get availability tool
   */
  private async handleGetAvailability(args: any): Promise<any> {
    logger.info('Received get_availability tool request', { args });
    
    const input = GetAvailabilityInputSchema.parse(args);
    logger.info('Parsed availability input', { 
      participants: input.participants.length,
      windowStart: input.windowStart,
      windowEnd: input.windowEnd
    });
    
    const result = await this.availabilityService.getAvailability(input);
    
    logger.info('Successfully retrieved availability', { 
      users: result.users.length,
      timeZone: result.timeZone
    });
    
    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(result)
        }
      ]
    };
  }

  /**
   * Handle debug availability tool
   */
  private async handleDebugAvailability(args: any): Promise<any> {
    logger.info('Received debug_availability tool request', { args });
    
    const input = GetAvailabilityInputSchema.parse(args);
    logger.info('Parsed debug availability input', { 
      participants: input.participants.length,
      windowStart: input.windowStart,
      windowEnd: input.windowEnd
    });
    
    const result = await this.availabilityService.debugAvailability(input);
    
    logger.info('Successfully debugged availability', { 
      participants: input.participants.length,
      timeZone: result.timeZone
    });
    
    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(result, null, 2)
        }
      ]
    };
  }

  /**
   * Handle propose meeting times tool
   */
  private async handleProposeMeetingTimes(args: any): Promise<any> {
    logger.info('Received propose_meeting_times tool request', { args });
    
    const input = ProposeMeetingTimesInputSchema.parse(args);
    logger.info('Parsed propose meeting times input', { 
      participants: input.participants.length,
      durationMinutes: input.durationMinutes,
      windowStart: input.windowStart,
      windowEnd: input.windowEnd
    });
    
    const result = await this.findMeetingTimesService.findMeetingTimes(input);
    
    logger.info('Successfully proposed meeting times', { 
      candidates: result.candidates.length,
      source: result.source
    });
    
    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(result)
        }
      ]
    };
  }

  /**
   * Handle book meeting tool
   */
  private async handleBookMeeting(args: any): Promise<any> {
    logger.info('Received book_meeting tool request', { args });
    
    const input = BookMeetingInputSchema.parse(args);
    logger.info('Parsed book meeting input', { 
      subject: input.subject,
      start: input.start,
      end: input.end,
      participants: input.participants.length
    });
    
    const result = await this.bookingService.bookMeeting(input);
    
    logger.info('Successfully booked meeting', { 
      eventId: result.eventId,
      organizer: result.organizer
    });
    
    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(result)
        }
      ]
    };
  }

  /**
   * Handle cancel meeting tool
   */
  private async handleCancelMeeting(args: any): Promise<any> {
    const input = CancelMeetingInputSchema.parse(args);
    const result = await this.bookingService.cancelMeeting(input);
    
    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(result)
        }
      ]
    };
  }

  /**
   * Start the MCP server
   */
  async start(): Promise<void> {
    logger.info('Starting MCP Outlook Scheduler server...');
    
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
    
    logger.info('MCP server started successfully');
  }
}
