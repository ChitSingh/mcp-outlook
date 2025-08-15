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
            description: 'Check the health and status of the MCP server',
            inputSchema: {
              type: 'object',
              properties: {},
              required: []
            }
          },
          {
            name: 'get_availability',
            description: 'Get free/busy availability for multiple participants',
            inputSchema: {
              type: 'object',
              properties: {
                participants: {
                  type: 'array',
                  items: { type: 'string', format: 'email' },
                  description: 'List of participant email addresses'
                },
                windowStart: {
                  type: 'string',
                  format: 'date-time',
                  description: 'Start of availability window (ISO 8601)'
                },
                windowEnd: {
                  type: 'string',
                  format: 'date-time',
                  description: 'End of availability window (ISO 8601)'
                },
                granularityMinutes: {
                  type: 'number',
                  minimum: 1,
                  maximum: 1440,
                  default: 30,
                  description: 'Time granularity in minutes'
                },
                workHoursOnly: {
                  type: 'boolean',
                  default: true,
                  description: 'Only show availability during working hours'
                },
                timeZone: {
                  type: 'string',
                  description: 'Timezone for working hours (IANA format)'
                }
              },
              required: ['participants', 'windowStart', 'windowEnd']
            }
          },
          {
            name: 'propose_meeting_times',
            description: 'Find available meeting time slots for participants',
            inputSchema: {
              type: 'object',
              properties: {
                participants: {
                  type: 'array',
                  items: { type: 'string', format: 'email' },
                  description: 'List of participant email addresses'
                },
                durationMinutes: {
                  type: 'number',
                  minimum: 1,
                  maximum: 1440,
                  description: 'Meeting duration in minutes'
                },
                windowStart: {
                  type: 'string',
                  format: 'date-time',
                  description: 'Start of search window (ISO 8601)'
                },
                windowEnd: {
                  type: 'string',
                  format: 'date-time',
                  description: 'End of search window (ISO 8601)'
                },
                maxCandidates: {
                  type: 'number',
                  minimum: 1,
                  maximum: 20,
                  default: 5,
                  description: 'Maximum number of time slot candidates'
                },
                bufferBeforeMinutes: {
                  type: 'number',
                  minimum: 0,
                  maximum: 120,
                  default: 0,
                  description: 'Buffer time before meeting in minutes'
                },
                bufferAfterMinutes: {
                  type: 'number',
                  minimum: 0,
                  maximum: 120,
                  default: 0,
                  description: 'Buffer time after meeting in minutes'
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
                  description: 'Meeting organizer email'
                },
                timeZone: {
                  type: 'string',
                  description: 'Timezone for working hours (IANA format)'
                }
              },
              required: ['participants', 'durationMinutes', 'windowStart', 'windowEnd']
            }
          },
          {
            name: 'book_meeting',
            description: 'Book a meeting by creating a calendar event',
            inputSchema: {
              type: 'object',
              properties: {
                subject: {
                  type: 'string',
                  minLength: 1,
                  description: 'Meeting subject/title'
                },
                participants: {
                  type: 'array',
                  items: { type: 'string', format: 'email' },
                  description: 'List of participant email addresses'
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
                  description: 'Meeting start time (ISO 8601)'
                },
                end: {
                  type: 'string',
                  format: 'date-time',
                  description: 'Meeting end time (ISO 8601)'
                },
                organizer: {
                  type: 'string',
                  format: 'email',
                  description: 'Meeting organizer email'
                },
                bodyHtml: {
                  type: 'string',
                  description: 'Meeting body content in HTML'
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
              required: ['subject', 'participants', 'start', 'end']
            }
          },
          {
            name: 'cancel_meeting',
            description: 'Cancel an existing meeting',
            inputSchema: {
              type: 'object',
              properties: {
                eventId: {
                  type: 'string',
                  description: 'Calendar event ID to cancel'
                },
                organizer: {
                  type: 'string',
                  format: 'email',
                  description: 'Meeting organizer email'
                },
                comment: {
                  type: 'string',
                  description: 'Cancellation comment'
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
