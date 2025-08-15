# MCP Outlook Scheduler

An MCP (Model Context Protocol) server that connects to Microsoft Outlook via Microsoft Graph to schedule meetings among multiple participants by checking their calendars and proposing or booking the earliest viable time slots.

## Features

- **Smart Meeting Scheduling**: Uses Microsoft Graph's `findMeetingTimes` API when available, with fallback to local intersection logic
- **Multi-Participant Support**: Check availability across N participants and find common free time slots
- **Flexible Constraints**: Support for working hours, buffer times, minimum attendance requirements, and timezone handling
- **Automatic Booking**: Create Outlook calendar events with Teams integration
- **MCP Integration**: Exposes tools for AI assistants to schedule meetings programmatically

## Prerequisites

- Node.js 20+ 
- Microsoft 365 account with Exchange Online
- Azure AD application registration (for authentication)

## Setup

### 1. Azure Application Registration

1. Go to [Azure Portal](https://portal.azure.com) → Azure Active Directory → App registrations
2. Click "New registration"
3. Enter a name (e.g., "MCP Outlook Scheduler")
4. Select "Accounts in this organizational directory only" for single tenant, or "Accounts in any organizational directory" for multi-tenant
5. Click "Register"

### 2. Configure API Permissions

1. In your app registration, go to "API permissions"
2. Click "Add a permission" → "Microsoft Graph"
3. Add these **Application permissions** (for app-only auth) or **Delegated permissions** (for user auth):
   - `Calendars.Read` - Read calendar availability
   - `Calendars.ReadWrite` - Create and manage calendar events
   - `User.Read` - Read user information
   - `OnlineMeetings.ReadWrite` - Create Teams meetings (optional)

4. Click "Grant admin consent" (requires admin privileges)

### 3. Get Application Credentials

1. Go to "Certificates & secrets" → "Client secrets"
2. Create a new client secret and copy the value
3. Copy the Application (client) ID from the Overview page
4. Copy the Directory (tenant) ID from the Overview page

### 4. Environment Configuration

1. Copy `env.example` to `.env`
2. Fill in your Azure application details:

```bash
# Microsoft Graph Configuration
GRAPH_TENANT_ID=your_tenant_id_or_common
GRAPH_CLIENT_ID=your_app_id
GRAPH_CLIENT_SECRET=your_client_secret
GRAPH_AUTH_MODE=app  # or 'delegated' for user auth

# Default Organizer
ORGANIZER_EMAIL=organizer@yourdomain.com

# Time Zone and Server Settings
DEFAULT_TIMEZONE=Asia/Dubai
PORT=7337

# Token Cache
TOKEN_CACHE_PATH=.tokens.json

# Logging
LOG_LEVEL=info
```

### 5. Install Dependencies

```bash
npm install
```

### 6. Build and Run

```bash
# Development mode
npm run dev

# Build and run production
npm run build
npm start
```

## MCP Tools

The server exposes the following tools:

### 1. `health_check`

Check server health and authentication status.

**Input**: None

**Output**:
```json
{
  "ok": true,
  "graphScopes": ["Calendars.Read", "Calendars.ReadWrite"],
  "organizer": "organizer@example.com"
}
```

### 2. `get_availability`

Get free/busy availability for multiple participants.

**Input**:
```json
{
  "participants": ["user1@example.com", "user2@example.com"],
  "windowStart": "2025-01-15T09:00:00+04:00",
  "windowEnd": "2025-01-15T17:00:00+04:00",
  "granularityMinutes": 30,
  "workHoursOnly": true,
  "timeZone": "Asia/Dubai"
}
```

**Output**:
```json
{
  "timeZone": "Asia/Dubai",
  "users": [
    {
      "email": "user1@example.com",
      "workingHours": {
        "start": "09:00",
        "end": "17:00",
        "days": [1, 2, 3, 4, 5]
      },
      "busy": [
        {
          "start": "2025-01-15T10:00:00+04:00",
          "end": "2025-01-15T11:00:00+04:00",
          "subject": "Team Meeting"
        }
      ],
      "free": [
        {
          "start": "2025-01-15T09:00:00+04:00",
          "end": "2025-01-15T10:00:00+04:00"
        }
      ]
    }
  ]
}
```

### 3. `propose_meeting_times`

Find available meeting time slots for participants.

**Input**:
```json
{
  "participants": ["user1@example.com", "user2@example.com", "user3@example.com"],
  "durationMinutes": 45,
  "windowStart": "2025-01-15T06:00:00+04:00",
  "windowEnd": "2025-01-20T20:00:00+04:00",
  "maxCandidates": 5,
  "bufferBeforeMinutes": 10,
  "bufferAfterMinutes": 10,
  "workHoursOnly": true,
  "minRequiredAttendees": 3,
  "timeZone": "Asia/Dubai"
}
```

**Output**:
```json
{
  "source": "graph_findMeetingTimes",
  "candidates": [
    {
      "start": "2025-01-17T10:00:00+04:00",
      "end": "2025-01-17T10:45:00+04:00",
      "attendeeAvailability": {
        "user1@example.com": "free",
        "user2@example.com": "free",
        "user3@example.com": "free"
      },
      "confidence": 1.0
    }
  ]
}
```

### 4. `book_meeting`

Create a calendar event and invite attendees.

**Input**:
```json
{
  "subject": "Q3 Planning Sync",
  "participants": ["user1@example.com", "user2@example.com", "user3@example.com"],
  "start": "2025-01-17T10:00:00+04:00",
  "end": "2025-01-17T10:45:00+04:00",
  "onlineMeeting": true,
  "location": "Teams",
  "allowConflicts": false,
  "remindersMinutesBeforeStart": 10
}
```

**Output**:
```json
{
  "eventId": "AAMkAGI2TG93AAA=",
  "iCalUid": "040000008200E00074C5B7101A82E00800000000",
  "webLink": "https://outlook.office365.com/owa/?itemid=AAMkAGI2TG93AAA%3D",
  "organizer": "organizer@example.com"
}
```

### 5. `cancel_meeting`

Cancel an existing meeting.

**Input**:
```json
{
  "eventId": "AAMkAGI2TG93AAA=",
  "comment": "Meeting cancelled due to scheduling conflict"
}
```

**Output**:
```json
{
  "cancelled": true,
  "eventId": "AAMkAGI2TG93AAA="
}
```

## Scheduling Logic

### Primary Method: Microsoft Graph findMeetingTimes

The server first attempts to use Microsoft Graph's `findMeetingTimes` API, which provides intelligent meeting scheduling with:
- Conflict detection
- Working hours consideration
- Attendee preferences
- Travel time calculation (when available)

### Fallback Method: Local Intersection

When `findMeetingTimes` is unavailable (due to tenant policies, insufficient permissions, or API limitations), the server falls back to local intersection logic:

1. Fetch free/busy data for each participant
2. Apply buffer times and working hour constraints
3. Compute intersecting free time slots
4. Rank candidates by:
   - Earliest start time
   - Least fragmented calendars
   - Maximum attendee availability

## Configuration Options

### Authentication Modes

- **Delegated**: User authentication via device code flow (recommended for personal use)
- **Application**: Service principal authentication (recommended for production/automation)

### Timezone Handling

- All internal calculations use UTC
- Input/output times include timezone offsets (ISO 8601 format)
- Working hours are interpreted in the specified timezone

### Working Hours

- Default: Monday-Friday, 9:00 AM - 5:00 PM
- Can be overridden per user via mailbox settings
- Configurable via `workHoursOnly` parameter

## Error Handling

The server provides consistent error responses:

```json
{
  "error": {
    "code": "ERROR_CODE",
    "message": "Human-readable error message",
    "details": "Additional error context"
  }
}
```

Common error codes:
- `INVALID_INPUT`: Input validation failed
- `AUTHENTICATION_FAILED`: Microsoft Graph authentication error
- `INSUFFICIENT_PERMISSIONS`: Missing required API permissions
- `SCHEDULING_CONFLICT`: No common free time slots found
- `BOOKING_FAILED`: Failed to create calendar event

## Development

### Running Tests

```bash
# Run all tests
npm test

# Run tests in watch mode
npm run test:watch

# Run tests with coverage
npm run test -- --coverage
```

### Code Quality

```bash
# Lint code
npm run lint

# Format code
npm run format
```

### Building

```bash
# Build TypeScript
npm run build

# Clean build artifacts
npm run clean
```

## Architecture

```
src/
├── index.ts              # Server bootstrap
├── mcp.ts                # MCP tool registration
├── config.ts             # Configuration management
├── types.ts              # Zod schemas and types
├── graph/
│   ├── auth.ts          # OAuth authentication
│   └── client.ts        # Graph client factory
├── scheduling/
│   ├── availability.ts  # Free/busy data processing
│   ├── intersect.ts     # Time slot intersection
│   ├── find.ts          # Meeting time discovery
│   └── book.ts          # Calendar event management
└── utils/
    └── time.ts          # Timezone and interval utilities
```

## Troubleshooting

### Authentication Issues

1. **Invalid client secret**: Regenerate the client secret in Azure
2. **Insufficient permissions**: Ensure admin consent is granted for all required permissions
3. **Tenant restrictions**: Check if your tenant allows the required Graph API calls

### Scheduling Issues

1. **No meeting times found**: 
   - Verify participants have accessible calendars
   - Check if working hours are too restrictive
   - Ensure the time window is reasonable

2. **Graph API fallback**: 
   - The server automatically falls back to local intersection when `findMeetingTimes` fails
   - Check logs for fallback reasons

### Performance Issues

1. **Slow availability checks**: 
   - Reduce the number of participants
   - Use smaller time windows
   - Consider caching availability data

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests for new functionality
5. Ensure all tests pass
6. Submit a pull request

## License

MIT License - see LICENSE file for details.

## Support

For issues and questions:
1. Check the troubleshooting section
2. Review Microsoft Graph API documentation
3. Open an issue on GitHub

## Changelog

### v1.0.0
- Initial release
- MCP server with 5 core tools
- Microsoft Graph integration
- Local intersection fallback
- Comprehensive test coverage
