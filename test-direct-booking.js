import { MCPServer } from './dist/mcp.js';
import { logger } from './dist/config.js';

async function testDirectBooking() {
  try {
    console.log('Testing direct MCP server booking functionality...');
    
    // Create an instance of the MCP server
    const server = new MCPServer();
    
    // Test meeting data
    const testMeeting = {
      subject: "Direct Test Meeting - MCP Server",
      participants: ["chitsimran_singh@masaood.com"],
      start: "2025-08-16T14:00:00.000Z", // Tomorrow at 2 PM UTC
      end: "2025-08-16T14:30:00.000Z",   // 30 minutes duration
      organizer: "chitsimran_singh@masaood.com",
      bodyHtml: "<p>This is a direct test of the MCP server booking functionality.</p>",
      onlineMeeting: true,
      remindersMinutesBeforeStart: 15
    };
    
    console.log('Test meeting data:', JSON.stringify(testMeeting, null, 2));
    
    // Test the health check first
    console.log('\n--- Testing health_check tool ---');
    const healthResult = await server.handleHealthCheck({});
    console.log('Health check result:', healthResult);
    
    // Test basic authentication with a simple API call
    console.log('\n--- Testing basic authentication ---');
    try {
      // This should test if we can make basic API calls
      const basicAuthTest = await server.handleBookMeeting({
        subject: "Basic Auth Test",
        participants: ["chitsimran_singh@masaood.com"],
        start: "2025-08-16T15:00:00.000Z",
        end: "2025-08-16T15:30:00.000Z",
        organizer: "chitsimran_singh@masaood.com",
        allowConflicts: true // Skip conflict checking to isolate the issue
      });
      console.log('Basic auth test successful:', basicAuthTest);
    } catch (error) {
      console.log('Basic auth test failed:', error.message);
      if (error.message.includes('authentication') || error.message.includes('permission')) {
        console.log('\nThis looks like a permission issue. Please check:');
        console.log('1. Your Azure app has APPLICATION permissions (not delegated)');
        console.log('2. Admin consent has been granted for all permissions');
        console.log('3. The correct tenant ID and client ID are configured');
      }
    }
    
    // Test the book_meeting tool
    console.log('\n--- Testing book_meeting tool ---');
    console.log('Note: Using application permissions - no user interaction required.');
    console.log('The test will proceed with service principal authentication...');
    
    try {
      const bookingResult = await server.handleBookMeeting(testMeeting);
      console.log('Booking result:', bookingResult);
      console.log('\nDirect test completed successfully!');
    } catch (error) {
      if (error.message.includes('authentication') || error.message.includes('permission')) {
        console.log('\nThis looks like a permission issue. Please check:');
        console.log('1. Your Azure app has APPLICATION permissions (not delegated)');
        console.log('2. Admin consent has been granted for all permissions');
        console.log('3. The correct tenant ID and client ID are configured');
      } else {
        console.error('Error during booking:', error);
      }
    }
    
  } catch (error) {
    console.error('Error during direct test:', error);
    logger.error('Direct test failed:', error);
  }
}

// Run the test
testDirectBooking().catch(console.error);
