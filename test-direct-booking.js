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
      participants: ["singh.chitsimran@outlook.com"],
      start: "2025-08-16T14:00:00.000Z", // Tomorrow at 2 PM UTC
      end: "2025-08-16T14:30:00.000Z",   // 30 minutes duration
      organizer: "singh.chitsimran@outlook.com",
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
        participants: ["singh.chitsimran@outlook.com"],
        start: "2025-08-16T15:00:00.000Z",
        end: "2025-08-16T15:30:00.000Z",
        organizer: "singh.chitsimran@outlook.com",
        allowConflicts: true // Skip conflict checking to isolate the issue
      });
      console.log('Basic auth test successful:', basicAuthTest);
    } catch (error) {
      console.log('Basic auth test failed:', error.message);
      if (error.message.includes('fetch failed') || error.message.includes('authentication')) {
        console.log('\nAuthentication is in progress. Please check the console above for:');
        console.log('1. A URL to visit (usually https://microsoft.com/devicelogin)');
        console.log('2. A code to enter on that page');
        console.log('3. Complete the authentication in your browser');
        console.log('\nAfter completing authentication, run this test again.');
      }
    }
    
    // Test the book_meeting tool
    console.log('\n--- Testing book_meeting tool ---');
    console.log('Note: If this is the first time running, you may need to complete device code authentication.');
    console.log('Check the console output above for the authentication code and URL.');
    console.log('The test will wait for authentication to complete...');
    
    try {
      const bookingResult = await server.handleBookMeeting(testMeeting);
      console.log('Booking result:', bookingResult);
      console.log('\nDirect test completed successfully!');
    } catch (error) {
      if (error.message.includes('fetch failed') || error.message.includes('authentication')) {
        console.log('\nAuthentication is in progress. Please check the console above for:');
        console.log('1. A URL to visit (usually https://microsoft.com/devicelogin)');
        console.log('2. A code to enter on that page');
        console.log('3. Complete the authentication in your browser');
        console.log('\nAfter completing authentication, run this test again.');
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
