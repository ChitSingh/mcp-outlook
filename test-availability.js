import { MCPServer } from './dist/mcp.js';
import { logger } from './dist/config.js';

async function testAvailability() {
  try {
    console.log('Testing MCP server availability functionality...');
    
    // Create an instance of the MCP server
    const server = new MCPServer();
    
    // Calculate next Monday
    const today = new Date();
    const daysUntilMonday = (8 - today.getDay()) % 7; // 0 = Sunday, 1 = Monday, etc.
    const nextMonday = new Date(today);
    nextMonday.setDate(today.getDate() + daysUntilMonday);
    nextMonday.setHours(8, 0, 0, 0); // 8 AM
    
    const nextMondayEnd = new Date(nextMonday);
    nextMondayEnd.setHours(18, 0, 0, 0); // 6 PM
    
    // Test availability data for next Monday
    const availabilityRequest = {
      participants: ["chitsimran_singh@masaood.com"],
      windowStart: nextMonday.toISOString(),
      windowEnd: nextMondayEnd.toISOString(),
      timeZone: "Asia/Dubai",
      granularityMinutes: 30,
      workHoursOnly: true
    };
    
    console.log('Next Monday date:', nextMonday.toLocaleDateString('en-US', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' }));
    console.log('Availability request:', JSON.stringify(availabilityRequest, null, 2));
    
    // Test the health check first
    console.log('\n--- Testing health_check tool ---');
    const healthResult = await server.handleHealthCheck({});
    console.log('Health check result:', healthResult);
    
    // Test the get_availability tool
    console.log('\n--- Testing get_availability tool ---');
    const availabilityResult = await server.handleGetAvailability(availabilityRequest);
    console.log('Availability result:', availabilityResult);
    
    console.log('\nAvailability test completed successfully!');
    
  } catch (error) {
    console.error('Error during availability test:', error);
    logger.error('Availability test failed:', error);
  }
}

// Run the test
testAvailability().catch(console.error);
