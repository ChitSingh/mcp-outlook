import { MCPServer } from './dist/mcp.js';
import { logger } from './dist/config.js';

async function testAvailability() {
  try {
    console.log('Testing MCP server availability functionality...');
    
    // Create an instance of the MCP server
    const server = new MCPServer();
    
    // Test availability data
    const availabilityRequest = {
      participants: ["singh.chitsimran@outlook.com"],
      windowStart: "2025-08-16T08:00:00.000Z", // Tomorrow 8 AM UTC
      windowEnd: "2025-08-16T18:00:00.000Z",   // Tomorrow 6 PM UTC
      timeZone: "Asia/Dubai",
      granularityMinutes: 30,
      workHoursOnly: true
    };
    
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
