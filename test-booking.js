import { spawn } from 'child_process';
import { join } from 'path';

// Test booking a meeting with yourself
const testMeeting = {
  subject: "Test Meeting - MCP Server Test",
  participants: ["chitsimran_singh@masaood.com"],
  start: "2025-08-16T10:00:00.000Z", // Tomorrow at 10 AM UTC
  end: "2025-08-16T10:30:00.000Z",   // 30 minutes duration
  organizer: "chitsimran_singh@masaood.com",
  bodyHtml: "<p>This is a test meeting to verify the MCP server booking functionality is working correctly.</p>",
  onlineMeeting: true,
  remindersMinutesBeforeStart: 15
};

console.log('Starting MCP server test...');
console.log('Test meeting details:', JSON.stringify(testMeeting, null, 2));

// Start the MCP server
const serverProcess = spawn('node', ['dist/index.js'], {
  stdio: ['pipe', 'pipe', 'pipe']
});

let serverOutput = '';
let serverError = '';

serverProcess.stdout.on('data', (data) => {
  serverOutput += data.toString();
  console.log('Server output:', data.toString());
});

serverProcess.stderr.on('data', (data) => {
  serverError += data.toString();
  console.error('Server error:', data.toString());
});

serverProcess.on('close', (code) => {
  console.log(`Server process exited with code ${code}`);
  console.log('Final server output:', serverOutput);
  if (serverError) {
    console.error('Final server errors:', serverError);
  }
});

// Wait a moment for server to start
setTimeout(() => {
  console.log('Server should be running now. You can test the booking functionality.');
  console.log('To test the MCP server in Claude, make sure it\'s configured to use this server.');
  console.log('The server is running with the following configuration:');
console.log('- Organizer: chitsimran_singh@masaood.com');
console.log('- Tenant ID: 31da92a8-2f5c-44ea-99bd-626d32113f36');
console.log('- Auth Mode: app (application permissions)');
console.log('- Timezone: Asia/Dubai');
}, 2000);

// Keep the process alive
process.on('SIGINT', () => {
  console.log('Shutting down test...');
  serverProcess.kill();
  process.exit(0);
});
