import { GraphAuth } from './dist/graph/auth.js';
import { logger } from './dist/config.js';

async function testAuthOnly() {
  try {
    console.log('Testing authentication only...');
    
    const auth = new GraphAuth();
    
    console.log('Attempting to get Graph client...');
    const client = await auth.getClient();
    
    console.log('✅ Authentication successful!');
    console.log('Graph client obtained:', typeof client);
    
  } catch (error) {
    console.error('❌ Authentication failed:');
    console.error('Error type:', typeof error);
    console.error('Error constructor:', error?.constructor?.name);
    
    if (error instanceof Error) {
      console.error('Error name:', error.name);
      console.error('Error message:', error.message);
      console.error('Error stack:', error.stack);
    } else {
      console.error('Raw error:', error);
      console.error('Error toString():', String(error));
    }
  }
}

// Run the test
testAuthOnly().catch(console.error);
