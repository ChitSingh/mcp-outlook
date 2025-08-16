import { GraphAuth } from './src/graph/auth.js';
import config from './src/config.js';

async function testAppAuth() {
  console.log('🔐 Testing Application Authentication');
  console.log('=====================================\n');

  console.log('Current configuration:');
  console.log(`- Auth Mode: ${config.graph.authMode}`);
  console.log(`- Tenant ID: ${config.graph.tenantId}`);
  console.log(`- Client ID: ${config.graph.clientId}`);
  console.log(`- Client Secret: ${config.graph.clientSecret ? '✅ Set' : '❌ Missing'}`);
  console.log(`- Organizer Email: ${config.organizer.email}\n`);

  if (config.graph.authMode !== 'app') {
    console.log('❌ GRAPH_AUTH_MODE is not set to "app"');
    console.log('Please update your .env file to use GRAPH_AUTH_MODE=app');
    return;
  }

  if (!config.graph.clientSecret) {
    console.log('❌ GRAPH_CLIENT_SECRET is missing');
    console.log('Please add your client secret to the .env file');
    return;
  }

  try {
    console.log('🔄 Creating GraphAuth instance...');
    const auth = new GraphAuth();
    
    console.log('🔄 Getting Graph client...');
    const client = await auth.getClient();
    console.log('✅ Graph client created successfully');

    // Test API calls
    console.log('\n🧪 Testing API calls...');
    
    // Test 1: Get current user info
    console.log('1. Testing /me endpoint...');
    try {
      const me = await client.api('/me').get();
      console.log('✅ /me endpoint works');
      console.log(`   User: ${me.displayName} (${me.userPrincipalName})`);
    } catch (error) {
      console.log('❌ /me endpoint failed:', error.message);
    }

    // Test 2: Get calendars
    console.log('\n2. Testing /me/calendars endpoint...');
    try {
      const calendars = await client.api('/me/calendars').get();
      console.log('✅ /me/calendars endpoint works');
      console.log(`   Found ${calendars.value.length} calendars`);
    } catch (error) {
      console.log('❌ /me/calendars endpoint failed:', error.message);
    }

    // Test 3: Get specific user (for organizer)
    console.log('\n3. Testing /users/{email} endpoint...');
    try {
      const user = await client.api(`/users/${config.organizer.email}`).get();
      console.log('✅ /users/{email} endpoint works');
      console.log(`   User: ${user.displayName} (${user.userPrincipalName})`);
    } catch (error) {
      console.log('❌ /users/{email} endpoint failed:', error.message);
    }

    console.log('\n🎉 Application authentication test completed!');

  } catch (error) {
    console.error('\n❌ Error during authentication:', error.message);
    
    if (error.message.includes('401')) {
      console.log('\n💡 This looks like a permission issue.');
      console.log('Make sure you have:');
      console.log('1. Added APPLICATION permissions (not delegated)');
      console.log('2. Granted admin consent for all permissions');
      console.log('3. The correct tenant ID and client ID');
    }
  }
}

testAppAuth().catch(console.error);
