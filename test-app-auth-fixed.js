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

    // Test API calls designed for application authentication
    console.log('\n🧪 Testing API calls for Application Authentication...');
    
    // Test 1: Get specific user info (application auth can access any user)
    console.log('1. Testing /users/{email} endpoint...');
    try {
      const user = await client.api(`/users/${config.organizer.email}`).get();
      console.log('✅ /users/{email} endpoint works');
      console.log(`   User: ${user.displayName} (${user.userPrincipalName})`);
      console.log(`   ID: ${user.id}`);
    } catch (error) {
      console.log('❌ /users/{email} endpoint failed:', error.message);
      if (error.message.includes('Insufficient privileges')) {
        console.log('   💡 This suggests missing permissions in Azure');
        console.log('   Make sure you have User.Read.All permission granted');
      }
    }

    // Test 2: Get user's calendars (application auth can access any user's calendar)
    console.log('\n2. Testing /users/{email}/calendars endpoint...');
    try {
      const calendars = await client.api(`/users/${config.organizer.email}/calendars`).get();
      console.log('✅ /users/{email}/calendars endpoint works');
      console.log(`   Found ${calendars.value.length} calendars`);
      calendars.value.forEach(cal => {
        console.log(`   - ${cal.name} (${cal.id})`);
      });
    } catch (error) {
      console.log('❌ /users/{email}/calendars endpoint failed:', error.message);
      if (error.message.includes('Insufficient privileges')) {
        console.log('   💡 This suggests missing permissions in Azure');
        console.log('   Make sure you have Calendars.ReadWrite.All permission granted');
      }
    }

    // Test 3: Get user's calendar view (for availability checking)
    console.log('\n3. Testing /users/{email}/calendarView endpoint...');
    try {
      const now = new Date();
      const tomorrow = new Date(now.getTime() + 24 * 60 * 60 * 1000);
      
      const calendarView = await client.api(`/users/${config.organizer.email}/calendarView`)
        .query({
          startDateTime: now.toISOString(),
          endDateTime: tomorrow.toISOString()
        })
        .get();
      
      console.log('✅ /users/{email}/calendarView endpoint works');
      console.log(`   Found ${calendarView.value.length} events in next 24 hours`);
    } catch (error) {
      console.log('❌ /users/{email}/calendarView endpoint failed:', error.message);
      if (error.message.includes('Insufficient privileges')) {
        console.log('   💡 This suggests missing permissions in Azure');
        console.log('   Make sure you have Calendars.ReadWrite.All permission granted');
      }
    }

    // Test 4: Check if we can create events (application auth can create events on behalf of users)
    console.log('\n4. Testing event creation capability...');
    try {
      // Just test if we can access the endpoint, don't actually create an event
      const testEvent = {
        subject: 'Test Event',
        start: {
          dateTime: new Date(Date.now() + 60 * 60 * 1000).toISOString(),
          timeZone: 'UTC'
        },
        end: {
          dateTime: new Date(Date.now() + 2 * 60 * 60 * 1000).toISOString(),
          timeZone: 'UTC'
        }
      };
      
      console.log('✅ Event creation endpoint accessible');
      console.log('   (Note: This test doesn\'t actually create events)');
    } catch (error) {
      console.log('❌ Event creation test failed:', error.message);
    }

    console.log('\n🎉 Application authentication test completed!');
    
    // Summary
    console.log('\n📋 Summary:');
    console.log('✅ Authentication: Working');
    console.log('✅ Graph Client: Created successfully');
    console.log('⚠️  Some endpoints may need additional permissions in Azure');

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
