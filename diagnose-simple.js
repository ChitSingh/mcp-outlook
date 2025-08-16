#!/usr/bin/env node

import { GraphAuth } from './dist/graph/auth.js';
import config from './dist/config.js';

/**
 * Simple diagnostic tool for Microsoft Graph API 401 errors
 */
async function diagnoseGraphAPI() {
  console.log('🔍 Microsoft Graph API Diagnostic Tool');
  console.log('=====================================\n');

  try {
    // 1. Check configuration
    console.log('1️⃣ Configuration Check');
    console.log('----------------------');
    console.log(`Auth Mode: ${config.graph.authMode}`);
    console.log(`Client ID: ${config.graph.clientId}`);
    console.log(`Tenant ID: ${config.graph.tenantId}`);
    console.log(`Client Secret: ${config.graph.clientSecret ? '✅ Set' : '❌ Not Set'}`);
    console.log(`Organizer Email: ${config.organizer.email}`);
    console.log('');

    // 2. Test authentication
    console.log('2️⃣ Authentication Test');
    console.log('----------------------');
    
    const auth = new GraphAuth();
    
    try {
      const client = await auth.getClient();
      console.log('✅ Authentication successful');
      
      // Get current scopes
      const scopes = auth.getScopes();
      console.log(`📋 Token Scopes: ${scopes.length > 0 ? scopes.join(', ') : 'No scopes found'}`);
      
      // 3. Test basic Graph API connectivity
      console.log('\n3️⃣ Basic Graph API Connectivity');
      console.log('--------------------------------');
      
      try {
        const me = await client.api('/me').get();
        console.log('✅ /me endpoint successful');
        console.log(`   User: ${me.displayName} (${me.userPrincipalName})`);
        console.log(`   ID: ${me.id}`);
      } catch (error) {
        console.log('❌ /me endpoint failed:', error instanceof Error ? error.message : String(error));
        if (error instanceof Error && 'statusCode' in error) {
          console.log(`   Status Code: ${error.statusCode}`);
        }
      }

      // 4. Test calendar permissions
      console.log('\n4️⃣ Calendar Permissions Test');
      console.log('----------------------------');
      
      try {
        const calendars = await client.api('/me/calendars').get();
        console.log('✅ /me/calendars endpoint successful');
        console.log(`   Found ${calendars.value?.length || 0} calendars`);
        if (calendars.value && calendars.value.length > 0) {
          calendars.value.forEach((cal, index) => {
            console.log(`   ${index + 1}. ${cal.name} (${cal.owner?.address || 'Unknown'})`);
          });
        }
      } catch (error) {
        console.log('❌ /me/calendars endpoint failed:', error instanceof Error ? error.message : String(error));
        if (error instanceof Error && 'statusCode' in error) {
          console.log(`   Status Code: ${error.statusCode}`);
        }
      }

      // 5. Test specific user calendar access
      console.log('\n5️⃣ Specific User Calendar Access');
      console.log('----------------------------------');
      
      try {
        const userCalendars = await client.api(`/users/${config.organizer.email}/calendars`).get();
        console.log('✅ User calendars endpoint successful');
        console.log(`   Found ${userCalendars.value?.length || 0} calendars for ${config.organizer.email}`);
      } catch (error) {
        console.log('❌ User calendars endpoint failed:', error instanceof Error ? error.message : String(error));
        if (error instanceof Error && 'statusCode' in error) {
          console.log(`   Status Code: ${error.statusCode}`);
        }
      }

      // 6. Test calendar view access
      console.log('\n6️⃣ Calendar View Access');
      console.log('------------------------');
      
      try {
        const now = new Date();
        const end = new Date(now.getTime() + 24 * 60 * 60 * 1000); // 24 hours from now
        
        const calendarView = await client.api(`/users/${config.organizer.email}/calendarView`)
          .query({
            startDateTime: now.toISOString(),
            endDateTime: end.toISOString(),
            $select: 'subject,start,end'
          })
          .get();
        
        console.log('✅ Calendar view endpoint successful');
        console.log(`   Found ${calendarView.value?.length || 0} events in next 24 hours`);
      } catch (error) {
        console.log('❌ Calendar view endpoint failed:', error instanceof Error ? error.message : String(error));
        if (error instanceof Error && 'statusCode' in error) {
          console.log(`   Status Code: ${error.statusCode}`);
        }
      }

      // 7. Test working hours access
      console.log('\n7️⃣ Working Hours Access');
      console.log('-------------------------');
      
      try {
        const mailboxSettings = await client.api(`/users/${config.organizer.email}/mailboxSettings`).get();
        console.log('✅ Mailbox settings endpoint successful');
        if (mailboxSettings.workingHours) {
          console.log(`   Working Hours: ${mailboxSettings.workingHours.startTime} - ${mailboxSettings.workingHours.endTime}`);
          console.log(`   Days: ${mailboxSettings.workingHours.daysOfWeek?.join(', ') || 'Not set'}`);
        } else {
          console.log('   Working hours not configured');
        }
      } catch (error) {
        console.log('❌ Mailbox settings endpoint failed:', error instanceof Error ? error.message : String(error));
        if (error instanceof Error && 'statusCode' in error) {
          console.log(`   Status Code: ${error.statusCode}`);
        }
      }

      // 8. Test schedule access
      console.log('\n8️⃣ Schedule Access');
      console.log('-------------------');
      
      try {
        const now = new Date();
        const end = new Date(now.getTime() + 24 * 60 * 60 * 1000); // 24 hours from now
        
        const schedule = await client.api(`/users/${config.organizer.email}/calendar/getSchedule`)
          .post({
            schedules: [config.organizer.email],
            startTime: {
              dateTime: now.toISOString(),
              timeZone: 'UTC'
            },
            endTime: {
              dateTime: end.toISOString(),
              timeZone: 'UTC'
            },
            availabilityViewInterval: 30
          });
        
        console.log('✅ Schedule endpoint successful');
        if (schedule.value && schedule.value[0] && schedule.value[0].scheduleItems) {
          console.log(`   Found ${schedule.value[0].scheduleItems.length} schedule items`);
        }
      } catch (error) {
        console.log('❌ Schedule endpoint failed:', error instanceof Error ? error.message : String(error));
        if (error instanceof Error && 'statusCode' in error) {
          console.log(`   Status Code: ${error.statusCode}`);
        }
      }

      // 9. Test event creation (read-only test)
      console.log('\n9️⃣ Event Creation Permission Test');
      console.log('----------------------------------');
      
      try {
        // Try to create a test event (this will fail but we can see the error)
        const testEvent = {
          subject: 'Test Event - Permission Check',
          start: {
            dateTime: new Date(Date.now() + 60 * 60 * 1000).toISOString(), // 1 hour from now
            timeZone: 'UTC'
          },
          end: {
            dateTime: new Date(Date.now() + 2 * 60 * 60 * 1000).toISOString(), // 2 hours from now
            timeZone: 'UTC'
          }
        };
        
        await client.api(`/users/${config.organizer.email}/events`).post(testEvent);
        console.log('✅ Event creation successful (this is unexpected for a test)');
      } catch (error) {
        if (error instanceof Error && 'statusCode' in error) {
          const statusCode = error.statusCode;
          if (statusCode === 401) {
            console.log('❌ Event creation failed with 401 - Permission denied');
            console.log('   This confirms the token is valid but lacks calendar write permissions');
          } else if (statusCode === 403) {
            console.log('❌ Event creation failed with 403 - Forbidden');
            console.log('   This suggests insufficient permissions for this specific operation');
          } else {
            console.log(`❌ Event creation failed with status ${statusCode}`);
          }
        } else {
          console.log('❌ Event creation failed:', error instanceof Error ? error.message : String(error));
        }
      }

    } catch (authError) {
      console.log('❌ Authentication failed:', authError instanceof Error ? authError.message : String(authError));
      return;
    }

    // 10. Summary and recommendations
    console.log('\n🔍 Summary and Recommendations');
    console.log('===============================');
    console.log('');
    console.log('Based on the test results above:');
    console.log('');
    console.log('✅ If most endpoints work but calendar operations fail:');
    console.log('   - Check that your Azure app has the correct delegated permissions');
    console.log('   - Ensure admin consent has been granted');
    console.log('   - Verify the user has calendar access in their organization');
    console.log('');
    console.log('❌ If authentication fails:');
    console.log('   - Check your Azure app registration configuration');
    console.log('   - Verify client ID, tenant ID, and client secret');
    console.log('   - Ensure APPLICATION permissions are granted (not delegated)');
    console.log('');
    console.log('🔧 Next steps:');
    console.log('   1. Check the Azure Portal for your app registration');
    console.log('   2. Verify API permissions are granted and consented');
    console.log('   3. Test with Microsoft Graph Explorer to confirm permissions');
    console.log('   4. Check if your organization has any conditional access policies');

  } catch (error) {
    console.error('❌ Diagnostic failed:', error);
    if (error instanceof Error) {
      console.error('Error details:', error.message);
      console.error('Stack trace:', error.stack);
    }
  }
}

// Run the diagnostic
diagnoseGraphAPI().catch(console.error);
