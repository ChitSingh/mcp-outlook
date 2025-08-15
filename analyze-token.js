#!/usr/bin/env node

import { GraphAuth } from './dist/graph/auth.js';
import config from './dist/config.js';

/**
 * Detailed token analysis tool for Microsoft Graph API
 */
async function analyzeToken() {
  console.log('🔐 Detailed Token Analysis for Microsoft Graph API');
  console.log('==================================================\n');

  try {
    const auth = new GraphAuth();
    
    console.log('1️⃣ Getting fresh token...');
    const client = await auth.getClient();
    const scopes = auth.getScopes();
    
    console.log('✅ Token obtained successfully');
    console.log(`📋 Requested Scopes: ${scopes.join(', ')}`);
    
    // Test a simple endpoint to see the actual error
    console.log('\n2️⃣ Testing Calendar Endpoint with Detailed Error Analysis');
    console.log('--------------------------------------------------------');
    
    try {
      const calendars = await client.api('/me/calendars').get();
      console.log('✅ SUCCESS: /me/calendars endpoint worked!');
      console.log(`   Found ${calendars.value?.length || 0} calendars`);
    } catch (error) {
      console.log('❌ FAILED: /me/calendars endpoint');
      console.log(`   Error Type: ${error.constructor.name}`);
      console.log(`   Error Message: ${error.message}`);
      
      // Check for GraphError properties
      if (error.statusCode) {
        console.log(`   Status Code: ${error.statusCode}`);
      }
      if (error.code) {
        console.log(`   Error Code: ${error.code}`);
      }
      if (error.body) {
        console.log(`   Error Body: ${JSON.stringify(error.body, null, 2)}`);
      }
      if (error.headers) {
        console.log(`   Response Headers: ${JSON.stringify(error.headers, null, 2)}`);
      }
      
      // Check for additional error properties
      console.log('\n3️⃣ Additional Error Analysis');
      console.log('-----------------------------');
      console.log(`   Error Keys: ${Object.keys(error).join(', ')}`);
      
      // Try to extract more details from the error
      if (error.message && error.message.includes('401')) {
        console.log('\n🔍 401 Error Analysis:');
        console.log('   - Token is being rejected by Graph API');
        console.log('   - This suggests a permission or consent issue');
        console.log('   - Even though scopes look correct');
      }
    }

    // Test with a different approach - try to get user info first
    console.log('\n4️⃣ Testing User Context');
    console.log('-------------------------');
    
    try {
      const me = await client.api('/me').get();
      console.log('✅ User context successful:');
      console.log(`   Display Name: ${me.displayName}`);
      console.log(`   User Principal Name: ${me.userPrincipalName}`);
      console.log(`   Mail: ${me.mail || 'Not set'}`);
      console.log(`   ID: ${me.id}`);
      
      // Check if user has mailbox
      if (me.mail) {
        console.log(`   Has Mailbox: Yes (${me.mail})`);
      } else {
        console.log('   Has Mailbox: No (this could be the issue!)');
      }
      
    } catch (error) {
      console.log('❌ User context failed:', error.message);
    }

    // Test with explicit user email
    console.log('\n5️⃣ Testing with Explicit User Email');
    console.log('------------------------------------');
    
    try {
      const userCalendars = await client.api(`/users/${config.organizer.email}/calendars`).get();
      console.log('✅ SUCCESS: User calendars with explicit email worked!');
      console.log(`   Found ${userCalendars.value?.length || 0} calendars`);
    } catch (error) {
      console.log('❌ FAILED: User calendars with explicit email');
      console.log(`   Error: ${error.message}`);
      if (error.statusCode) {
        console.log(`   Status Code: ${error.statusCode}`);
      }
    }

    console.log('\n🔍 Summary and Next Steps');
    console.log('==========================');
    console.log('');
    console.log('Based on the detailed analysis above:');
    console.log('');
    console.log('🚨 If you still get 401 errors:');
    console.log('   1. Test the same endpoints in Microsoft Graph Explorer');
    console.log('   2. Check if your account has calendar access in the organization');
    console.log('   3. Verify there are no conditional access policies');
    console.log('   4. Check if your account type supports calendar operations');
    console.log('');
    console.log('🔧 Additional troubleshooting:');
    console.log('   1. Try with a different user account in the same tenant');
    console.log('   2. Check Azure AD user permissions');
    console.log('   3. Verify Exchange Online licensing');
    console.log('   4. Contact your IT admin about calendar access');

  } catch (error) {
    console.error('❌ Token analysis failed:', error);
    if (error instanceof Error) {
      console.error('Error details:', error.message);
      console.error('Stack trace:', error.stack);
    }
  }
}

// Run the token analysis
analyzeToken().catch(console.error);
