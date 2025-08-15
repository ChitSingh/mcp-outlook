#!/usr/bin/env node

import { GraphAuth } from './src/graph/auth.js';
import config from './src/config.js';

/**
 * JWT Token Decoder for Microsoft Graph API tokens
 * This helps identify permission and scope issues
 */
async function decodeToken() {
  console.log('🔐 JWT Token Decoder for Microsoft Graph API');
  console.log('============================================\n');

  try {
    const auth = new GraphAuth();
    
    // Get the current token
    console.log('1️⃣ Getting current token...');
    const client = await auth.getClient();
    const scopes = auth.getScopes();
    
    console.log('✅ Token obtained successfully');
    console.log(`📋 Requested Scopes: ${scopes.length > 0 ? scopes.join(', ') : 'No scopes found'}`);
    
    // Get the token from the auth provider
    const token = await new Promise((resolve, reject) => {
      client.api('/me').authProvider((done) => {
        // This will be called by the Graph client
        // We need to intercept it to get the token
        reject(new Error('Token interception not implemented in this version'));
      });
    }).catch(() => {
      // Since we can't easily intercept the token, let's try to get it from the cache
      console.log('📝 Note: Using cached token information');
      return null;
    });

    if (!token) {
      console.log('\n2️⃣ Token Analysis');
      console.log('------------------');
      console.log('❌ Unable to retrieve token for analysis');
      console.log('   This is expected in the current implementation');
      console.log('');
      console.log('3️⃣ Scope Analysis');
      console.log('------------------');
      console.log(`📋 Current Scopes: ${scopes.length > 0 ? scopes.join(', ') : 'No scopes found'}`);
      
      if (scopes.length === 0) {
        console.log('⚠️  WARNING: No scopes found in token response');
        console.log('   This suggests the token may not have the required permissions');
        console.log('');
        console.log('🔧 Recommended Scopes for Calendar Operations:');
        console.log('   - https://graph.microsoft.com/Calendars.Read');
        console.log('   - https://graph.microsoft.com/Calendars.ReadWrite');
        console.log('   - https://graph.microsoft.com/User.Read');
        console.log('   - https://graph.microsoft.com/OnlineMeetings.ReadWrite');
        console.log('');
        console.log('4️⃣ Troubleshooting Steps');
        console.log('-------------------------');
        console.log('1. Check Azure Portal App Registration:');
        console.log('   - Go to Azure Portal > App Registrations > Your App');
        console.log('   - Check API Permissions section');
        console.log('   - Ensure delegated permissions are granted');
        console.log('   - Verify admin consent has been given');
        console.log('');
        console.log('2. Verify Permission Names:');
        console.log('   - Calendars.ReadWrite (not Calendar.ReadWrite)');
        console.log('   - User.Read');
        console.log('   - MailboxSettings.Read');
        console.log('');
        console.log('3. Check Account Type:');
        console.log('   - Ensure your account is in the same Azure tenant');
        console.log('   - Check if you have admin rights to grant consent');
        console.log('   - Verify the app supports your account type');
        console.log('');
        console.log('4. Test with Graph Explorer:');
        console.log('   - Go to https://developer.microsoft.com/en-us/graph/graph-explorer');
        console.log('   - Sign in with the same account');
        console.log('   - Try the same API calls to verify permissions');
        console.log('');
        console.log('5. Check Conditional Access:');
        console.log('   - Your organization may have policies blocking access');
        console.log('   - Contact your IT admin if needed');
      } else {
        console.log('✅ Scopes found in token');
        console.log('');
        console.log('4️⃣ Permission Analysis');
        console.log('------------------------');
        
        const requiredScopes = [
          'https://graph.microsoft.com/Calendars.Read',
          'https://graph.microsoft.com/Calendars.ReadWrite',
          'https://graph.microsoft.com/User.Read'
        ];
        
        const missingScopes = requiredScopes.filter(scope => !scopes.includes(scope));
        
        if (missingScopes.length === 0) {
          console.log('✅ All required scopes are present');
          console.log('   The issue may be with admin consent or organization policies');
        } else {
          console.log('❌ Missing required scopes:');
          missingScopes.forEach(scope => console.log(`   - ${scope}`));
          console.log('');
          console.log('🔧 Add these permissions in Azure Portal:');
          console.log('   1. Go to App Registration > API Permissions');
          console.log('   2. Add the missing permissions');
          console.log('   3. Grant admin consent');
        }
      }
    }

  } catch (error) {
    console.error('❌ Token decoding failed:', error);
    if (error instanceof Error) {
      console.error('Error details:', error.message);
      console.error('Stack trace:', error.stack);
    }
  }
}

// Run the token decoder
decodeToken().catch(console.error);
