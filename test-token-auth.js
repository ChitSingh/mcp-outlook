import { GraphAuthWithToken } from './src/graph/auth-with-token.js';

async function testTokenAuth() {
  console.log('🔐 Testing Token-Based Authentication');
  console.log('=====================================\n');

  const auth = new GraphAuthWithToken();
  
  // Check token status
  const tokenInfo = auth.getTokenInfo();
  console.log(`Token available: ${tokenInfo.hasToken ? '✅ Yes' : '❌ No'}`);
  if (tokenInfo.hasToken) {
    console.log(`Token preview: ${tokenInfo.tokenPreview}`);
  }

  if (!auth.hasToken()) {
    console.log('\n❌ No token found!');
    console.log('\nTo fix this:');
    console.log('1. Run: node test-auth-only.js (outside of Claude)');
    console.log('2. Copy the token from the output');
    console.log('3. Set environment variable: GRAPH_ACCESS_TOKEN=your_token_here');
    console.log('4. Run this script again');
    return;
  }

  try {
    console.log('\n🔄 Getting Graph client...');
    const client = await auth.getClient();
    console.log('✅ Graph client created successfully');

    // Test a simple API call
    console.log('\n🧪 Testing API call...');
    const me = await client.api('/me').get();
    console.log('✅ API call successful!');
    console.log(`User: ${me.displayName} (${me.userPrincipalName})`);

  } catch (error) {
    console.error('❌ Error:', error.message);
    
    if (error.message.includes('401')) {
      console.log('\n💡 This looks like an authentication error.');
      console.log('The token might be expired or invalid.');
      console.log('Try running authentication again: node test-auth-only.js');
    }
  }
}

testTokenAuth().catch(console.error);
