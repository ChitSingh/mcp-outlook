import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Read the token cache file
const tokenCachePath = path.join(process.cwd(), '.tokens.json');

if (fs.existsSync(tokenCachePath)) {
  try {
    const tokenData = JSON.parse(fs.readFileSync(tokenCachePath, 'utf8'));
    
    if (tokenData.accessToken) {
      console.log('✅ Token found!');
      console.log(`Token expires at: ${new Date(tokenData.expiresAt).toISOString()}`);
      console.log(`Scopes: ${tokenData.scopes.join(', ')}`);
      
      // Check if token is still valid
      const now = Date.now();
      const buffer = 10 * 60 * 1000; // 10 minutes buffer
      
      if (tokenData.expiresAt > (now + buffer)) {
        console.log('✅ Token is valid and can be used');
        console.log('\nTo use this token, set this environment variable:');
        console.log(`GRAPH_ACCESS_TOKEN=${tokenData.accessToken}`);
      } else {
        console.log('❌ Token is expired or will expire soon');
        console.log('You need to re-authenticate');
      }
    } else {
      console.log('❌ No access token found in cache');
    }
  } catch (error) {
    console.error('Error reading token cache:', error.message);
  }
} else {
  console.log('❌ No token cache file found');
  console.log('Run authentication first with: node test-auth-only.js');
}
