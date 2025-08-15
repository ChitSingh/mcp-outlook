import config from './dist/config.js';

console.log('=== MCP Server Configuration ===');
console.log('Graph Configuration:');
console.log('  Tenant ID:', config.graph.tenantId);
console.log('  Client ID:', config.graph.clientId);
console.log('  Client Secret:', config.graph.clientSecret ? '[SET]' : '[NOT SET]');
console.log('  Auth Mode:', config.graph.authMode);

console.log('\nOrganizer Configuration:');
console.log('  Email:', config.organizer.email);

console.log('\nServer Configuration:');
console.log('  Port:', config.server.port);
console.log('  Timezone:', config.server.defaultTimezone);

console.log('\nAuth Configuration:');
console.log('  Token Cache Path:', config.auth.tokenCachePath);

console.log('\nLogging Configuration:');
console.log('  Level:', config.logging.level);

console.log('\nEnvironment Variables:');
console.log('  GRAPH_TENANT_ID:', process.env.GRAPH_TENANT_ID || '[NOT SET]');
console.log('  GRAPH_CLIENT_ID:', process.env.GRAPH_CLIENT_ID || '[NOT SET]');
console.log('  GRAPH_CLIENT_SECRET:', process.env.GRAPH_CLIENT_SECRET ? '[SET]' : '[NOT SET]');
console.log('  GRAPH_AUTH_MODE:', process.env.GRAPH_AUTH_MODE || '[NOT SET]');
console.log('  GRAPH_ORGANIZER_EMAIL:', process.env.GRAPH_ORGANIZER_EMAIL || '[NOT SET]');
