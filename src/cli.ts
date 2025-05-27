#!/usr/bin/env node
import yargs from 'yargs';
import { hideBin } from 'yargs/helpers';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
// import { SSEServerTransport } from '@modelcontextprotocol/sdk/server/sse.js';
import { createServer } from './server.js';

const argv = yargs(hideBin(process.argv))
  .command('stdio', 'Start the server with stdio transport')
  .command('sse', 'Start the server with SSE transport')
  .demandCommand(1)
  .help()
  .parseSync();

const command = argv._[0];

async function main() {
  const server = createServer();

  if (command === 'stdio') {
    const transport = new StdioServerTransport();
    await server.connect(transport);
    console.error('Excel MCP Server running with stdio transport');
  } else if (command === 'sse') {
    const port = parseInt(process.env.FASTMCP_PORT || '8000');
    // Note: SSE transport implementation would go here
    // For now, we'll focus on stdio transport
    console.error('SSE transport not yet implemented. Use stdio transport instead.');
    process.exit(1);
  }
}

main().catch((error) => {
  console.error('Server error:', error);
  process.exit(1);
});