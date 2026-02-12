#!/usr/bin/env node

const { spawn, execFileSync } = require('child_process');
const path = require('path');

// Get the directory where this package is installed
const packageDir = path.dirname(__dirname);

// Function to find the correct Python executable with mcp module
function findPythonWithMCP() {
  // Try environment variable first, then common paths
  const pythonPaths = [
    process.env.PYTHON_PATH,
    process.env.ENHANCED_WORD_PYTHON,
    'python3',
    'python',
    '/usr/bin/python3',
    '/usr/local/bin/python3'
  ].filter(Boolean); // Remove undefined values
  
  for (const pythonPath of pythonPaths) {
    try {
      // Verify required runtime modules (FastMCP + python-docx).
      execFileSync(
        pythonPath,
        ['-c', "from mcp.server.fastmcp import FastMCP; import docx; print('OK')"],
        { stdio: 'pipe', timeout: 5000 }
      );
      return pythonPath;
    } catch (err) {
      continue;
    }
  }
  
  throw new Error(
    "No compatible Python found. Need Python 3.10+ with required deps. " +
    "Install with: pip install -r requirements.txt"
  );
}

try {
  const pythonExecutable = findPythonWithMCP();
  
  // Run the Python MCP server
  const python = spawn(pythonExecutable, ['-m', 'word_document_server.main'], {
    cwd: packageDir,
    stdio: 'inherit',
    env: { ...process.env }
  });

  python.on('close', (code) => {
    process.exit(code);
  });

  python.on('error', (err) => {
    console.error('Failed to start Enhanced Word MCP Server:', err.message);
    console.error('Make sure Python 3.10+ is installed and requirements are met.');
    process.exit(1);
  });

} catch (err) {
  console.error('Setup Error:', err.message);
  console.error('Ensure Python 3.10+ is in your PATH and dependencies are installed.');
  console.error('Install with: pip install -r requirements.txt');
  process.exit(1);
}
