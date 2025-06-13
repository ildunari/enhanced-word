/**
 * Enhanced Word MCP Server - Node.js Entry Point
 * 
 * This is a Node.js wrapper for the Python-based Enhanced Word MCP Server.
 * The actual server logic is implemented in Python using the word_document_server module.
 */

const { spawn, execSync } = require('child_process');
const path = require('path');

/**
 * Find the correct Python executable with mcp module
 * @returns {string} Path to Python executable
 */
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
      // Check if this Python has the mcp module
      execSync(`${pythonPath} -c "import mcp; print('MCP found')"`, { 
        stdio: 'pipe', 
        timeout: 5000 
      });
      return pythonPath;
    } catch (err) {
      continue;
    }
  }
  
  throw new Error('No Python installation found with MCP module. Please install: pip install mcp');
}

/**
 * Start the Enhanced Word MCP Server
 * @param {Object} options - Configuration options
 * @returns {ChildProcess} The spawned Python process
 */
function startServer(options = {}) {
  const packageDir = __dirname;
  
  try {
    const pythonExecutable = findPythonWithMCP();
    
    const python = spawn(pythonExecutable, ['-m', 'word_document_server.main'], {
      cwd: packageDir,
      stdio: options.stdio || 'inherit',
      env: { ...process.env, ...options.env }
    });

    return python;
  } catch (err) {
    console.error('Setup Error:', err.message);
    console.error('Please install the MCP module: pip install mcp');
    console.error('Or ensure Python 3.11+ is in your PATH.');
    throw err;
  }
}

module.exports = {
  startServer
};

// If this file is run directly, start the server
if (require.main === module) {
  try {
    const server = startServer();
    
    server.on('close', (code) => {
      process.exit(code);
    });
    
    server.on('error', (err) => {
      console.error('Failed to start Enhanced Word MCP Server:', err.message);
      console.error('Make sure Python 3.11+ is installed and requirements are met.');
      process.exit(1);
    });
  } catch (err) {
    console.error('Failed to start server:', err.message);
    process.exit(1);
  }
}
