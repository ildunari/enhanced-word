#!/bin/bash

# Enhanced Word MCP Server - ChatGPT Integration Setup Script
# This script sets up the Enhanced Word MCP Server for use with ChatGPT and other AI tools

set -e

echo "🚀 Enhanced Word MCP Server - ChatGPT Integration Setup"
echo "======================================================"
echo ""

PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_DIR="$PROJECT_DIR/venv"
PYTHON_CMD=""

# Function to detect Python
detect_python() {
    if command -v python3.11 &> /dev/null; then
        PYTHON_CMD="python3.11"
    elif command -v python3.10 &> /dev/null; then
        PYTHON_CMD="python3.10"
    elif command -v python3.9 &> /dev/null; then
        PYTHON_CMD="python3.9"
    elif command -v python3 &> /dev/null; then
        PYTHON_CMD="python3"
    elif command -v python &> /dev/null; then
        PYTHON_CMD="python"
    else
        echo "❌ Error: Python 3.9+ is required but not found"
        echo "   Please install Python 3.9 or higher from https://python.org"
        exit 1
    fi
    
    # Verify Python version
    PYTHON_VERSION=$($PYTHON_CMD --version 2>&1 | awk '{print $2}')
    echo "✅ Found Python $PYTHON_VERSION at $(which $PYTHON_CMD)"
}

# Function to create virtual environment
setup_virtual_env() {
    echo ""
    echo "📦 Setting up Python virtual environment..."
    
    if [ -d "$VENV_DIR" ]; then
        echo "   Virtual environment already exists at $VENV_DIR"
        echo "   Removing existing environment..."
        rm -rf "$VENV_DIR"
    fi
    
    echo "   Creating new virtual environment..."
    $PYTHON_CMD -m venv "$VENV_DIR"
    
    # Activate virtual environment
    source "$VENV_DIR/bin/activate"
    
    # Upgrade pip
    echo "   Upgrading pip..."
    pip install --upgrade pip > /dev/null 2>&1
    
    echo "✅ Virtual environment created and activated"
}

# Function to install dependencies
install_dependencies() {
    echo ""
    echo "📋 Installing Python dependencies..."
    
    if [ -f "$PROJECT_DIR/requirements.txt" ]; then
        echo "   Installing from requirements.txt..."
        pip install -r "$PROJECT_DIR/requirements.txt"
    else
        echo "   Installing core dependencies..."
        pip install python-docx fastmcp websockets
    fi
    
    echo "✅ Dependencies installed successfully"
}

# Function to test installation
test_installation() {
    echo ""
    echo "🧪 Testing installation..."
    
    # Test Python imports
    echo "   Testing Python imports..."
    $PYTHON_CMD -c "
import sys
sys.path.insert(0, '$PROJECT_DIR')
try:
    from word_document_server.main import main
    print('   ✅ MCP server imports working')
except ImportError as e:
    print(f'   ❌ Import error: {e}')
    sys.exit(1)
"
    
    # Test server startup (dry run)
    echo "   Testing server startup..."
    cd "$PROJECT_DIR"
    timeout 5s $PYTHON_CMD -m word_document_server.main --help > /dev/null 2>&1 || true
    
    echo "✅ Installation test completed"
}

# Function to create startup script
create_startup_script() {
    echo ""
    echo "📝 Creating startup script..."
    
    cat > "$PROJECT_DIR/start-server.sh" << 'EOF'
#!/bin/bash

# Enhanced Word MCP Server Startup Script
# Use this script to start the MCP server for ChatGPT integration

PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_DIR="$PROJECT_DIR/venv"

echo "🚀 Starting Enhanced Word MCP Server..."

# Activate virtual environment if it exists
if [ -d "$VENV_DIR" ]; then
    echo "   Activating virtual environment..."
    source "$VENV_DIR/bin/activate"
fi

# Change to project directory
cd "$PROJECT_DIR"

# Start the MCP server
echo "   Starting MCP server on stdio transport..."
echo "   Server is ready for ChatGPT integration!"
echo ""

python -m word_document_server.main "$@"
EOF
    
    chmod +x "$PROJECT_DIR/start-server.sh"
    echo "✅ Startup script created at $PROJECT_DIR/start-server.sh"
}

# Function to create ChatGPT configuration
create_chatgpt_config() {
    echo ""
    echo "⚙️  Creating ChatGPT configuration..."
    
    cat > "$PROJECT_DIR/chatgpt-mcp-config.json" << EOF
{
  "mcpServers": {
    "enhanced-word": {
      "command": "$PROJECT_DIR/start-server.sh",
      "args": [],
      "env": {}
    }
  }
}
EOF
    
    echo "✅ ChatGPT MCP configuration created at $PROJECT_DIR/chatgpt-mcp-config.json"
}

# Function to display setup instructions
display_instructions() {
    echo ""
    echo "🎉 Setup completed successfully!"
    echo "================================="
    echo ""
    echo "📋 Next steps for ChatGPT integration:"
    echo ""
    echo "1. **Copy the MCP configuration:**"
    echo "   File location: $PROJECT_DIR/chatgpt-mcp-config.json"
    echo ""
    echo "2. **Add to your ChatGPT MCP settings:**"
    echo "   - Open ChatGPT settings"
    echo "   - Navigate to 'Advanced' → 'Model Context Protocol'"
    echo "   - Add the configuration from chatgpt-mcp-config.json"
    echo ""
    echo "3. **Manual server start (if needed):**"
    echo "   cd $PROJECT_DIR"
    echo "   ./start-server.sh"
    echo ""
    echo "4. **Test the integration:**"
    echo "   - Start a new ChatGPT conversation"
    echo "   - Ask ChatGPT to help with Word document tasks"
    echo "   - The enhanced-word tools should be available"
    echo ""
    echo "🔧 Available tools include:"
    echo "   • Document creation and editing"
    echo "   • Advanced search and replace (with regex)"
    echo "   • Section management and analysis"
    echo "   • Comments and track changes"
    echo "   • Formatting and styling"
    echo "   • Tables and images"
    echo "   • PDF conversion"
    echo ""
    echo "📚 For detailed usage, see: $PROJECT_DIR/README_ENHANCED.md"
    echo ""
    echo "🆘 Need help? Check the documentation or file an issue at:"
    echo "   https://github.com/ildunari/enhanced-word/issues"
    echo ""
}

# Main execution
main() {
    detect_python
    setup_virtual_env
    install_dependencies
    test_installation
    create_startup_script
    create_chatgpt_config
    display_instructions
}

# Run main function
main