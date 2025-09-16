# FSS Parse Word

**Professional-grade bidirectional parser for Word (.docx) ↔ Markdown (.md) conversion**

Part of the **FSS Parsers** collection - individual parser tools with the `fss-parse-*` CLI prefix. **Completely standalone** - no dependencies on other FSS parsers.

🛡️ **Built with production safety and enterprise quality standards**

## 🚀 Quick Start

### Installation
```bash
# Clone the repository
git clone https://github.com/FSSCoding/fss-parse-word.git
cd fss-parse-word
python3 install.py

# Your tool is now available as 'fss-parse-word'
```

### Basic Usage
```bash
# Convert Word to Markdown
fss-parse-word document.docx document.md

# Convert Markdown to Word  
fss-parse-word document.md document.docx

# Get help
fss-parse-word --help
```

## ✨ Key Features

### 🛡️ **Production Safety**
- **Hash Validation**: SHA256 checksums prevent data corruption
- **Collision Detection**: Prevents overwriting different files with same name
- **Automatic Backups**: Creates `.backup` files before overwriting
- **Confirmation Prompts**: Interactive safety checks before destructive operations
- **Never Destroys Documents**: Multiple safety layers protect your files

### 🎯 **Professional Quality**
- **Word Theme Compatibility**: Uses Word's built-in styles and colors
- **Configurable Formatting**: YAML/JSON config files for custom styling
- **Template Support**: Use corporate Word templates for branding
- **Metadata Preservation**: Maintains formatting information across conversions
- **Enterprise-Grade Output**: Professional documents that look native to Word

### 🔧 **Robust Parsing**
- **Universal Markdown Support**: Handles any markdown file format
- **Rich Formatting**: Bold, italic, headings, lists, tables, code blocks
- **Inline Code & Links**: Preserves technical documentation formatting
- **Table Conversion**: Professional table formatting in both directions
- **Code Block Styling**: Proper syntax highlighting preparation

## 📋 Usage Examples

### Basic Conversion
```bash
# Auto-detect conversion direction
fss-parse-word presentation.docx presentation.md
fss-parse-word notes.md notes.docx
```

### Advanced Options
```bash
# Skip confirmation prompts (for automation)
fss-parse-word --force document.md output.docx

# Skip backup creation
fss-parse-word --no-backup document.md output.docx

# Use custom configuration
fss-parse-word --config corporate_style.yaml document.md branded.docx

# Use Word template for branding
fss-parse-word --template company_template.docx document.md branded.docx
```

### Configuration Management
```bash
# Create sample config file
fss-parse-word --create-config my_style.yaml

# Edit the generated file to customize formatting
# Use with: fss-parse-word --config my_style.yaml input.md output.docx
```

## ⚙️ Configuration Options

### YAML Frontmatter (Per-Document)
```yaml
---
font_name: "Arial"
font_size: 12
heading_colors:
  1: "#2E75B6"
  2: "#C55A11"
use_builtin_styles: true
---

# Your markdown content here
```

### External Config File
```yaml
# Professional Word styling
font_name: "Calibri"
font_size: 11
heading_font: "Calibri"

# Word theme colors
heading_colors:
  1: "#2E75B6"  # Blue
  2: "#C55A11"  # Orange
  3: "#70AD47"  # Green

# Professional spacing
line_spacing: 1.15
paragraph_spacing_after: 6
```

## 🔒 Safety Features

### File Protection
- **Hash Checking**: Detects file changes and prevents accidental overwrites
- **Backup Creation**: Automatically creates backups with `.backup` extension
- **Collision Detection**: Warns when converting would overwrite different content
- **Confirmation Prompts**: Interactive checks for destructive operations

### Safety Flags
```bash
--force           # Skip confirmation prompts (use with caution)
--no-backup       # Skip backup creation
--no-hash-check   # Skip hash validation
```

### Manual Safety Override (for agents/automation)
```bash
# For automated workflows, use explicit safety flags:
fss-parse-word --force --no-backup input.md output.docx  # Maximum automation
fss-parse-word --no-backup input.md output.docx          # Skip backup only
```

## 📁 Installation Details

### Automatic Installation
The installer handles multiple scenarios:
- **User Local**: `~/.local/bin/fss-parse-word` (preferred, no sudo required)
- **System Wide**: `/usr/local/bin/fss-parse-word` (requires sudo)
- **PATH Configuration**: Automatic detection and guidance

### Manual Installation (for agents)
```bash
# Dependencies
pip install python-docx markdown PyYAML

# Copy script
mkdir -p ~/.local/bin
cp src/word_converter.py ~/.local/bin/fss-parse-word
chmod +x ~/.local/bin/fss-parse-word

# Add to PATH (if needed)
echo 'export PATH="$HOME/.local/bin:$PATH"' >> ~/.bashrc
```

### Dependencies
- **Required**: `python-docx`, `markdown`
- **Optional**: `PyYAML` (for YAML configuration files)

## 🧪 Testing & Validation

### Hash Verification
```bash
# The tool displays hash information for verification
fss-parse-word document.md output.docx
# ✅ Successfully converted document.md → output.docx
# 📊 Output hash: a1b2c3d4e5f6g7h8...
```

### Backup Management
```bash
# Backups are created automatically
ls -la
# document.docx
# document.docx.backup     # Created before overwrite
# document.docx.backup.1   # Multiple backups if needed
```

## 🤝 Agent-Friendly Design

### Command-Line Integration
- **Clear Error Codes**: Exit 0 for success, 1 for failure
- **Structured Output**: Consistent formatting for parsing
- **Safety Flags**: Explicit control over interactive features
- **Manual Instructions**: Fallback commands for automated workflows

### Automation Support
```bash
# Non-interactive usage
fss-parse-word --force document.md output.docx

# With hash verification but no prompts
fss-parse-word --force --no-backup input.md output.docx

# Maximum safety (recommended for agents)
fss-parse-word input.md output.docx  # Will prompt and create backups
```

## 📊 Technical Specifications

### File Format Support
- **Input**: `.docx` (Word 2007+), `.md` (CommonMark + extensions)
- **Output**: `.docx` (Word compatible), `.md` (Universal markdown)
- **Configuration**: `.yaml`, `.json`

### Conversion Features
- **Bidirectional**: Full round-trip conversion support
- **Metadata Preservation**: Formatting information maintained
- **Theme Compatibility**: Word built-in styles and colors
- **Professional Quality**: Enterprise-ready output

### Safety Architecture
- **SHA256 Hashing**: File integrity validation
- **Atomic Operations**: All-or-nothing conversions
- **Backup Management**: Automatic and versioned
- **Error Recovery**: Graceful handling of edge cases

---

## 📦 Part of FSS Parsers Collection

This tool uses the `fss-parse-word` CLI command as part of the broader **FSS Parsers** ecosystem. Future parsers will follow the same `fss-parse-*` pattern for consistency.

**Repository**: https://github.com/FSSCoding/fss-parse-word  
**License**: MIT  
**Author**: FssCoding

## 📋 Known Limitations

### Image Handling
- **Images are not currently processed** during conversion (python-docx limitation)
- **Text content extracts perfectly** - images are silently skipped
- **No crashes or errors** with image-heavy documents
- **Recommendation**: Excellent for text-focused documents

### Advanced Elements  
- **Charts and SmartArt**: Not supported
- **Headers/Footers**: Not currently processed
- **Complex tables**: Basic conversion only (merged cells not supported)

---