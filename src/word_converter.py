#!/usr/bin/env python3
"""
FSS Parse Word - Professional document processing toolkit
Safe, hash-validated bidirectional conversion between .docx and .md formats.

CLI Framework: Click-based with universal options
Safety Features: Hash checking, collision detection, confirmation prompts
"""

import click
import hashlib
import json
import re
import sys
import yaml
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple
from dataclasses import dataclass, asdict
import tempfile
import shutil
from datetime import datetime
from rich.console import Console
from rich.table import Table

try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.style import WD_STYLE_TYPE
    from docx.enum.dml import MSO_THEME_COLOR_INDEX
    from docx.oxml.ns import qn
except ImportError:
    print("❌ Error: python-docx not installed. Run: pip install python-docx")
    sys.exit(1)

try:
    import markdown
    from markdown.extensions import codehilite, tables, toc, fenced_code
except ImportError:
    print("❌ Error: markdown not installed. Run: pip install markdown")
    sys.exit(1)

try:
    import yaml
except ImportError:
    print("⚠️  Warning: PyYAML not installed. YAML config support disabled. Install with: pip install PyYAML")
    yaml = None

console = Console()


# Enhanced Error Handling Classes
class WordParserError(Exception):
    """Base exception for Word parser with user-friendly messages"""
    def __init__(self, message: str, suggested_action: str = None, error_code: str = None):
        self.message = message
        self.suggested_action = suggested_action
        self.error_code = error_code
        super().__init__(self.message)

class UnsupportedFormatError(WordParserError):
    """Raised when file format is not supported"""
    pass

class ProtectedFileError(WordParserError):
    """Raised when file is password-protected"""
    pass

class FileSizeError(WordParserError):
    """Raised when file is too large"""
    pass

class CorruptedFileError(WordParserError):
    """Raised when file appears corrupted"""
    pass


@dataclass
class SafetyConfig:
    """Configuration for safety mechanisms."""
    require_confirmation: bool = True
    create_backup: bool = True
    check_hash: bool = True
    prevent_overwrite: bool = True
    backup_suffix: str = ".backup"


@dataclass
class FormatMetadata:
    """Stores formatting information for reconstruction."""
    bold_ranges: List[Tuple[int, int]] = None
    italic_ranges: List[Tuple[int, int]] = None
    heading_levels: Dict[int, int] = None
    lists: List[Dict[str, Any]] = None
    tables: List[Dict[str, Any]] = None
    hyperlinks: List[Dict[str, Any]] = None
    images: List[Dict[str, Any]] = None
    styles: Dict[str, str] = None
    file_hash: str = ""
    conversion_timestamp: str = ""
    
    def __post_init__(self):
        if self.bold_ranges is None:
            self.bold_ranges = []
        if self.italic_ranges is None:
            self.italic_ranges = []
        if self.heading_levels is None:
            self.heading_levels = {}
        if self.lists is None:
            self.lists = []
        if self.tables is None:
            self.tables = []
        if self.hyperlinks is None:
            self.hyperlinks = []
        if self.images is None:
            self.images = []
        if self.styles is None:
            self.styles = {}


@dataclass
class ConversionConfig:
    """Configuration for markdown to Word conversion."""
    # Document settings
    font_name: str = "Calibri"
    font_size: int = 11
    line_spacing: float = 1.15
    
    # Heading settings
    heading_font: str = "Calibri"
    heading_colors: Dict[int, str] = None
    heading_sizes: Dict[int, int] = None
    heading_spacing_before: Dict[int, int] = None
    heading_spacing_after: Dict[int, int] = None
    
    # Paragraph settings
    paragraph_spacing_after: int = 6
    paragraph_first_line_indent: float = 0.0
    
    # List settings
    list_spacing: int = 0
    list_indent: float = 0.25
    
    # Table settings
    table_style: str = "Table Grid"
    table_autofit: bool = True
    
    # Code block settings
    code_font: str = "Consolas"
    code_size: int = 9
    code_background: str = "#F5F5F5"
    
    # Use Word built-in styles
    use_builtin_styles: bool = True
    custom_style_map: Dict[str, str] = None
    
    def __post_init__(self):
        if self.heading_colors is None:
            self.heading_colors = {
                1: "#2E75B6",  # Blue
                2: "#C55A11",  # Orange
                3: "#70AD47",  # Green
                4: "#7030A0",  # Purple
                5: "#264478",  # Dark Blue
                6: "#E7E6E6"   # Gray
            }
        
        if self.heading_sizes is None:
            self.heading_sizes = {
                1: 16, 2: 14, 3: 12, 4: 11, 5: 11, 6: 10
            }
        
        if self.heading_spacing_before is None:
            self.heading_spacing_before = {
                1: 12, 2: 10, 3: 8, 4: 6, 5: 6, 6: 4
            }
        
        if self.heading_spacing_after is None:
            self.heading_spacing_after = {
                1: 6, 2: 6, 3: 4, 4: 4, 5: 2, 6: 2
            }
        
        if self.custom_style_map is None:
            self.custom_style_map = {}


class FileValidator:
    """Enhanced file validation with accessibility-focused error handling"""
    
    MAX_FILE_SIZE = 100 * 1024 * 1024  # 100MB limit
    
    def detect_file_format(self, file_path: str) -> Dict[str, Any]:
        """Detect file format and provide clear feedback"""
        extension = Path(file_path).suffix.lower()
        
        format_info = {
            '.docx': {'supported': True, 'processor': 'python-docx'},
            '.rtf': {'supported': False, 'alternative': 'Please convert to .docx format using Microsoft Word first'},
            '.doc': {'supported': False, 'alternative': 'Please save as .docx format using Microsoft Word first'},
            '.odt': {'supported': False, 'alternative': 'Please export as .docx format using LibreOffice first'},
            '.md': {'supported': True, 'processor': 'markdown'},
            '.txt': {'supported': False, 'alternative': 'Please rename to .md extension if this is markdown content'}
        }
        
        return {
            'format': extension,
            'supported': format_info.get(extension, {}).get('supported', False),
            'message': format_info.get(extension, {}).get('alternative', f'Unknown format: {extension}'),
            'suggested_action': f"Convert {extension} to .docx or .md format" if extension not in format_info else format_info[extension].get('alternative')
        }
    
    def check_file_protection(self, file_path: str) -> Dict[str, Any]:
        """Check if file is password-protected before processing"""
        try:
            if Path(file_path).suffix.lower() == '.docx':
                Document(file_path)
            return {'protected': False, 'accessible': True}
        except Exception as e:
            error_msg = str(e).lower()
            if any(term in error_msg for term in ['password', 'encrypted', 'protected', 'access denied']):
                return {
                    'protected': True, 
                    'accessible': False,
                    'message': 'Document is password-protected or encrypted',
                    'suggested_action': 'Remove password protection in Microsoft Word: File → Info → Protect Document → Encrypt with Password (remove password)'
                }
            return {
                'protected': False,
                'accessible': False, 
                'error': str(e),
                'suggested_action': 'Check if file is corrupted or try opening in Microsoft Word first'
            }
    
    def validate_file_size(self, file_path: str) -> Dict[str, Any]:
        """Check file size before processing"""
        try:
            size = Path(file_path).stat().st_size
            size_mb = round(size / 1024 / 1024, 1)
            limit_mb = round(self.MAX_FILE_SIZE / 1024 / 1024, 1)
            
            if size > self.MAX_FILE_SIZE:
                return {
                    'valid': False,
                    'size_mb': size_mb,
                    'limit_mb': limit_mb,
                    'message': f'File too large ({size_mb}MB). Maximum supported: {limit_mb}MB',
                    'suggested_action': 'Split document into smaller sections using Microsoft Word or try processing in smaller chunks'
                }
            
            return {'valid': True, 'size_mb': size_mb}
        except Exception as e:
            return {
                'valid': False,
                'error': str(e),
                'suggested_action': 'Check if file exists and is accessible'
            }
    
    def analyze_document_structure(self, doc_path: str) -> Dict[str, Any]:
        """Analyze document for embedded objects and complex elements"""
        try:
            doc = Document(doc_path)
            analysis = {
                'total_paragraphs': len(doc.paragraphs),
                'embedded_objects': [],
                'images': 0,
                'tables': len(doc.tables),
                'warnings': [],
                'accessibility_notes': []
            }
            
            # Count images and check for alt text
            image_count = 0
            for rel in doc.part.rels.values():
                if 'image' in rel.target_ref:
                    image_count += 1
            analysis['images'] = image_count
            
            # Generate accessibility warnings
            if image_count > 0:
                analysis['warnings'].append(f"Found {image_count} images")
                analysis['accessibility_notes'].append("Images may lose alt text during conversion. Consider adding descriptions in surrounding text.")
            
            if analysis['tables'] > 0:
                analysis['accessibility_notes'].append("Tables detected. Conversion will preserve structure but may lose complex formatting.")
            
            # Check for complex formatting
            has_complex_formatting = False
            for paragraph in doc.paragraphs[:10]:  # Sample first 10 paragraphs
                for run in paragraph.runs:
                    if run.bold or run.italic or run.underline:
                        has_complex_formatting = True
                        break
                if has_complex_formatting:
                    break
            
            if has_complex_formatting:
                analysis['accessibility_notes'].append("Document contains formatting that will be preserved in conversion.")
            
            return analysis
        except Exception as e:
            return {
                'error': str(e),
                'suggested_action': 'Unable to analyze document structure. File may be corrupted.'
            }


class FileSafetyManager:
    """Handles file safety operations: hashing, collision detection, backups."""
    
    def __init__(self, safety_config: SafetyConfig = None):
        self.config = safety_config or SafetyConfig()
        self.validator = FileValidator()
    
    def calculate_file_hash(self, file_path: Path) -> str:
        """Calculate SHA256 hash of file."""
        if not file_path.exists():
            return ""
        
        sha256_hash = hashlib.sha256()
        with open(file_path, "rb") as f:
            for byte_block in iter(lambda: f.read(4096), b""):
                sha256_hash.update(byte_block)
        return sha256_hash.hexdigest()
    
    def detect_conversion_collision(self, source_file: Path, target_file: Path) -> bool:
        """Check if target file would create a conversion collision."""
        if not target_file.exists():
            return False
        
        # Check if target has same basename but different extension
        if source_file.stem == target_file.stem:
            source_hash = self.calculate_file_hash(source_file)
            target_hash = self.calculate_file_hash(target_file)
            
            # If hashes are different, it's a collision
            return source_hash != target_hash
        
        return False
    
    def create_backup(self, file_path: Path) -> Optional[Path]:
        """Create backup of existing file."""
        if not file_path.exists():
            return None
        
        backup_path = file_path.with_suffix(f"{file_path.suffix}{self.config.backup_suffix}")
        counter = 1
        
        while backup_path.exists():
            backup_path = file_path.with_suffix(f"{file_path.suffix}{self.config.backup_suffix}.{counter}")
            counter += 1
        
        try:
            shutil.copy2(file_path, backup_path)
            return backup_path
        except Exception as e:
            print(f"⚠️  Warning: Could not create backup: {e}")
            return None
    
    def confirm_overwrite(self, file_path: Path) -> bool:
        """Get user confirmation for file overwrite."""
        if not self.config.require_confirmation:
            return True
        
        response = input(f"⚠️  File '{file_path}' exists. Overwrite? [y/N]: ").lower().strip()
        return response in ['y', 'yes']
    
    def safe_document_processing(self, file_path: str) -> Dict[str, Any]:
        """Process document with comprehensive error handling"""
        try:
            # Pre-flight checks
            format_check = self.validator.detect_file_format(file_path)
            if not format_check['supported']:
                raise UnsupportedFormatError(
                    format_check['message'],
                    suggested_action=format_check['suggested_action'],
                    error_code='UNSUPPORTED_FORMAT'
                )
            
            size_check = self.validator.validate_file_size(file_path)
            if not size_check['valid']:
                raise FileSizeError(
                    size_check['message'],
                    suggested_action=size_check['suggested_action'],
                    error_code='FILE_TOO_LARGE'
                )
            
            protection_check = self.validator.check_file_protection(file_path)
            if protection_check.get('protected', False):
                raise ProtectedFileError(
                    protection_check['message'],
                    suggested_action=protection_check['suggested_action'],
                    error_code='PASSWORD_PROTECTED'
                )
            
            # Analyze document structure
            structure_analysis = self.validator.analyze_document_structure(file_path)
            
            return {
                'status': 'ready',
                'checks_passed': True,
                'file_info': {
                    'size_mb': size_check.get('size_mb'),
                    'format': format_check['format']
                },
                'structure_analysis': structure_analysis
            }
            
        except WordParserError as e:
            return {
                'status': 'error',
                'error_code': e.error_code,
                'message': e.message,
                'suggested_action': e.suggested_action,
                'user_friendly': True
            }
        except Exception as e:
            return {
                'status': 'error',
                'error_code': 'UNKNOWN_ERROR',
                'message': f"Unexpected error: {str(e)}",
                'suggested_action': 'Please check file integrity and try again. If problem persists, the file may be corrupted.',
                'user_friendly': False
            }
    
    def safe_write_check(self, source_file: Path, target_file: Path) -> Tuple[bool, str]:
        """
        Comprehensive safety check before writing.
        Returns (can_proceed, reason)
        """
        # Check for collision
        if self.detect_conversion_collision(source_file, target_file):
            return False, f"Collision detected: {target_file} exists with different content"
        
        # Check if target exists and get confirmation
        if target_file.exists():
            if self.config.prevent_overwrite:
                if not self.confirm_overwrite(target_file):
                    return False, "User cancelled overwrite"
            
            # Create backup if requested
            if self.config.create_backup:
                backup_path = self.create_backup(target_file)
                if backup_path:
                    print(f"✅ Backup created: {backup_path}")
        
        return True, "Safe to proceed"


class WordToMarkdownConverter:
    """Converts Word documents to Markdown with metadata preservation and safety."""
    
    def __init__(self, safety_manager: FileSafetyManager = None):
        self.metadata = FormatMetadata()
        self.current_line = 0
        self.safety = safety_manager or FileSafetyManager()
    
    def convert_docx_to_md(self, docx_path: str, md_path: str) -> bool:
        """Convert a Word document to Markdown with safety checks."""
        source_file = Path(docx_path)
        target_file = Path(md_path)
        
        if not source_file.exists():
            print(f"❌ Error: Source file {source_file} does not exist")
            return False
        
        # Safety check
        can_proceed, reason = self.safety.safe_write_check(source_file, target_file)
        if not can_proceed:
            print(f"❌ Safety check failed: {reason}")
            return False
        
        try:
            doc = Document(docx_path)
            markdown_content = self._extract_content_and_metadata(doc)
            
            # Add file hash to metadata
            self.metadata.file_hash = self.safety.calculate_file_hash(source_file)
            from datetime import datetime
            self.metadata.conversion_timestamp = datetime.now().isoformat()
            
            # Add metadata footer
            metadata_json = json.dumps(asdict(self.metadata), indent=2)
            full_content = f"{markdown_content}\n\n<!-- WORD_CONVERSION_METADATA\n{metadata_json}\n-->\n"
            
            with open(target_file, 'w', encoding='utf-8') as f:
                f.write(full_content)
            
            print(f"✅ Successfully converted {source_file} → {target_file}")
            print(f"📊 File hash: {self.metadata.file_hash[:16]}...")
            return True
            
        except Exception as e:
            print(f"❌ Error converting {source_file}: {e}")
            return False
    
    def _extract_content_and_metadata(self, doc: Document) -> str:
        """Extract content and build metadata."""
        content_lines = []
        self.current_line = 0
        
        for paragraph in doc.paragraphs:
            line_content = self._process_paragraph(paragraph)
            if line_content.strip():
                content_lines.append(line_content)
                self.current_line += 1
        
        # Process tables
        for table in doc.tables:
            table_md = self._process_table(table)
            if table_md:
                content_lines.append(table_md)
                self.current_line += table_md.count('\n')
        
        return '\n'.join(content_lines)
    
    def _process_paragraph(self, paragraph) -> str:
        """Process a paragraph and extract formatting."""
        if not paragraph.text.strip():
            return ""
        
        # Check for heading
        if paragraph.style.name.startswith('Heading') or paragraph.style.name.startswith('Title'):
            if paragraph.style.name == 'Title':
                level = 1
            else:
                level = int(re.findall(r'\d+', paragraph.style.name)[0]) if re.findall(r'\d+', paragraph.style.name) else 1
            
            self.metadata.heading_levels[self.current_line] = level
            return f"{'#' * level} {paragraph.text}"
        
        # Check for list items
        if any(list_style in paragraph.style.name for list_style in ['List', 'Bullet', 'Number']):
            list_info = {
                'line': self.current_line,
                'style': paragraph.style.name
                # Note: text content is already in markdown, no need to store again
            }
            self.metadata.lists.append(list_info)
            
            if any(bullet in paragraph.style.name for bullet in ['Bullet', 'bullet']):
                return f"- {paragraph.text}"
            else:
                return f"1. {paragraph.text}"
        
        # Process runs for inline formatting
        formatted_text = self._process_runs(paragraph.runs)
        
        # Check alignment and other paragraph properties
        if paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER:
            self.metadata.styles[str(self.current_line)] = "center"
        elif paragraph.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
            self.metadata.styles[str(self.current_line)] = "right"
        
        return formatted_text
    
    def _process_runs(self, runs) -> str:
        """Process runs within a paragraph for inline formatting."""
        result = ""
        char_position = 0
        
        for run in runs:
            start_pos = char_position
            end_pos = char_position + len(run.text)
            
            text = run.text
            
            # Handle formatting - order matters for nested formatting
            formatting_applied = False
            
            # Handle superscript/subscript first
            if hasattr(run.font, 'superscript') and run.font.superscript:
                text = f"<sup>{text}</sup>"
                formatting_applied = True
            elif hasattr(run.font, 'subscript') and run.font.subscript:
                text = f"<sub>{text}</sub>"
                formatting_applied = True
            
            # Handle strikethrough
            if hasattr(run.font, 'strike') and run.font.strike:
                text = f"~~{text}~~"
                formatting_applied = True
            
            # Handle underline
            if run.underline:
                text = f"<u>{text}</u>"
                formatting_applied = True
            
            # Handle bold/italic combinations
            if run.bold and run.italic:
                text = f"***{text}***"
            elif run.bold:
                text = f"**{text}**"
            elif run.italic:
                text = f"*{text}*"
            
            # Handle hyperlinks
            if hasattr(run.element, 'xpath') and run.element.xpath('.//w:hyperlink'):
                hyperlink_info = {
                    'start': start_pos,
                    'end': end_pos,
                    'line': self.current_line
                    # Note: text content is preserved in markdown syntax
                }
                self.metadata.hyperlinks.append(hyperlink_info)
                text = f"[{run.text}](#{run.text.replace(' ', '-').lower()})"
            
            result += text
            char_position = end_pos
        
        return result
    
    def _process_table(self, table) -> str:
        """Convert Word table to Markdown table."""
        if not table.rows:
            return ""
        
        table_data = []
        for row in table.rows:
            row_data = [cell.text.strip() for cell in row.cells]
            table_data.append(row_data)
        
        if not table_data or not table_data[0]:
            return ""
        
        # Store table metadata
        table_info = {
            'line': self.current_line,
            'rows': len(table_data),
            'cols': len(table_data[0]),
            'data': table_data
        }
        self.metadata.tables.append(table_info)
        
        # Convert to Markdown table
        markdown_table = []
        markdown_table.append("| " + " | ".join(table_data[0]) + " |")
        markdown_table.append("| " + " | ".join(["---"] * len(table_data[0])) + " |")
        
        for row in table_data[1:]:
            markdown_table.append("| " + " | ".join(row) + " |")
        
        return "\n".join(markdown_table)


class MarkdownToWordConverter:
    """Converts Markdown to Word documents with configurable formatting and safety."""
    
    def __init__(self, config: ConversionConfig = None, template_path: str = None, safety_manager: FileSafetyManager = None):
        self.config = config or ConversionConfig()
        self.template_path = template_path
        self.safety = safety_manager or FileSafetyManager()
    
    def convert_md_to_docx(self, md_path: str, docx_path: str) -> bool:
        """Convert Markdown to Word document with safety checks."""
        source_file = Path(md_path)
        target_file = Path(docx_path)
        
        if not source_file.exists():
            print(f"❌ Error: Source file {source_file} does not exist")
            return False
        
        # Safety check
        can_proceed, reason = self.safety.safe_write_check(source_file, target_file)
        if not can_proceed:
            print(f"❌ Safety check failed: {reason}")
            return False
        
        try:
            with open(md_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Check for YAML frontmatter config
            frontmatter_config = self._extract_frontmatter_config(content)
            if frontmatter_config:
                self._update_config_from_dict(frontmatter_config)
                content = self._strip_frontmatter(content)
            
            # Extract metadata if present
            metadata = self._extract_metadata(content)
            markdown_content = self._strip_metadata(content)
            
            # Create Word document
            if self.template_path and Path(self.template_path).exists():
                doc = Document(self.template_path)
            else:
                doc = Document()
                self._setup_default_styles(doc)
            
            self._build_document(doc, markdown_content, metadata)
            
            doc.save(docx_path)
            
            # Calculate and display hash
            output_hash = self.safety.calculate_file_hash(target_file)
            print(f"✅ Successfully converted {source_file} → {target_file}")
            print(f"📊 Output hash: {output_hash[:16]}...")
            return True
            
        except Exception as e:
            print(f"❌ Error converting {source_file}: {e}")
            return False
    
    def _extract_frontmatter_config(self, content: str) -> Optional[Dict]:
        """Extract YAML frontmatter configuration."""
        if not yaml or not content.startswith('---\n'):
            return None
        
        try:
            end_marker = content.find('\n---\n', 4)
            if end_marker == -1:
                return None
            
            yaml_content = content[4:end_marker]
            return yaml.safe_load(yaml_content)
        except:
            return None
    
    def _strip_frontmatter(self, content: str) -> str:
        """Remove YAML frontmatter from content."""
        if not content.startswith('---\n'):
            return content
        
        end_marker = content.find('\n---\n', 4)
        if end_marker == -1:
            return content
        
        return content[end_marker + 5:]
    
    def _update_config_from_dict(self, config_dict: Dict) -> None:
        """Update configuration from dictionary."""
        for key, value in config_dict.items():
            if hasattr(self.config, key):
                setattr(self.config, key, value)
    
    def _extract_metadata(self, content: str) -> Optional[FormatMetadata]:
        """Extract metadata from Markdown content."""
        metadata_match = re.search(r'<!-- WORD_CONVERSION_METADATA\n(.*?)\n-->', content, re.DOTALL)
        if metadata_match:
            try:
                metadata_dict = json.loads(metadata_match.group(1))
                return FormatMetadata(**metadata_dict)
            except json.JSONDecodeError:
                pass
        return FormatMetadata()
    
    def _strip_metadata(self, content: str) -> str:
        """Remove metadata from content."""
        return re.sub(r'\n\n<!-- WORD_CONVERSION_METADATA.*?-->\n?$', '', content, flags=re.DOTALL)
    
    def _setup_default_styles(self, doc: Document) -> None:
        """Set up default document styles."""
        # Update Normal style
        normal_style = doc.styles['Normal']
        normal_font = normal_style.font
        normal_font.name = self.config.font_name
        normal_font.size = Pt(self.config.font_size)
        
        normal_paragraph = normal_style.paragraph_format
        normal_paragraph.space_after = Pt(self.config.paragraph_spacing_after)
        normal_paragraph.line_spacing = self.config.line_spacing
        normal_paragraph.first_line_indent = Inches(self.config.paragraph_first_line_indent)
        
        # Update heading styles
        for level in range(1, 7):
            try:
                heading_style = doc.styles[f'Heading {level}']
                self._configure_heading_style(heading_style, level)
            except KeyError:
                # Create heading style if it doesn't exist
                heading_style = doc.styles.add_style(f'Heading {level}', WD_STYLE_TYPE.PARAGRAPH)
                self._configure_heading_style(heading_style, level)
    
    def _configure_heading_style(self, style, level: int) -> None:
        """Configure a heading style."""
        font = style.font
        font.name = self.config.heading_font
        font.size = Pt(self.config.heading_sizes.get(level, 12))
        font.bold = True
        
        # Set color if specified
        if level in self.config.heading_colors:
            color_hex = self.config.heading_colors[level].replace('#', '')
            try:
                r, g, b = int(color_hex[:2], 16), int(color_hex[2:4], 16), int(color_hex[4:6], 16)
                font.color.rgb = RGBColor(r, g, b)
            except ValueError:
                pass
        
        # Set spacing
        paragraph = style.paragraph_format
        paragraph.space_before = Pt(self.config.heading_spacing_before.get(level, 6))
        paragraph.space_after = Pt(self.config.heading_spacing_after.get(level, 3))
        paragraph.keep_with_next = True
    
    def _build_document(self, doc: Document, content: str, metadata: FormatMetadata) -> None:
        """Build Word document from Markdown content and metadata."""
        # Parse markdown with extensions
        md = markdown.Markdown(extensions=['fenced_code', 'tables', 'toc'])
        
        lines = content.split('\n')
        current_list_type = None
        in_code_block = False
        code_block_content = []
        
        for line_num, line in enumerate(lines):
            line = line.rstrip()
            
            # Handle code blocks
            if line.startswith('```'):
                if in_code_block:
                    # End of code block
                    self._add_code_block(doc, '\n'.join(code_block_content))
                    code_block_content = []
                    in_code_block = False
                else:
                    # Start of code block
                    in_code_block = True
                continue
            
            if in_code_block:
                code_block_content.append(line)
                continue
            
            if not line.strip():
                if not in_code_block:
                    # Add empty paragraph for spacing
                    doc.add_paragraph()
                continue
            
            # Handle horizontal rules and dividers
            if self._is_horizontal_rule(line):
                self._add_horizontal_rule(doc)
                continue
            
            # Handle header boxes (equals dividers around text)
            if self._is_header_box_divider(line, lines, line_num):
                header_text = self._extract_header_box_text(lines, line_num)
                if header_text:
                    self._add_header_box(doc, header_text)
                    # Skip the header text line and closing divider
                    continue
            
            # Handle headings
            if line.startswith('#'):
                level = len(line) - len(line.lstrip('#'))
                heading_text = line.lstrip('#').strip()
                if self.config.use_builtin_styles:
                    doc.add_heading(heading_text, level)
                else:
                    p = doc.add_paragraph(heading_text)
                    self._apply_custom_heading_format(p, level)
            
            # Handle lists
            elif line.strip().startswith(('-', '*', '+')):
                list_text = line.strip()[1:].strip()
                p = doc.add_paragraph(list_text, style='List Bullet')
                current_list_type = 'bullet'
            
            elif re.match(r'^\s*\d+\.', line):
                list_text = re.sub(r'^\s*\d+\.\s*', '', line)
                p = doc.add_paragraph(list_text, style='List Number')
                current_list_type = 'number'
            
            # Handle tables
            elif '|' in line and line.strip().startswith('|'):
                table_lines = [line]
                # Collect all table lines
                for next_line_num in range(line_num + 1, len(lines)):
                    next_line = lines[next_line_num]
                    if '|' in next_line and next_line.strip().startswith('|'):
                        table_lines.append(next_line)
                    else:
                        break
                
                if len(table_lines) >= 2:  # Header + separator minimum
                    self._add_markdown_table(doc, table_lines)
                    # Skip the lines we just processed
                    line_num += len(table_lines) - 1
            
            # Handle blockquotes
            elif line.strip().startswith('>'):
                quote_text = line.strip()[1:].strip()
                p = doc.add_paragraph(quote_text)
                p.style = 'Quote'
            
            # Regular paragraphs
            else:
                p = doc.add_paragraph()
                self._apply_inline_formatting(p, line)
                current_list_type = None
    
    def _add_code_block(self, doc: Document, code_content: str) -> None:
        """Add a code block to the document."""
        p = doc.add_paragraph(code_content)
        
        # Style the code block
        font = p.runs[0].font if p.runs else p.style.font
        font.name = self.config.code_font
        font.size = Pt(self.config.code_size)
        
        # Add background color (limited support in python-docx)
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.right_indent = Inches(0.25)
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(6)
    
    def _apply_custom_heading_format(self, paragraph, level: int) -> None:
        """Apply custom heading formatting."""
        run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
        
        font = run.font
        font.name = self.config.heading_font
        font.size = Pt(self.config.heading_sizes.get(level, 12))
        font.bold = True
        
        if level in self.config.heading_colors:
            color_hex = self.config.heading_colors[level].replace('#', '')
            try:
                r, g, b = int(color_hex[:2], 16), int(color_hex[2:4], 16), int(color_hex[4:6], 16)
                font.color.rgb = RGBColor(r, g, b)
            except ValueError:
                pass
        
        paragraph.paragraph_format.space_before = Pt(self.config.heading_spacing_before.get(level, 6))
        paragraph.paragraph_format.space_after = Pt(self.config.heading_spacing_after.get(level, 3))
    
    def _apply_inline_formatting(self, paragraph, text: str) -> None:
        """Apply inline formatting to paragraph text."""
        # Parse inline markdown
        current_pos = 0
        
        # Combined pattern for bold, italic, code
        pattern = r'(\*\*\*(.+?)\*\*\*|\*\*(.+?)\*\*|\*(.+?)\*|`(.+?)`|\[(.+?)\]\((.+?)\))'
        
        for match in re.finditer(pattern, text):
            # Add text before formatting
            if match.start() > current_pos:
                paragraph.add_run(text[current_pos:match.start()])
            
            # Determine formatting type and add formatted run
            full_match = match.group(1)
            if full_match.startswith('***') and full_match.endswith('***'):
                # Bold + italic
                run = paragraph.add_run(match.group(2))
                run.bold = True
                run.italic = True
            elif full_match.startswith('**') and full_match.endswith('**'):
                # Bold
                run = paragraph.add_run(match.group(3))
                run.bold = True
            elif full_match.startswith('*') and full_match.endswith('*'):
                # Italic
                run = paragraph.add_run(match.group(4))
                run.italic = True
            elif full_match.startswith('`') and full_match.endswith('`'):
                # Inline code
                run = paragraph.add_run(match.group(5))
                run.font.name = self.config.code_font
                run.font.size = Pt(self.config.code_size)
            elif '[' in full_match and '](' in full_match:
                # Hyperlink
                run = paragraph.add_run(match.group(6))
                # Note: Creating actual hyperlinks in python-docx is complex
                run.font.color.rgb = RGBColor(0, 0, 255)
                run.underline = True
            
            current_pos = match.end()
        
        # Add remaining text
        if current_pos < len(text):
            paragraph.add_run(text[current_pos:])
    
    def _add_markdown_table(self, doc: Document, table_lines: List[str]) -> None:
        """Add a markdown table to the document."""
        # Parse table data
        rows = []
        for line in table_lines:
            if '---' in line:  # Skip separator line
                continue
            cells = [cell.strip() for cell in line.split('|')[1:-1]]  # Remove empty first/last
            if cells:
                rows.append(cells)
        
        if not rows:
            return
        
        # Create Word table
        table = doc.add_table(rows=len(rows), cols=len(rows[0]))
        table.style = self.config.table_style
        
        if self.config.table_autofit:
            table.autofit = True
        
        # Populate table
        for row_idx, row_data in enumerate(rows):
            for col_idx, cell_data in enumerate(row_data):
                if col_idx < len(table.rows[row_idx].cells):
                    cell = table.rows[row_idx].cells[col_idx]
                    # Apply inline formatting to cell text
                    cell.text = cell_data
                    if row_idx == 0:  # Header row
                        for paragraph in cell.paragraphs:
                            for run in paragraph.runs:
                                run.font.bold = True
    
    def _is_horizontal_rule(self, line: str) -> bool:
        """Check if line is a horizontal divider pattern."""
        stripped = line.strip()
        if not stripped:
            return False
        
        # Standard markdown horizontal rules
        if stripped in ['---', '***', '___']:
            return True
        
        # Long divider lines (10+ characters of same symbol)
        if len(stripped) >= 10:
            # Check for repeated dashes, equals, or unicode box drawing
            if all(c == '-' for c in stripped):  # ----------------
                return True
            if all(c == '=' for c in stripped):  # ================
                return True
            if all(c == '─' for c in stripped):  # ────────────────
                return True
        
        return False
    
    def _add_horizontal_rule(self, doc: Document) -> None:
        """Add a horizontal rule to the Word document."""
        from docx.shared import Pt, RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        
        # Create paragraph with underline to simulate horizontal line
        p = doc.add_paragraph("_" * 50)  # Create underline characters
        
        # Style as a horizontal line with minimal spacing
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(3)  # Reduced from 12
        p.paragraph_format.space_before = Pt(3)  # Reduced from 6
        
        # Style the text to look like a line
        if p.runs:
            run = p.runs[0]
            run.font.color.rgb = RGBColor(128, 128, 128)  # Gray color
            run.font.size = Pt(8)
    
    def _is_header_box_divider(self, line: str, lines: List[str], line_num: int) -> bool:
        """Check if line is start of a header box pattern (equals dividers)."""
        stripped = line.strip()
        
        # Must be long line of equals signs
        if len(stripped) >= 20 and all(c == '=' for c in stripped):
            # Check if there's a text line and closing divider following
            if line_num + 2 < len(lines):
                text_line = lines[line_num + 1].strip()
                closing_line = lines[line_num + 2].strip()
                
                # Text line should not be empty and should not be another divider
                if text_line and not all(c in '=-_' for c in text_line if c.strip()):
                    # Closing line should also be equals divider
                    if len(closing_line) >= 20 and all(c == '=' for c in closing_line):
                        return True
        
        return False
    
    def _extract_header_box_text(self, lines: List[str], line_num: int) -> Optional[str]:
        """Extract text from header box pattern and mark lines as processed."""
        if line_num + 2 < len(lines):
            header_text = lines[line_num + 1].strip()
            
            # Mark the header text line and closing divider as processed
            # by setting them to empty (they'll be skipped in the main loop)
            lines[line_num + 1] = ""
            lines[line_num + 2] = ""
            
            return header_text
        return None
    
    def _add_header_box(self, doc: Document, header_text: str) -> None:
        """Add a bordered header box to the Word document."""
        from docx.shared import Pt, RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        
        # Create the header paragraph with minimal spacing
        p = doc.add_paragraph(header_text)
        
        # Style the header box
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Add top and bottom borders
        p.paragraph_format.border_top.color.rgb = RGBColor(0, 0, 0)  # Black
        p.paragraph_format.border_top.width = Pt(2)
        p.paragraph_format.border_bottom.color.rgb = RGBColor(0, 0, 0)  # Black  
        p.paragraph_format.border_bottom.width = Pt(2)
        
        # Minimal padding
        p.paragraph_format.space_before = Pt(4)  # Reduced from 12
        p.paragraph_format.space_after = Pt(4)   # Reduced from 12
        
        # Style the text
        if p.runs:
            run = p.runs[0]
            run.font.bold = True
            run.font.size = Pt(14)
        else:
            # If no runs, add the text as a run
            run = p.add_run(header_text)
            run.font.bold = True
            run.font.size = Pt(14)


def load_config_file(config_path: str) -> ConversionConfig:
    """Load configuration from file."""
    if not Path(config_path).exists():
        print(f"⚠️  Config file {config_path} not found, using defaults")
        return ConversionConfig()
    
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            if config_path.endswith('.json'):
                config_dict = json.load(f)
            elif config_path.endswith(('.yml', '.yaml')) and yaml:
                config_dict = yaml.safe_load(f)
            else:
                print("⚠️  Unsupported config format, using defaults")
                return ConversionConfig()
        
        # Create config object
        config = ConversionConfig()
        for key, value in config_dict.items():
            if hasattr(config, key):
                setattr(config, key, value)
        
        return config
    except Exception as e:
        print(f"⚠️  Error loading config: {e}, using defaults")
        return ConversionConfig()


def create_sample_config(config_path: str) -> None:
    """Create a sample configuration file."""
    config = ConversionConfig()
    config_dict = asdict(config)
    
    try:
        if config_path.endswith('.json'):
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config_dict, f, indent=2)
        elif config_path.endswith(('.yml', '.yaml')) and yaml:
            with open(config_path, 'w', encoding='utf-8') as f:
                yaml.dump(config_dict, f, default_flow_style=False)
        else:
            print("❌ Unsupported config format")
            return
        
        print(f"✅ Sample configuration created at {config_path}")
    except Exception as e:
        print(f"❌ Error creating config: {e}")


@click.group()
@click.option('--verbose', '-v', is_flag=True, help='Enable verbose output')
@click.option('--quiet', '-q', is_flag=True, help='Minimal output for automation')
@click.option('--config', type=click.Path(exists=True), help='Configuration file path (JSON or YAML)')
@click.pass_context
def cli(ctx, verbose, quiet, config):
    """FSS Parse Word - Professional document processing toolkit
    
    Safe, hash-validated bidirectional conversion between .docx and .md formats.
    
    Examples:
      fss-parse-word convert document.docx output.md
      fss-parse-word info document.docx --json
      fss-parse-word extract document.docx --sections headers
    """
    ctx.ensure_object(dict)
    ctx.obj['verbose'] = verbose
    ctx.obj['quiet'] = quiet
    ctx.obj['config'] = config
    
    # Store global settings for output control
    # Note: Rich Console doesn't have quiet/verbose attributes by default
    # We'll handle this in individual commands


@cli.command()
@click.argument('input_file', type=click.Path(exists=True))
@click.argument('output_file', type=click.Path())
@click.option('--force', is_flag=True, help='Skip confirmation prompts')
@click.option('--backup/--no-backup', default=False, help='Create backup (not needed for conversions)')
@click.option('--format', type=click.Choice(['markdown', 'txt', 'docx']), default='markdown', help='Output format')
@click.option('--template', type=click.Path(exists=True), help='Word template for DOCX output')
@click.option('--strip-metadata', is_flag=True, help='Remove document metadata for privacy')
@click.option('--preserve-metadata', is_flag=True, default=True, help='Keep document metadata (default)')
@click.option('--json', 'output_json', is_flag=True, help='JSON output for automation')
@click.pass_context
def convert(ctx, input_file, output_file, force, backup, format, template, strip_metadata, preserve_metadata, output_json):
    """Convert between Word and Markdown formats
    
    Auto-detects conversion direction based on file extensions.
    Conversions leave original files untouched by default.
    """
    input_path = Path(input_file)
    output_path = Path(output_file)
    
    # Validate metadata privacy options
    if strip_metadata and preserve_metadata and preserve_metadata is not True:
        click.echo("❌ Error: Cannot both strip and preserve metadata. Choose one option.")
        sys.exit(1)
    
    # Set metadata handling mode
    privacy_mode = strip_metadata if strip_metadata else False
    
    if privacy_mode:
        click.echo("🔒 Privacy mode: Document metadata will be removed during conversion")
    else:
        click.echo("📝 Standard mode: Document metadata will be preserved")
    
    # Configure safety settings
    safety_config = SafetyConfig(
        require_confirmation=not force,
        create_backup=backup,  # Default False for conversions
        check_hash=True,
        prevent_overwrite=True
    )
    
    safety_manager = FileSafetyManager(safety_config)
    
    # Load configuration
    config = load_config_file(ctx.obj['config']) if ctx.obj['config'] else ConversionConfig()
    
    # Auto-detect conversion direction
    if input_path.suffix.lower() == '.docx':
        direction = 'docx2md'
    elif input_path.suffix.lower() == '.md':
        direction = 'md2docx'
    else:
        click.echo("❌ Error: Cannot auto-detect conversion direction from file extensions")
        sys.exit(1)
    
    # Enhanced validation with accessibility-focused error handling
    validation_result = safety_manager.safe_document_processing(str(input_path))
    
    if validation_result['status'] == 'error':
        if output_json:
            result = {
                "operation": "convert",
                "status": "error",
                "error_code": validation_result.get('error_code'),
                "message": validation_result['message'],
                "suggested_action": validation_result['suggested_action'],
                "timestamp": datetime.now().isoformat()
            }
            click.echo(json.dumps(result, indent=2))
        else:
            click.echo(f"❌ Error: {validation_result['message']}")
            if validation_result.get('suggested_action'):
                click.echo(f"💡 Suggestion: {validation_result['suggested_action']}")
        sys.exit(1)
    
    # Display structure analysis for user awareness
    if validation_result.get('structure_analysis'):
        analysis = validation_result['structure_analysis']
        if analysis.get('accessibility_notes'):
            for note in analysis['accessibility_notes']:
                click.echo(f"ℹ️  {note}")
    
    start_time = datetime.now()
    
    try:
        if direction == 'docx2md':
            converter = WordToMarkdownConverter(safety_manager)
            success = converter.convert_docx_to_md(str(input_path), str(output_path))
        else:
            converter = MarkdownToWordConverter(config, template, safety_manager)
            success = converter.convert_md_to_docx(str(input_path), str(output_path))
        
        processing_time = (datetime.now() - start_time).total_seconds()
        
        if output_json:
            result = {
                "operation": "convert",
                "input": {
                    "filename": str(input_path),
                    "size_bytes": input_path.stat().st_size,
                    "format": input_path.suffix[1:]
                },
                "output": {
                    "filename": str(output_path),
                    "size_bytes": output_path.stat().st_size if output_path.exists() else 0,
                    "format": output_path.suffix[1:]
                },
                "status": "success" if success else "error",
                "processing_time": f"{processing_time:.2f}s",
                "timestamp": datetime.now().isoformat()
            }
            click.echo(json.dumps(result, indent=2))
        elif success and not ctx.obj['quiet']:
            click.echo(f"✅ Conversion completed: {input_path.name} → {output_path.name}")
        
        sys.exit(0 if success else 1)
            
    except Exception as e:
        if output_json:
            error_result = {
                "operation": "convert",
                "status": "error",
                "error": str(e),
                "timestamp": datetime.now().isoformat()
            }
            click.echo(json.dumps(error_result, indent=2))
        else:
            click.echo(f"❌ Conversion failed: {e}")
        sys.exit(1)


@cli.command()
@click.argument('input_file', type=click.Path(exists=True))
@click.option('--json', 'output_json', is_flag=True, help='JSON output for automation')
@click.pass_context
def info(ctx, input_file, output_json):
    """Display document information and metadata
    
    Shows document properties, word count, structure analysis, and file metadata.
    """
    input_path = Path(input_file)
    
    try:
        if input_path.suffix.lower() == '.docx':
            doc_info = get_docx_info(input_path)
        elif input_path.suffix.lower() == '.md':
            doc_info = get_markdown_info(input_path)
        else:
            click.echo("❌ Error: Unsupported file format. Only .docx and .md files supported.")
            sys.exit(1)
        
        if output_json:
            click.echo(json.dumps(doc_info, indent=2, default=str))
        else:
            display_info_table(doc_info, quiet=ctx.obj['quiet'])
            
    except Exception as e:
        if output_json:
            error_result = {
                "operation": "info",
                "status": "error",
                "error": str(e),
                "timestamp": datetime.now().isoformat()
            }
            click.echo(json.dumps(error_result, indent=2))
        else:
            click.echo(f"❌ Error reading document: {e}")
        sys.exit(1)


@cli.command()
@click.argument('input_dir', type=click.Path(exists=True, file_okay=False))
@click.option('--output', type=click.Path(), required=True, help='Output directory')
@click.option('--pattern', default='*.{docx,md}', help='File pattern to match')
@click.option('--operation', type=click.Choice(['convert', 'info', 'extract']), default='convert', help='Operation to perform')
@click.option('--format', type=click.Choice(['markdown', 'txt', 'json']), default='markdown', help='Output format')
@click.option('--sections', type=click.Choice(['headers', 'tables', 'images', 'paragraphs', 'all']),
              default='headers', help='Sections to extract (for extract operation)')
@click.option('--threads', default=4, help='Number of processing threads')
@click.option('--json', 'output_json', is_flag=True, help='JSON output for automation')
@click.pass_context
def batch(ctx, input_dir, output, pattern, operation, format, sections, threads, output_json):
    """Process multiple documents in directory
    
    Processes all matching documents in a directory with the specified operation.
    Supports convert, info, and extract operations on multiple files.
    """
    import glob
    import concurrent.futures
    from pathlib import Path
    
    input_path = Path(input_dir)
    output_path = Path(output)
    output_path.mkdir(parents=True, exist_ok=True)
    
    # Find matching files
    if '{' in pattern and '}' in pattern:
        # Handle pattern like "*.{docx,md}"
        pattern_parts = pattern.replace('{', '').replace('}', '').split(',')
        all_files = []
        for part in pattern_parts:
            part = part.strip()
            if not part.startswith('*.'):
                part = '*.' + part
            all_files.extend(glob.glob(str(input_path / part)))
    else:
        # Handle simple pattern like "*.docx"
        all_files = glob.glob(str(input_path / pattern))
    
    if not all_files:
        click.echo(f"❌ No files found matching pattern {pattern} in {input_dir}")
        sys.exit(1)
    
    click.echo(f"📁 Found {len(all_files)} files to process")
    
    def process_file(file_path):
        file_path = Path(file_path)
        
        try:
            if operation == 'convert':
                if file_path.suffix.lower() == '.docx':
                    output_file = output_path / f"{file_path.stem}.md"
                    converter = WordToMarkdownConverter()
                    success = converter.convert_docx_to_md(str(file_path), str(output_file))
                elif file_path.suffix.lower() == '.md':
                    output_file = output_path / f"{file_path.stem}.docx"
                    converter = MarkdownToWordConverter()
                    success = converter.convert_md_to_docx(str(file_path), str(output_file))
                else:
                    return {'file': str(file_path), 'status': 'error', 'error': 'Unsupported format'}
                
                return {'file': str(file_path), 'status': 'success' if success else 'error', 'output': str(output_file)}
            
            elif operation == 'info':
                if file_path.suffix.lower() == '.docx':
                    doc_info = get_docx_info(file_path)
                elif file_path.suffix.lower() == '.md':
                    doc_info = get_markdown_info(file_path)
                else:
                    return {'file': str(file_path), 'status': 'error', 'error': 'Unsupported format'}
                
                # Save info to JSON file
                info_file = output_path / f"{file_path.stem}_info.json"
                with open(info_file, 'w', encoding='utf-8') as f:
                    json.dump(doc_info, f, indent=2, default=str)
                
                return {'file': str(file_path), 'status': 'success', 'output': str(info_file)}
            
            elif operation == 'extract':
                if file_path.suffix.lower() == '.docx':
                    extracted_data = extract_docx_sections(file_path, sections)
                elif file_path.suffix.lower() == '.md':
                    extracted_data = extract_markdown_sections(file_path, sections)
                else:
                    return {'file': str(file_path), 'status': 'error', 'error': 'Unsupported format'}
                
                # Save extracted data
                if format == 'json':
                    extract_file = output_path / f"{file_path.stem}_{sections}.json"
                    with open(extract_file, 'w', encoding='utf-8') as f:
                        json.dump(extracted_data, f, indent=2, default=str)
                else:
                    extract_file = output_path / f"{file_path.stem}_{sections}.md"
                    formatted_output = format_extracted_markdown(extracted_data, sections, file_path)
                    with open(extract_file, 'w', encoding='utf-8') as f:
                        f.write(formatted_output)
                
                return {'file': str(file_path), 'status': 'success', 'output': str(extract_file)}
                
        except Exception as e:
            return {'file': str(file_path), 'status': 'error', 'error': str(e)}
    
    # Process files with threading
    results = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=threads) as executor:
        future_to_file = {executor.submit(process_file, f): f for f in all_files}
        for future in concurrent.futures.as_completed(future_to_file):
            result = future.result()
            results.append(result)
            
            if not output_json:
                if result['status'] == 'success':
                    click.echo(f"✅ {Path(result['file']).name} → {Path(result['output']).name}")
                else:
                    click.echo(f"❌ {Path(result['file']).name}: {result.get('error', 'Unknown error')}")
    
    # Summary
    successful = len([r for r in results if r['status'] == 'success'])
    failed = len([r for r in results if r['status'] == 'error'])
    
    if output_json:
        summary = {
            'operation': 'batch',
            'total_files': len(all_files),
            'successful': successful,
            'failed': failed,
            'results': results,
            'timestamp': datetime.now().isoformat()
        }
        click.echo(json.dumps(summary, indent=2))
    else:
        click.echo(f"\n📊 Batch processing complete: {successful} successful, {failed} failed")


@cli.command()
@click.argument('input_file', type=click.Path(exists=True))
@click.option('--sections', type=click.Choice(['headers', 'tables', 'images', 'paragraphs', 'all']),
              default='headers', help='Content sections to extract')
@click.option('--output', type=click.Path(), help='Output file path')
@click.option('--format', type=click.Choice(['markdown', 'txt', 'json']), default='markdown', help='Output format')
@click.option('--json', 'output_json', is_flag=True, help='JSON output for automation')
@click.pass_context
def extract(ctx, input_file, sections, output, format, output_json):
    """Extract specific sections from document
    
    Extracts headers, tables, images, or paragraphs from Word documents.
    Outputs to file or stdout in specified format.
    """
    input_path = Path(input_file)
    
    try:
        if input_path.suffix.lower() == '.docx':
            extracted_data = extract_docx_sections(input_path, sections)
        elif input_path.suffix.lower() == '.md':
            extracted_data = extract_markdown_sections(input_path, sections)
        else:
            click.echo("❌ Error: Unsupported file format. Only .docx and .md files supported.")
            sys.exit(1)
        
        # Format output
        if format == 'json' or output_json:
            result = {
                "operation": "extract",
                "input": str(input_path),
                "sections": sections,
                "data": extracted_data,
                "timestamp": datetime.now().isoformat()
            }
            formatted_output = json.dumps(result, indent=2, default=str)
        elif format == 'markdown':
            formatted_output = format_extracted_markdown(extracted_data, sections, input_path)
        else:  # txt format
            formatted_output = format_extracted_text(extracted_data, sections)
        
        # Output to file or stdout
        if output:
            with open(output, 'w', encoding='utf-8') as f:
                f.write(formatted_output)
            if not ctx.obj['quiet']:
                click.echo(f"✅ Extracted {sections} from {input_path} → {output}")
        else:
            click.echo(formatted_output)
            
    except Exception as e:
        if output_json:
            error_result = {
                "operation": "extract",
                "status": "error",
                "error": str(e),
                "timestamp": datetime.now().isoformat()
            }
            click.echo(json.dumps(error_result, indent=2))
        else:
            click.echo(f"❌ Error extracting sections: {e}")
        sys.exit(1)


def get_docx_info(file_path: Path) -> Dict[str, Any]:
    """Extract comprehensive information from Word document"""
    doc = Document(file_path)
    
    # Basic file information
    file_stats = file_path.stat()
    
    # Document properties
    props = doc.core_properties
    
    # Count elements
    paragraphs = len(doc.paragraphs)
    tables = len(doc.tables)
    
    # Count words (approximate)
    total_text = ' '.join([p.text for p in doc.paragraphs])
    word_count = len(total_text.split())
    
    # Count headings
    headings = []
    for p in doc.paragraphs:
        if p.style.name.startswith('Heading'):
            level = int(re.findall(r'\d+', p.style.name)[0]) if re.findall(r'\d+', p.style.name) else 1
            headings.append({
                'level': level,
                'text': p.text[:50] + '...' if len(p.text) > 50 else p.text
            })
    
    return {
        'filename': file_path.name,
        'file_size_bytes': file_stats.st_size,
        'file_size_human': f"{file_stats.st_size / 1024:.1f} KB" if file_stats.st_size < 1024*1024 else f"{file_stats.st_size / (1024*1024):.1f} MB",
        'created': datetime.fromtimestamp(file_stats.st_ctime),
        'modified': datetime.fromtimestamp(file_stats.st_mtime),
        'format': 'Word Document (.docx)',
        'title': props.title or 'Untitled',
        'author': props.author or 'Unknown',
        'subject': props.subject or '',
        'paragraphs': paragraphs,
        'tables': tables,
        'headings_count': len(headings),
        'word_count': word_count,
        'headings': headings[:10]  # First 10 headings
    }


def get_markdown_info(file_path: Path) -> Dict[str, Any]:
    """Extract comprehensive information from Markdown document"""
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    file_stats = file_path.stat()
    
    # Count elements
    lines = content.split('\n')
    paragraphs = len([line for line in lines if line.strip() and not line.startswith('#')])
    
    # Count headings
    headings = []
    for line in lines:
        if line.strip().startswith('#'):
            level = len(line) - len(line.lstrip('#'))
            text = line.lstrip('#').strip()
            headings.append({
                'level': level,
                'text': text[:50] + '...' if len(text) > 50 else text
            })
    
    # Count tables
    table_count = len([line for line in lines if '|' in line and line.strip().startswith('|')])
    
    # Word count
    word_count = len(content.split())
    
    return {
        'filename': file_path.name,
        'file_size_bytes': file_stats.st_size,
        'file_size_human': f"{file_stats.st_size / 1024:.1f} KB" if file_stats.st_size < 1024*1024 else f"{file_stats.st_size / (1024*1024):.1f} MB",
        'created': datetime.fromtimestamp(file_stats.st_ctime),
        'modified': datetime.fromtimestamp(file_stats.st_mtime),
        'format': 'Markdown (.md)',
        'lines': len(lines),
        'paragraphs': paragraphs,
        'tables': table_count,
        'headings_count': len(headings),
        'word_count': word_count,
        'headings': headings[:10]  # First 10 headings
    }


def display_info_table(doc_info: Dict[str, Any], quiet: bool = False) -> None:
    """Display document information in formatted table"""
    if quiet:
        # Minimal output for quiet mode
        print(f"{doc_info['filename']}: {doc_info['word_count']} words, {doc_info['paragraphs']} paragraphs, {doc_info['headings_count']} headings")
        return
        
    table = Table(title=f"Document Information: {doc_info['filename']}")
    table.add_column("Property", style="cyan", width=15)
    table.add_column("Value", style="white")
    
    # File properties
    table.add_row("File Size", doc_info['file_size_human'])
    table.add_row("Format", doc_info['format'])
    table.add_row("Created", str(doc_info['created']))
    table.add_row("Modified", str(doc_info['modified']))
    
    # Document properties
    if 'title' in doc_info:
        table.add_row("Title", doc_info['title'])
    if 'author' in doc_info:
        table.add_row("Author", doc_info['author'])
    
    # Content statistics
    table.add_row("Word Count", str(doc_info['word_count']))
    table.add_row("Paragraphs", str(doc_info['paragraphs']))
    table.add_row("Tables", str(doc_info['tables']))
    table.add_row("Headings", str(doc_info['headings_count']))
    
    console.print(table)
    
    # Show headings if any
    if doc_info['headings']:
        console.print("\n📋 Document Structure (first 10 headings):")
        for heading in doc_info['headings']:
            indent = "  " * (heading['level'] - 1)
            console.print(f"{indent}{'#' * heading['level']} {heading['text']}")


def extract_docx_sections(file_path: Path, sections: str) -> Dict[str, Any]:
    """Extract specific sections from Word document"""
    doc = Document(file_path)
    extracted = {}
    
    if sections in ['headers', 'all']:
        headers = []
        for p in doc.paragraphs:
            if p.style.name.startswith('Heading') or p.style.name == 'Title':
                level = 1 if p.style.name == 'Title' else int(re.findall(r'\d+', p.style.name)[0])
                headers.append({
                    'level': level,
                    'text': p.text,
                    'style': p.style.name
                })
        extracted['headers'] = headers
    
    if sections in ['tables', 'all']:
        tables = []
        for table in doc.tables:
            table_data = []
            for row in table.rows:
                row_data = [cell.text.strip() for cell in row.cells]
                table_data.append(row_data)
            tables.append({
                'rows': len(table_data),
                'cols': len(table_data[0]) if table_data else 0,
                'data': table_data
            })
        extracted['tables'] = tables
    
    if sections in ['paragraphs', 'all']:
        paragraphs = []
        for p in doc.paragraphs:
            if p.text.strip() and not p.style.name.startswith('Heading'):
                paragraphs.append({
                    'text': p.text,
                    'style': p.style.name
                })
        extracted['paragraphs'] = paragraphs
    
    if sections in ['images', 'all']:
        # Image extraction is complex in python-docx, placeholder for now
        extracted['images'] = []
    
    return extracted


def extract_markdown_sections(file_path: Path, sections: str) -> Dict[str, Any]:
    """Extract specific sections from Markdown document"""
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    lines = content.split('\n')
    extracted = {}
    
    if sections in ['headers', 'all']:
        headers = []
        for line in lines:
            if line.strip().startswith('#'):
                level = len(line) - len(line.lstrip('#'))
                text = line.lstrip('#').strip()
                headers.append({
                    'level': level,
                    'text': text
                })
        extracted['headers'] = headers
    
    if sections in ['tables', 'all']:
        tables = []
        current_table = []
        in_table = False
        
        for line in lines:
            if '|' in line and line.strip().startswith('|'):
                if '---' not in line:  # Skip separator line
                    current_table.append([cell.strip() for cell in line.split('|')[1:-1]])
                in_table = True
            else:
                if in_table and current_table:
                    tables.append({
                        'rows': len(current_table),
                        'cols': len(current_table[0]) if current_table else 0,
                        'data': current_table
                    })
                    current_table = []
                in_table = False
        
        if current_table:  # Handle table at end of file
            tables.append({
                'rows': len(current_table),
                'cols': len(current_table[0]) if current_table else 0,
                'data': current_table
            })
        
        extracted['tables'] = tables
    
    if sections in ['paragraphs', 'all']:
        paragraphs = []
        for line in lines:
            if line.strip() and not line.startswith('#') and not line.startswith('|'):
                paragraphs.append({'text': line.strip()})
        extracted['paragraphs'] = paragraphs
    
    if sections in ['images', 'all']:
        images = []
        for line in lines:
            if '![' in line:
                images.append({'markdown': line.strip()})
        extracted['images'] = images
    
    return extracted


def format_extracted_markdown(data: Dict[str, Any], sections: str, input_path: Path) -> str:
    """Format extracted data as Markdown"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    md_content = f"# Word Document Analysis: {input_path.name}\n"
    md_content += f"*Generated on: {timestamp}*\n\n"
    
    md_content += f"## Metadata\n"
    md_content += f"- **File:** {input_path.name}\n"
    md_content += f"- **Sections Extracted:** {sections}\n"
    md_content += f"- **Processing Date:** {timestamp}\n\n"
    
    if 'headers' in data:
        md_content += f"## Headers ({len(data['headers'])})\n\n"
        for header in data['headers']:
            md_content += f"{'#' * header['level']} {header['text']}\n"
        md_content += "\n"
    
    if 'tables' in data:
        md_content += f"## Tables ({len(data['tables'])})\n\n"
        for i, table in enumerate(data['tables'], 1):
            md_content += f"### Table {i} ({table['rows']} rows × {table['cols']} columns)\n\n"
            if table['data']:
                # Format as markdown table
                md_content += "| " + " | ".join(table['data'][0]) + " |\n"
                md_content += "| " + " | ".join(["---"] * len(table['data'][0])) + " |\n"
                for row in table['data'][1:]:
                    md_content += "| " + " | ".join(row) + " |\n"
            md_content += "\n"
    
    if 'paragraphs' in data:
        md_content += f"## Text Content ({len(data['paragraphs'])} paragraphs)\n\n"
        for para in data['paragraphs'][:10]:  # Limit to first 10
            md_content += f"{para['text']}\n\n"
        if len(data['paragraphs']) > 10:
            md_content += f"... and {len(data['paragraphs']) - 10} more paragraphs\n\n"
    
    md_content += "---\n*Generated by FSS Parse Word v1.0.0*"
    return md_content


def format_extracted_text(data: Dict[str, Any], sections: str) -> str:
    """Format extracted data as plain text"""
    text_content = ""
    
    if 'headers' in data:
        text_content += f"HEADERS ({len(data['headers'])}):\n"
        for header in data['headers']:
            indent = "  " * (header['level'] - 1)
            text_content += f"{indent}{header['text']}\n"
        text_content += "\n"
    
    if 'tables' in data:
        text_content += f"TABLES ({len(data['tables'])}):\n"
        for i, table in enumerate(data['tables'], 1):
            text_content += f"Table {i}: {table['rows']} rows × {table['cols']} columns\n"
        text_content += "\n"
    
    if 'paragraphs' in data:
        text_content += f"PARAGRAPHS ({len(data['paragraphs'])}):\n"
        for i, para in enumerate(data['paragraphs'][:5], 1):  # First 5 only for text format
            text_content += f"{i}. {para['text'][:100]}...\n"
        text_content += "\n"
    
    return text_content


@cli.command()
@click.argument('input_file', type=click.Path(exists=True))
@click.option('--check-styles', is_flag=True, help='Check document style consistency')
@click.option('--fix-formatting', is_flag=True, help='Automatically fix common formatting issues')
@click.option('--json', 'output_json', is_flag=True, help='JSON output for automation')
@click.pass_context
def validate(ctx, input_file, check_styles, fix_formatting, output_json):
    """Validate Word document safety and integrity
    
    Checks document for formatting issues, style inconsistencies, and structural problems.
    Can automatically fix common issues when requested.
    """
    input_path = Path(input_file)
    
    try:
        if input_path.suffix.lower() == '.docx':
            validation_result = validate_docx_document(input_path, check_styles, fix_formatting)
        else:
            click.echo("❌ Error: Unsupported file format. Only .docx files supported for validation.")
            sys.exit(1)
        
        if output_json:
            click.echo(json.dumps(validation_result, indent=2, default=str))
        else:
            display_validation_results(validation_result, ctx.obj['quiet'])
            
    except Exception as e:
        if output_json:
            error_result = {
                "operation": "validate",
                "status": "error",
                "error": str(e),
                "timestamp": datetime.now().isoformat()
            }
            click.echo(json.dumps(error_result, indent=2))
        else:
            click.echo(f"❌ Error validating document: {e}")
        sys.exit(1)


def validate_docx_document(file_path: Path, check_styles: bool = True, fix_formatting: bool = False) -> Dict[str, Any]:
    """Validate Word document and optionally fix issues"""
    doc = Document(file_path)
    issues = []
    fixes_applied = []
    
    # Check basic document structure
    if not doc.paragraphs:
        issues.append({
            'type': 'structure',
            'severity': 'warning',
            'message': 'Document has no paragraphs'
        })
    
    # Check for empty paragraphs
    empty_paragraphs = len([p for p in doc.paragraphs if not p.text.strip()])
    if empty_paragraphs > 5:
        issues.append({
            'type': 'formatting',
            'severity': 'warning', 
            'message': f'Document has {empty_paragraphs} empty paragraphs (consider cleanup)'
        })
    
    # Check heading structure if style checking enabled
    if check_styles:
        heading_issues = validate_heading_structure(doc)
        issues.extend(heading_issues)
    
    # Check for very long paragraphs
    long_paragraphs = [p for p in doc.paragraphs if len(p.text) > 1000]
    if long_paragraphs:
        issues.append({
            'type': 'readability',
            'severity': 'info',
            'message': f'Found {len(long_paragraphs)} very long paragraphs (>1000 chars)'
        })
    
    # Auto-fix issues if requested
    if fix_formatting:
        if empty_paragraphs > 5:
            # Remove excessive empty paragraphs (keep some for spacing)
            removed_count = remove_excessive_empty_paragraphs(doc)
            if removed_count > 0:
                fixes_applied.append(f'Removed {removed_count} excessive empty paragraphs')
                
                # Save fixed document
                backup_path = file_path.with_suffix('.backup.docx')
                doc.save(backup_path)
                doc.save(file_path)
                fixes_applied.append(f'Document saved with fixes applied (backup: {backup_path.name})')
    
    return {
        'filename': file_path.name,
        'validation_status': 'passed' if not any(i['severity'] == 'error' for i in issues) else 'failed',
        'issues_found': len(issues),
        'issues': issues,
        'fixes_applied': fixes_applied,
        'document_stats': {
            'paragraphs': len(doc.paragraphs),
            'tables': len(doc.tables),
            'empty_paragraphs': empty_paragraphs,
            'long_paragraphs': len(long_paragraphs)
        },
        'timestamp': datetime.now().isoformat()
    }


def validate_heading_structure(doc: Document) -> List[Dict[str, Any]]:
    """Validate document heading structure"""
    issues = []
    heading_levels = []
    
    for p in doc.paragraphs:
        if p.style.name.startswith('Heading'):
            level = int(re.findall(r'\d+', p.style.name)[0]) if re.findall(r'\d+', p.style.name) else 1
            heading_levels.append(level)
    
    # Check for missing heading hierarchy
    for i, level in enumerate(heading_levels[1:], 1):
        prev_level = heading_levels[i-1]
        if level > prev_level + 1:
            issues.append({
                'type': 'style',
                'severity': 'warning',
                'message': f'Heading hierarchy skip detected: H{prev_level} followed by H{level}'
            })
    
    # Check if document starts with non-H1 heading
    if heading_levels and heading_levels[0] != 1:
        issues.append({
            'type': 'style',
            'severity': 'info',
            'message': f'Document starts with H{heading_levels[0]} instead of H1'
        })
    
    return issues


def remove_excessive_empty_paragraphs(doc: Document) -> int:
    """Remove excessive empty paragraphs, keeping some for spacing"""
    paragraphs_to_remove = []
    consecutive_empty = 0
    removed_count = 0
    
    for i, p in enumerate(doc.paragraphs):
        if not p.text.strip():
            consecutive_empty += 1
            # Keep first empty paragraph in sequence, remove subsequent ones
            if consecutive_empty > 2:  # Allow up to 2 consecutive empty paragraphs
                paragraphs_to_remove.append(p)
        else:
            consecutive_empty = 0
    
    # Remove paragraphs (in reverse order to maintain indices)
    for p in reversed(paragraphs_to_remove):
        p_element = p.element
        p_element.getparent().remove(p_element)
        removed_count += 1
    
    return removed_count


def display_validation_results(result: Dict[str, Any], quiet: bool = False) -> None:
    """Display validation results in formatted table"""
    if quiet:
        status = result['validation_status']
        issues = result['issues_found']
        print(f"{result['filename']}: {status}, {issues} issues")
        return
    
    # Main status
    status_color = "green" if result['validation_status'] == 'passed' else "red"
    console.print(f"\n✅ Validation Results: [bold {status_color}]{result['validation_status'].upper()}[/bold {status_color}]")
    console.print(f"📄 File: {result['filename']}")
    console.print(f"🔍 Issues Found: {result['issues_found']}")
    
    # Document stats
    stats = result['document_stats']
    console.print(f"\n📊 Document Statistics:")
    console.print(f"  • Paragraphs: {stats['paragraphs']}")
    console.print(f"  • Tables: {stats['tables']}")
    console.print(f"  • Empty Paragraphs: {stats['empty_paragraphs']}")
    console.print(f"  • Long Paragraphs: {stats['long_paragraphs']}")
    
    # Issues
    if result['issues']:
        console.print(f"\n⚠️  Issues Detected:")
        for issue in result['issues']:
            severity_color = {"error": "red", "warning": "yellow", "info": "blue"}.get(issue['severity'], "white")
            console.print(f"  • [{severity_color}]{issue['severity'].upper()}[/{severity_color}]: {issue['message']} ({issue['type']})")
    
    # Fixes applied
    if result['fixes_applied']:
        console.print(f"\n🔧 Fixes Applied:")
        for fix in result['fixes_applied']:
            console.print(f"  • {fix}")


if __name__ == '__main__':
    cli()