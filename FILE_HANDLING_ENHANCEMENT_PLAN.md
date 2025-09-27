# Word Parser File Handling Enhancement Plan

**Focus**: Address file handling gaps and failure modes for improved accessibility and reliability

## Current Format Support Analysis

### Python Word Parser (`word/`)
- **Supported**: `.docx` (Word 2007+), `.md` (Markdown)
- **Missing**: `.rtf`, `.doc` (legacy), `.odt`, password-protected files
- **Dependencies**: `python-docx` (limited to DOCX format only)

### TypeScript Word Parser (`word-ts/`)  
- **Claimed Support**: DOCX, DOC, RTF, ODT (per README line 8)
- **Reality Check Needed**: Verify actual implementation vs documented support
- **Dependencies**: Unknown - needs investigation

## Priority File Handling Gaps

### 1. Legacy Format Support ⚠️ HIGH IMPACT
**Problem**: `.rtf` and `.doc` files still common in government/legacy systems
**Current Behavior**: Hard failure with unclear error messages
**Accessible Solution**:
```python
def detect_file_format(file_path: str) -> Dict[str, Any]:
    """Detect file format and provide clear feedback"""
    extension = Path(file_path).suffix.lower()
    
    format_info = {
        '.docx': {'supported': True, 'processor': 'python-docx'},
        '.rtf': {'supported': False, 'alternative': 'Convert to .docx first'},
        '.doc': {'supported': False, 'alternative': 'Use Word to save as .docx'},
        '.odt': {'supported': False, 'alternative': 'Use LibreOffice to export as .docx'}
    }
    
    return {
        'format': extension,
        'supported': format_info.get(extension, {}).get('supported', False),
        'message': format_info.get(extension, {}).get('alternative', 'Unknown format'),
        'suggested_action': f"Please convert {extension} to .docx format"
    }
```

### 2. Password-Protected Files 🔒 MEDIUM IMPACT  
**Problem**: Encrypted Word documents cause crashes
**Current Behavior**: python-docx throws cryptic errors
**Accessible Solution**:
```python
def check_file_protection(file_path: str) -> Dict[str, Any]:
    """Check if file is password-protected before processing"""
    try:
        # Attempt to open document
        Document(file_path)
        return {'protected': False, 'accessible': True}
    except Exception as e:
        if 'password' in str(e).lower() or 'encrypted' in str(e).lower():
            return {
                'protected': True, 
                'accessible': False,
                'message': 'Document is password-protected',
                'suggested_action': 'Remove password protection in Word first'
            }
        return {
            'protected': False,
            'accessible': False, 
            'error': str(e),
            'suggested_action': 'Check file integrity'
        }
```

### 3. Large File Streaming 📊 MEDIUM IMPACT
**Problem**: Memory exhaustion on large documents (>100MB)
**Current Behavior**: System freeze or crash
**Accessible Solution**:
```python
class FileHandler:
    MAX_FILE_SIZE = 100 * 1024 * 1024  # 100MB
    
    def validate_file_size(self, file_path: str) -> Dict[str, Any]:
        """Check file size before processing"""
        size = Path(file_path).stat().st_size
        
        if size > self.MAX_FILE_SIZE:
            return {
                'valid': False,
                'size_mb': round(size / 1024 / 1024, 1),
                'limit_mb': round(self.MAX_FILE_SIZE / 1024 / 1024, 1),
                'message': f'File too large ({size_mb}MB). Maximum supported: {limit_mb}MB',
                'suggested_action': 'Split document into smaller sections'
            }
        
        return {'valid': True, 'size_mb': round(size / 1024 / 1024, 1)}
```

### 4. Embedded Object Detection 🔍 LOW IMPACT
**Problem**: Embedded Excel/images silently skipped
**Current Behavior**: No indication that content was lost
**Accessible Solution**:
```python
def analyze_document_structure(doc_path: str) -> Dict[str, Any]:
    """Analyze document for embedded objects and complex elements"""
    doc = Document(doc_path)
    analysis = {
        'total_paragraphs': len(doc.paragraphs),
        'embedded_objects': [],
        'images': 0,
        'tables': len(doc.tables),
        'warnings': []
    }
    
    # Check for embedded objects (shapes, OLE objects)
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            if hasattr(run, '_element'):
                # Look for embedded objects in XML
                if 'object' in str(run._element.xml).lower():
                    analysis['embedded_objects'].append({
                        'type': 'unknown_object',
                        'paragraph': paragraph.text[:50]
                    })
    
    # Count images
    for rel in doc.part.rels.values():
        if 'image' in rel.target_ref:
            analysis['images'] += 1
    
    # Generate accessibility warnings
    if analysis['embedded_objects']:
        analysis['warnings'].append(
            f"Found {len(analysis['embedded_objects'])} embedded objects that may not convert properly"
        )
    
    if analysis['images'] > 0:
        analysis['warnings'].append(
            f"Found {analysis['images']} images. Alt text may not be preserved in conversion"
        )
    
    return analysis
```

### 5. Metadata Privacy Controls 🛡️ MEDIUM IMPACT
**Problem**: Document metadata can leak sensitive information
**Current Behavior**: All metadata preserved without user control
**Accessible Solution**:
```python
@click.option('--strip-metadata', is_flag=True, help='Remove document metadata for privacy')
@click.option('--preserve-metadata', is_flag=True, help='Keep all document metadata (default)')
def convert_with_privacy_controls(input_file, output_file, strip_metadata, preserve_metadata):
    """Convert with configurable metadata handling"""
    
    if strip_metadata and preserve_metadata:
        click.echo("❌ Error: Cannot both strip and preserve metadata. Choose one option.")
        return
    
    # Default behavior
    strip_mode = strip_metadata if strip_metadata else False
    
    if strip_mode:
        click.echo("🔒 Privacy mode: Removing document metadata")
        # Implementation to strip author, creation date, revision history
    else:
        click.echo("📝 Standard mode: Preserving document metadata")
```

## Enhanced Error Handling Implementation

### Graceful Failure System
```python
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

def safe_document_processing(file_path: str) -> Dict[str, Any]:
    """Process document with comprehensive error handling"""
    try:
        # Pre-flight checks
        format_check = detect_file_format(file_path)
        if not format_check['supported']:
            raise UnsupportedFormatError(
                format_check['message'],
                suggested_action=format_check['suggested_action'],
                error_code='UNSUPPORTED_FORMAT'
            )
        
        size_check = validate_file_size(file_path)
        if not size_check['valid']:
            raise FileSizeError(
                size_check['message'],
                suggested_action=size_check['suggested_action'],
                error_code='FILE_TOO_LARGE'
            )
        
        protection_check = check_file_protection(file_path)
        if protection_check['protected']:
            raise ProtectedFileError(
                protection_check['message'],
                suggested_action=protection_check['suggested_action'],
                error_code='PASSWORD_PROTECTED'
            )
        
        # Proceed with processing
        return {'status': 'ready', 'checks_passed': True}
        
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
            'suggested_action': 'Please check file integrity and try again',
            'user_friendly': False
        }
```

## Batch Processing Improvements

### Enhanced Batch Error Handling
```python
def process_batch_with_resilience(input_dir: Path, output_dir: Path, 
                                pattern: str = "*.docx") -> Dict[str, Any]:
    """Process multiple files with detailed error reporting"""
    results = {
        'processed': 0,
        'successful': 0,
        'failed': 0,
        'skipped': 0,
        'errors': [],
        'warnings': []
    }
    
    files = list(input_dir.glob(pattern))
    
    for file_path in files:
        results['processed'] += 1
        
        try:
            # Pre-flight validation
            validation = safe_document_processing(str(file_path))
            
            if validation['status'] == 'error':
                results['failed'] += 1
                results['errors'].append({
                    'file': str(file_path),
                    'error': validation['message'],
                    'action': validation['suggested_action'],
                    'code': validation.get('error_code')
                })
                continue
            
            # Process file
            # ... conversion logic ...
            results['successful'] += 1
            
        except Exception as e:
            results['failed'] += 1
            results['errors'].append({
                'file': str(file_path),
                'error': str(e),
                'action': 'Check file manually',
                'code': 'PROCESSING_ERROR'
            })
    
    return results
```

## Implementation Priority

### Phase 1: Core Safety (Week 1)
1. ✅ Implement file format detection with clear messaging
2. ✅ Add password protection detection  
3. ✅ Create file size validation with user-friendly limits
4. ✅ Enhance error messages for accessibility

### Phase 2: Advanced Features (Week 2)  
1. ✅ Add metadata privacy controls
2. ✅ Implement embedded object detection
3. ✅ Create batch processing resilience
4. ✅ Add comprehensive validation reporting

### Phase 3: TypeScript Alignment (Week 3)
1. ✅ Verify TypeScript parser actual format support
2. ✅ Implement parallel enhancements in word-ts/
3. ✅ Ensure feature parity between implementations
4. ✅ Create unified error handling approach

## Testing Strategy

### Real-World Test Files
- Password-protected Word documents
- Large documents (>50MB)
- Legacy .rtf and .doc files
- Documents with embedded Excel/images
- Corrupted/partial Word files

### Accessibility Testing
- Screen reader compatibility with error messages
- Clear, actionable feedback for all failure modes
- Progressive disclosure of technical details
- Keyboard navigation for CLI prompts

## Success Metrics

- **Error Clarity**: 100% of errors include suggested actions
- **Format Support**: Clear documentation of what is/isn't supported
- **Performance**: Large file handling without system crashes
- **Privacy**: Configurable metadata handling
- **Reliability**: Batch processing continues despite individual file failures

---

*This plan focuses on making Word parsers more accessible, reliable, and user-friendly through comprehensive error handling and clear communication.*