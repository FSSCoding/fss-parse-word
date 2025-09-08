#!/usr/bin/env python3
"""
Word - Global Installation Script
Installs the Word conversion tool to system PATH for global usage.

Author: Isabella (Testing & Validation Specialist)
Safety Features: Path detection, sudo prompts, manual fallback instructions
"""

import os
import sys
import shutil
import stat
from pathlib import Path


class WordInstaller:
    """Handles safe installation of Word tool to system PATH."""
    
    def __init__(self):
        self.script_dir = Path(__file__).parent.absolute()
        self.src_path = self.script_dir / "src" / "word_converter.py"
        
        # Possible installation targets (in order of preference)
        self.install_targets = [
            Path.home() / ".local" / "bin",  # User local bin (preferred)
            Path("/usr/local/bin"),          # System-wide (requires sudo)
            Path("/usr/bin"),                # System bin (requires sudo, less preferred)
        ]
        
    def check_dependencies(self) -> bool:
        """Check if required dependencies are installed."""
        print("🔍 Checking dependencies...")
        
        missing_deps = []
        
        try:
            import docx
            print("  ✅ python-docx")
        except ImportError:
            missing_deps.append("python-docx")
            print("  ❌ python-docx")
        
        try:
            import markdown
            print("  ✅ markdown")
        except ImportError:
            missing_deps.append("markdown")
            print("  ❌ markdown")
        
        try:
            import yaml
            print("  ✅ PyYAML")
        except ImportError:
            print("  ⚠️  PyYAML (optional - for config files)")
        
        if missing_deps:
            print(f"\n❌ Missing dependencies: {', '.join(missing_deps)}")
            print("📦 Install with: pip install " + " ".join(missing_deps))
            if "PyYAML" not in missing_deps:
                print("💡 Optional: pip install PyYAML  # For YAML config support")
            return False
        
        print("✅ All required dependencies installed!")
        return True
    
    def find_best_install_path(self) -> tuple[Path, bool]:
        """
        Find the best installation path.
        Returns (path, requires_sudo)
        """
        for target in self.install_targets:
            if target.exists():
                # Check if we can write to this directory
                if os.access(target, os.W_OK):
                    return target, False
                elif target == Path.home() / ".local" / "bin":
                    # We can create ~/.local/bin if it doesn't exist
                    return target, False
                else:
                    # This path exists but requires sudo
                    return target, True
        
        # Default to user local bin (create if needed)
        return self.install_targets[0], False
    
    def create_executable(self, install_path: Path) -> Path:
        """Create the executable script with embedded source code."""
        executable_path = install_path / "word"
        
        # Read the source code
        with open(self.src_path, 'r', encoding='utf-8') as f:
            source_code = f.read()
        
        # Create standalone executable
        executable_content = f'''#!/usr/bin/env python3
"""
Word - Global Document Conversion Tool
Safe, hash-validated bidirectional conversion between .docx and .md formats.

Standalone executable - contains embedded source code.
Installed by Isabella's Word Installer.
"""

# Embedded source code
{source_code}
'''
        
        return executable_path, executable_content
    
    def install_with_sudo(self, install_path: Path, executable_content: str) -> bool:
        """Install using sudo with clear prompts."""
        executable_path = install_path / "word"
        
        print(f"\n🔐 SUDO REQUIRED - Installing to {install_path}")
        print(f"   Target: {executable_path}")
        print("   This requires administrator privileges.")
        print("\n📋 MANUAL INSTALLATION (for agents/automated workflows):")
        print(f"   sudo mkdir -p {install_path}")
        print(f"   sudo tee {executable_path} > /dev/null << 'WORD_SCRIPT_EOF'")
        print("   [script content would be written here]")
        print("   WORD_SCRIPT_EOF")
        print(f"   sudo chmod +x {executable_path}")
        print("\n🤖 AUTOMATED ALTERNATIVE:")
        print("   Install to ~/.local/bin instead (no sudo required)")
        print("   mkdir -p ~/.local/bin && [copy script] && chmod +x ~/.local/bin/word")
        
        response = input("\n❓ Proceed with sudo installation? [y/N]: ").lower().strip()
        if response not in ['y', 'yes']:
            return False
        
        try:
            # Create directory if needed
            if not install_path.exists():
                result = os.system(f"sudo mkdir -p {install_path}")
                if result != 0:
                    print(f"❌ Failed to create directory {install_path}")
                    return False
            
            # Write the file
            temp_file = Path("/tmp/word_installer_temp")
            with open(temp_file, 'w', encoding='utf-8') as f:
                f.write(executable_content)
            
            result = os.system(f"sudo cp {temp_file} {executable_path}")
            if result != 0:
                print(f"❌ Failed to copy executable to {executable_path}")
                temp_file.unlink(missing_ok=True)
                return False
            
            # Make executable
            result = os.system(f"sudo chmod +x {executable_path}")
            if result != 0:
                print(f"❌ Failed to make {executable_path} executable")
                temp_file.unlink(missing_ok=True)
                return False
            
            temp_file.unlink(missing_ok=True)
            return True
            
        except Exception as e:
            print(f"❌ Installation failed: {e}")
            return False
    
    def install_user_local(self, install_path: Path, executable_content: str) -> bool:
        """Install to user's local bin directory."""
        executable_path = install_path / "word"
        
        try:
            # Create directory if needed
            install_path.mkdir(parents=True, exist_ok=True)
            
            # Write executable
            with open(executable_path, 'w', encoding='utf-8') as f:
                f.write(executable_content)
            
            # Make executable
            executable_path.chmod(executable_path.stat().st_mode | stat.S_IEXEC)
            
            return True
            
        except Exception as e:
            print(f"❌ Installation failed: {e}")
            return False
    
    def check_path_configuration(self, install_path: Path) -> None:
        """Check if install path is in user's PATH."""
        path_env = os.environ.get('PATH', '')
        path_dirs = path_env.split(os.pathsep)
        
        if str(install_path) not in path_dirs:
            print(f"\n⚠️  WARNING: {install_path} is not in your PATH")
            print("🔧 To fix this, add the following to your shell profile:")
            print(f"   echo 'export PATH=\"{install_path}:$PATH\"' >> ~/.bashrc")
            print("   # OR for zsh:")
            print(f"   echo 'export PATH=\"{install_path}:$PATH\"' >> ~/.zshrc")
            print("\n🔄 Then restart your terminal or run:")
            print("   source ~/.bashrc  # or ~/.zshrc")
        else:
            print(f"✅ {install_path} is already in your PATH")
    
    def install(self) -> bool:
        """Perform the installation."""
        print("🚀 Word Document Converter - Global Installation")
        print("=" * 50)
        
        # Check if source exists
        if not self.src_path.exists():
            print(f"❌ Source file not found: {self.src_path}")
            return False
        
        # Check dependencies
        if not self.check_dependencies():
            return False
        
        # Find installation path
        install_path, requires_sudo = self.find_best_install_path()
        executable_path, executable_content = self.create_executable(install_path)
        
        print(f"\n📁 Installation target: {install_path}")
        print(f"🔧 Requires sudo: {'Yes' if requires_sudo else 'No'}")
        
        # Perform installation
        if requires_sudo:
            success = self.install_with_sudo(install_path, executable_content)
        else:
            print(f"📝 Installing to user directory: {install_path}")
            success = self.install_user_local(install_path, executable_content)
        
        if not success:
            print("\n❌ Installation failed!")
            print("\n📋 MANUAL INSTALLATION INSTRUCTIONS:")
            print("1. Ensure dependencies are installed:")
            print("   pip install python-docx markdown PyYAML")
            print("\n2. Copy the script manually:")
            print(f"   mkdir -p ~/.local/bin")
            print(f"   cp {self.src_path} ~/.local/bin/word")
            print(f"   chmod +x ~/.local/bin/word")
            print("\n3. Add to PATH if needed:")
            print("   echo 'export PATH=\"$HOME/.local/bin:$PATH\"' >> ~/.bashrc")
            return False
        
        # Verify installation
        executable_path = install_path / "word"
        if executable_path.exists() and os.access(executable_path, os.X_OK):
            print(f"\n✅ Installation successful!")
            print(f"📍 Installed to: {executable_path}")
            
            # Check PATH configuration
            self.check_path_configuration(install_path)
            
            print(f"\n🎯 Usage:")
            print("   word document.docx document.md    # Convert Word to Markdown")
            print("   word document.md document.docx    # Convert Markdown to Word")
            print("   word --help                       # Show all options")
            
            print(f"\n🛡️ Safety Features:")
            print("   • Hash validation prevents data loss")
            print("   • Automatic backup creation")
            print("   • Collision detection")
            print("   • Confirmation prompts for overwrites")
            
            return True
        else:
            print(f"\n❌ Installation verification failed!")
            print(f"   Expected: {executable_path}")
            return False


def main():
    """Main installation function."""
    installer = WordInstaller()
    success = installer.install()
    sys.exit(0 if success else 1)


if __name__ == '__main__':
    main()