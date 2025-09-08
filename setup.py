#!/usr/bin/env python3
"""Setup script for FSS Parse Word - Document Parser Tool."""

from setuptools import setup, find_packages
from pathlib import Path

# Read README for long description
readme_path = Path(__file__).parent / "README.md"
long_description = readme_path.read_text(encoding="utf-8") if readme_path.exists() else ""

setup(
    name="fss-parse-word",
    version="1.0.0",
    description="Robust bidirectional parser between Word documents (.docx) and Markdown (.md)",
    long_description=long_description,
    long_description_content_type="text/markdown",
    author="FssCoding",
    author_email="",
    url="https://github.com/FSSCoding/fss-parse-word",
    packages=find_packages(),
    package_dir={"": "src"},
    install_requires=[
        "python-docx>=0.8.11",
        "markdown>=3.4.0",
        "PyYAML>=6.0",
    ],
    entry_points={
        "console_scripts": [
            "fss-parse-word=word_converter:main",
        ],
    },
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "Intended Audience :: End Users/Desktop",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Topic :: Text Processing :: Markup",
        "Topic :: Office/Business :: Office Suites",
        "Topic :: Software Development :: Libraries :: Python Modules",
    ],
    python_requires=">=3.8",
    keywords="docx markdown word parser document conversion fss",
    project_urls={
        "Bug Reports": "https://github.com/FSSCoding/fss-parse-word/issues",
        "Source": "https://github.com/FSSCoding/fss-parse-word",
    },
)