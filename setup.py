#!/usr/bin/env python3
"""
Скрипт установки для системы перевода документов
"""

from setuptools import setup, find_packages
import os

# Читаем README для long_description
def read_readme():
    with open("README.md", "r", encoding="utf-8") as f:
        return f.read()

# Читаем требования
def read_requirements():
    with open("requirements.txt", "r", encoding="utf-8") as f:
        return [line.strip() for line in f if line.strip() and not line.startswith("#")]

setup(
    name="literary-translator",
    version="3.0.1",
    author="Literary Translator Team",
    author_email="support@translator.com",
    description="Система перевода документов с сохранением форматирования и изображений",
    long_description=read_readme(),
    long_description_content_type="text/markdown",
    url="https://github.com/your-repo/literary-translator",
    packages=find_packages(),
    classifiers=[
        "Development Status :: 5 - Production/Stable",
        "Intended Audience :: Developers",
        "Intended Audience :: End Users/Desktop",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Topic :: Office/Business",
        "Topic :: Text Processing :: Linguistic",
        "Topic :: Utilities",
    ],
    python_requires=">=3.8",
    install_requires=read_requirements(),
    extras_require={
        "dev": [
            "pytest>=7.0.0",
            "black>=22.0.0",
            "flake8>=4.0.0",
            "mypy>=0.900",
        ],
    },
    entry_points={
        "console_scripts": [
            "literary-translate=literary_translate:main",
        ],
    },
    include_package_data=True,
    package_data={
        "": ["*.md", "*.txt", "*.example"],
    },
    project_urls={
        "Bug Reports": "https://github.com/your-repo/literary-translator/issues",
        "Source": "https://github.com/your-repo/literary-translator",
        "Documentation": "https://github.com/your-repo/literary-translator/blob/main/README.md",
    },
    keywords="translation document docx formatting images literary",
    zip_safe=False,
) 