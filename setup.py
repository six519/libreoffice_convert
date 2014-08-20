#!/usr/bin/env python3

try:
    from setuptools import setup
except ImportError:
    from distutils.core import setup

setup(
    name="libreoffice_convert",
    version="1.0",
    description="Python module using LibreOffice API to convert file format to another file format",
    author="Ferdinand Silva",
    author_email="ferdinandsilva@ferdinandsilva.com",
    packages=["libreoffice_convert"],
    url="http://ferdinandsilva.com",
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "Natural Language :: English",
        "Programming Language :: Python",
        "Programming Language :: Python :: 3",
        "License :: Freeware",
    ],
    entry_points = {
        "console_scripts": [
            "libreoffice_convert = libreoffice_convert.commands:libreoffice_convert"]
    },
    download_url="http://ferdinandsilva.com",
)