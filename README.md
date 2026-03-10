# Python-Document-Automation-Tool
Automated system for generating personalized .docx files from Excel data and organizing them using Regex.
Multi-Template Document Automation and Organization Suite
Overview
This project provides a Python-based solution for automating the generation and organization of administrative documents. It is designed to handle large datasets from Excel and map them into multiple Word templates while maintaining structural integrity and formatting.

Features
Document Generator
Processes .docx templates using the python-docx library.

Maps Excel columns to document placeholders with formatting preservation.

Includes logic for automatic field generation such as formatted dates and name concatenations.

Supports text replacement in paragraphs, tables, headers, and footers.

File Organizer
Scans directory trees to identify and group .docx files by person.

Implements Regex patterns with support for Romanian diacritics to ensure accurate name extraction from filenames.

Automates the creation of a structured directory hierarchy (FISIER ORGANIZAT).

Handles file system operations including safe copying and error logging for missing or malformed data.

Requirements
Python 3.x

pandas

python-docx

openpyxl

Usage
Provide the source Excel file path and the directory containing the .docx templates.

The generator module populates the templates based on the Excel rows.

The organizer module sorts the resulting files into individual folders based on the identified subject names.
