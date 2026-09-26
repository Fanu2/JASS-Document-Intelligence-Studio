# JASS Document Intelligence Studio

### Universal Local-First Document & File Research Workspace

JASS Document Intelligence Studio is a **local-first research workspace** for importing, cataloguing, inspecting, searching, organizing and studying files on your own computer.

It is designed as part of **JASS Digital Lab** and follows a simple principle:

> **Bring your files together. Understand them. Organize them. Research them.**

The application uses a local SQLite database and does not require a cloud service for its core library functions.

---

## ✨ What It Does

JASS Document Intelligence Studio provides a unified library for many different kinds of files.

### Universal Import

The application can import:

- PDF
- DOCX
- TXT
- Markdown
- CSV
- TSV
- JSON
- XML
- HTML
- YAML
- RTF
- Python
- JavaScript
- TypeScript
- C / C++
- Java
- SQL
- PowerShell
- Shell scripts
- CSS
- Logs
- Configuration files
- Images
- Audio
- Video
- Other file types

The important design principle is:

> **A file does not have to be fully understood to be preserved in the library.**

Supported formats receive appropriate extraction and/or preview. Other formats can still be catalogued and retained as library items.

---

## 🖼️ Image Support

Images are first-class library items.

Supported image formats include:

- PNG
- JPG / JPEG
- GIF
- BMP
- WebP
- TIFF
- SVG
- ICO

Images can be imported into the library and viewed through the application's image workspace.

PDF documents can also have their embedded images extracted and displayed.

---

## 📄 Document Support

### PDF

PDF support includes:

- Text extraction
- Page count
- PDF page thumbnails
- Visual page rendering
- Page navigation
- Embedded image extraction
- Image gallery
- Searchable extracted text

### DOCX

DOCX files can be imported and their text extracted for research and searching.

### Text & Markdown

TXT and Markdown files are imported directly and become searchable library content.

---

## 🔎 Research Search

Search across the local document library using:

- Filename
- Extracted text
- Category
- Tags
- Notes

The Research Search workspace provides:

- Matching documents
- Match counts
- Context snippets
- Quick opening of matching documents

There is also a **Find** function for searching within the currently open document.

---

## 🗂️ Research Collections

Documents can be organized into research collections such as:

- Project Athena
- Punjabi Language
- Mizo Research
- Literature
- AI Research
- HRTK
- Personal Research

Collections do not move or modify the original files.

They simply provide a logical research organization layer.

---

## ⭐ Favorites

Important documents can be added to a personal reading list.

Features include:

- Add to Favorites
- Remove from Favorites
- Favorites workspace
- Persistent favorite status

This makes it easy to maintain a focused set of documents for current research.

---

## 🏷️ Metadata & Organization

Each library item can have:

- Category
- Tags
- Research notes
- Collections
- Relationships

Built-in categories include:

- Research
- Literature
- Language
- Technical
- Project
- Data
- Reference
- Personal
- Other

---

## 🔗 Document Relationships

Documents can be connected to one another using relationships such as:

- Related to
- References
- Derived from
- Part of
- Supports
- Uses

For example:

```text
RAG Research Paper
        │
        ├── References → Project Athena
        │
        └── Supports → Retrieval Notes
```

This provides the foundation for a future research knowledge graph.

---

## 📊 Document Statistics

The Statistics workspace provides information such as:

- Filename
- File type
- Category
- Collections
- File size
- Pages
- Word count
- Character count
- Extracted images
- Import time
- Modification time
- SHA-1 checksum
- Original file path

---

## 🔐 Local-First Design

JASS Document Intelligence Studio is designed around local processing.

### Core principles

- Local-first
- Offline-friendly
- Privacy-conscious
- SQLite-based
- Original files are never modified
- No mandatory cloud service
- User-controlled document library

The database is stored by default under:

```text
~/JASS/DocumentIntelligenceStudio/
```

On Windows this normally resolves to:

```text
C:\Users\<username>\JASS\DocumentIntelligenceStudio\
```

The primary database is:

```text
jass_documents.db
```

---

## 🧠 Current Architecture

```text
                     JASS Document Intelligence Studio
                                  │
                                  ▼
                         Universal File Import
                                  │
             ┌────────────────────┼────────────────────┐
             ▼                    ▼                    ▼
        Documents              Images          Media / Other
             │                    │                    │
             └────────────────────┼────────────────────┘
                                  ▼
                           SQLite Library
                                  │
              ┌───────────────────┼───────────────────┐
              ▼                   ▼                   ▼
           Search             Collections         Metadata
              │                   │                   │
              ├──────────────┬────┴────┬──────────────┤
              ▼              ▼         ▼              ▼
          Favorites      Relations   Notes          Tags
                                  │
                                  ▼
                           Research Workspace
```

---

## 🛠️ Technology Stack

- **Python**
- **PySide6**
- **SQLite**
- **PyMuPDF**
- **python-docx**

The application is designed as a lightweight desktop application rather than a browser-dependent research service.

---

## 📦 Installation

Create a Python environment and install the required packages:

```powershell
pip install PySide6 PyMuPDF python-docx
```

Then run:

```powershell
python jass_document_intelligence_studio_v0_7.py
```

---

## 💻 Typical Workflow

### 1. Import

Bring research material into the library:

```text
PDF
DOCX
TXT
Markdown
Images
CSV
JSON
Code
Media
Other files
```

### 2. Inspect

Open a document and inspect:

```text
Text
PDF Pages
Images
Statistics
```

### 3. Organize

Add:

```text
Category
Tags
Notes
Collections
Favorites
Relationships
```

### 4. Search

Use:

```text
Library Search
Research Search
Find in Document
```

### 5. Build Research Collections

For example:

```text
Project Athena
│
├── Architecture papers
├── RAG research
├── AI provider documentation
├── Retrieval experiments
└── Personal research notes
```

---

## 🚧 Current Scope

JASS Document Intelligence Studio v0.7 is primarily a **file intelligence, organization and research workspace**.

It already provides the foundation for more advanced document intelligence.

### Available now

- Universal file import
- Local SQLite library
- PDF rendering
- PDF text extraction
- PDF image extraction
- Image import and viewing
- DOCX extraction
- Text/Markdown extraction
- Search
- Research Search
- Collections
- Favorites
- Tags
- Notes
- Relationships
- Categories
- Statistics
- Duplicate detection
- Text export
- Local/offline operation

### Planned / Future Direction

The architecture can evolve toward:

- OCR
- Image-to-text extraction
- Semantic search
- Embeddings
- Vector databases
- RAG
- Document summarization
- AI-assisted research
- Question answering over documents
- Citation-aware answers
- Knowledge graph features
- Athena integration
- Local AI providers
- Document classification
- Multimodal research

The goal is to add these capabilities incrementally without turning the application into an unnecessarily heavy platform.

---

## 🔗 Relationship to Project Athena

JASS Document Intelligence Studio can serve as a practical document research foundation for **Project Athena**.

A possible future architecture is:

```text
JASS Document Intelligence Studio
              │
              ▼
       Document Library
              │
              ▼
       Document Intelligence
              │
       ┌──────┴──────┐
       ▼             ▼
   Retrieval       AI/RAG
       │             │
       └──────┬──────┘
              ▼
        Project Athena
```

This keeps document ingestion and organization useful independently while allowing Athena to consume the resulting knowledge layer later.

---

## 🎯 Design Philosophy

### Library First. Intelligence Second.

The first responsibility of the application is to build a reliable local library.

AI capabilities come afterward.

This approach provides:

- Better control of personal research material
- Clear document provenance
- Local organization before AI processing
- Easier debugging
- Modular development
- Reduced dependence on external services

---

## 🔒 Data Safety

The application calculates a SHA-1 checksum for imported files to help identify duplicates.

Removing a document from the library does **not** delete the original file.

The application stores references and extracted research data in its local database while leaving the source file untouched.

---

## 🌱 JASS Digital Lab

JASS Document Intelligence Studio is part of **JASS Digital Lab** — a personal software laboratory focused on:

- AI & Machine Learning
- Knowledge Systems
- Language & Literature
- Data Research
- Document Intelligence
- Creative Computing
- Legacy Software Modernization
- Desktop Software

The broader philosophy is:

> **Build useful software locally, preserve knowledge, and modernize what matters.**

---

## 📌 Project Status

**Version:** `0.7`

**Status:** Active Development

**Platform:** Desktop

**Primary language:** Python

**GUI:** PySide6

**Database:** SQLite

**Architecture:** Local-first

---

## 👤 Author

**SinghJasvir · Fanu2**

Independent Software Builder  
**JASS Digital Lab**

---

## 📄 License

Add the project's preferred license here when finalized.

---

### JASS Digital Lab

**AI · Knowledge · Language · Data · Creative Computing**
