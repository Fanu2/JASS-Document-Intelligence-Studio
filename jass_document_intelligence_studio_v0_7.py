#!/usr/bin/env python3
"""
JASS Document Intelligence Studio
v0.7 — Universal File Research Workspace

Local-first document workspace for:
    Universal file import with rich handling for documents, images, data, code, audio, video and unknown/binary files.

Core:
    - SQLite document library
    - Real PDF page rendering
    - PDF page thumbnails
    - Document text preview
    - Search across library
    - Find/highlight inside document
    - Direct image import and preview
    - Extracted PDF image gallery
    - Universal import for any file extension
    - Tags, categories and notes
    - Related documents
    - Document statistics
    - Export text
    - Original files are never modified

Optional:
    pip install PySide6 PyMuPDF python-docx
"""

from __future__ import annotations

import hashlib
import mimetypes
import re
import sqlite3
import sys
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import (
    QAction, QColor, QImage, QPixmap, QTextCharFormat, QTextCursor
)
from PySide6.QtWidgets import (
    QApplication, QAbstractItemView, QComboBox, QDialog, QFileDialog,
    QFormLayout, QHBoxLayout, QLabel, QLineEdit, QListWidget,
    QListWidgetItem, QMainWindow, QMessageBox, QPlainTextEdit, QPushButton,
    QScrollArea, QSplitter, QStatusBar, QTabWidget, QTextBrowser,
    QToolBar, QVBoxLayout, QWidget
)

try:
    import pymupdf
except ImportError:
    pymupdf = None

try:
    from docx import Document as DocxDocument
except ImportError:
    DocxDocument = None


APP_NAME = "JASS Document Intelligence Studio"
VERSION = "0.7"

# Known formats receive specialized extraction/preview. Everything else is
# still importable and preserved as a library item.
SUPPORTED = {
    ".pdf": "PDF", ".docx": "DOCX", ".txt": "Text", ".md": "Markdown",
    ".rtf": "Rich Text", ".csv": "CSV", ".tsv": "TSV", ".json": "JSON",
    ".xml": "XML", ".html": "HTML", ".htm": "HTML", ".yaml": "YAML",
    ".yml": "YAML", ".log": "Log", ".ini": "INI", ".cfg": "Config",
    ".py": "Python", ".js": "JavaScript", ".ts": "TypeScript",
    ".java": "Java", ".c": "C", ".cpp": "C++", ".h": "C/C++ Header",
    ".sql": "SQL", ".sh": "Shell", ".ps1": "PowerShell",
    ".jpg": "Image", ".jpeg": "Image", ".png": "Image", ".gif": "Image",
    ".bmp": "Image", ".webp": "Image", ".tif": "Image", ".tiff": "Image",
    ".svg": "Image", ".ico": "Image",
    ".mp3": "Audio", ".wav": "Audio", ".flac": "Audio", ".m4a": "Audio",
    ".aac": "Audio", ".ogg": "Audio", ".wma": "Audio",
    ".mp4": "Video", ".mkv": "Video", ".avi": "Video", ".mov": "Video",
    ".webm": "Video", ".wmv": "Video",
}

IMAGE_EXTENSIONS = {
    ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp",
    ".tif", ".tiff", ".svg", ".ico"
}

TEXT_EXTENSIONS = {
    ".txt", ".md", ".rst", ".csv", ".tsv", ".json", ".xml",
    ".html", ".htm", ".yaml", ".yml", ".log", ".ini", ".cfg",
    ".py", ".js", ".ts", ".java", ".c", ".cpp", ".h", ".hpp",
    ".sql", ".sh", ".ps1", ".bat", ".cmd", ".css", ".scss",
    ".tex", ".rtf",
}

CATEGORIES = [
    "Research", "Literature", "Language", "Technical",
    "Project", "Data", "Reference", "Personal", "Other"
]


def now_text():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def sha1_file(path: Path):
    h = hashlib.sha1()
    with path.open("rb") as f:
        for block in iter(lambda: f.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def extract_document(path: Path):
    """Return text, page_count, [(page, image_bytes)].

    Universal importer: text/document formats are extracted as text, images
    are stored for visual preview, and every other file remains importable
    as a preserved binary/media library item.
    """
    ext = path.suffix.lower()

    if ext in IMAGE_EXTENSIONS:
        return "", 1, [(0, path.read_bytes())]

    if ext in TEXT_EXTENSIONS:
        return path.read_text(encoding="utf-8", errors="replace"), 1, []

    if ext == ".docx":
        if DocxDocument is None:
            raise RuntimeError("Install python-docx for DOCX support.")
        doc = DocxDocument(str(path))
        text = "\n\n".join(p.text for p in doc.paragraphs if p.text.strip())
        return text, 1, []

    if ext == ".pdf":
        if pymupdf is None:
            raise RuntimeError("Install PyMuPDF for PDF support.")
        doc = pymupdf.open(str(path))
        chunks, images = [], []
        for page_no, page in enumerate(doc, 1):
            txt = page.get_text("text").strip()
            if txt:
                chunks.append(f"[Page {page_no}]\n{txt}")
            for img in page.get_images(full=True):
                try:
                    data = doc.extract_image(img[0])
                    images.append((page_no, data["image"]))
                except Exception:
                    pass
        text = "\n\n".join(chunks)
        count = len(doc)
        doc.close()
        return text, count, images

    # Unknown, binary, audio and video files are valid imports.
    return "", 1, []


class Database:
    def __init__(self, path: Path):
        self.path = path
        self.conn = sqlite3.connect(str(path))
        self.conn.row_factory = sqlite3.Row
        self.conn.executescript("""
        CREATE TABLE IF NOT EXISTS documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            path TEXT NOT NULL,
            filename TEXT NOT NULL,
            extension TEXT NOT NULL,
            category TEXT DEFAULT 'Other',
            tags TEXT DEFAULT '',
            notes TEXT DEFAULT '',
            text TEXT DEFAULT '',
            size INTEGER DEFAULT 0,
            pages INTEGER DEFAULT 0,
            image_count INTEGER DEFAULT 0,
            checksum TEXT DEFAULT '',
            imported_at TEXT NOT NULL,
            modified_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS relations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            source_id INTEGER NOT NULL,
            target_id INTEGER NOT NULL,
            relation TEXT DEFAULT 'Related to',
            UNIQUE(source_id, target_id, relation)
        );

        CREATE TABLE IF NOT EXISTS extracted_images (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            document_id INTEGER NOT NULL,
            page INTEGER DEFAULT 0,
            image BLOB NOT NULL
        );

        CREATE TABLE IF NOT EXISTS collections (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            description TEXT DEFAULT '',
            created_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS collection_documents (
            collection_id INTEGER NOT NULL,
            document_id INTEGER NOT NULL,
            UNIQUE(collection_id, document_id)
        );

        CREATE TABLE IF NOT EXISTS favorites (
            document_id INTEGER PRIMARY KEY,
            added_at TEXT NOT NULL
        );
        """)
        self.conn.commit()

    def documents(self, query="", category="All categories"):
        if query.strip():
            q = f"%{query.strip()}%"
            rows = self.conn.execute("""
                SELECT * FROM documents
                WHERE filename LIKE ? OR text LIKE ? OR tags LIKE ?
                   OR notes LIKE ? OR category LIKE ?
                ORDER BY filename COLLATE NOCASE
            """, (q, q, q, q, q)).fetchall()
        else:
            rows = self.conn.execute(
                "SELECT * FROM documents ORDER BY filename COLLATE NOCASE"
            ).fetchall()

        if category != "All categories":
            rows = [r for r in rows if r["category"] == category]

        return rows

    def get(self, doc_id):
        return self.conn.execute(
            "SELECT * FROM documents WHERE id=?", (doc_id,)
        ).fetchone()

    def add(self, data):
        cur = self.conn.execute("""
            INSERT INTO documents
            (path, filename, extension, category, tags, notes, text,
             size, pages, image_count, checksum, imported_at, modified_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            data["path"], data["filename"], data["extension"],
            data["category"], data.get("tags", ""), data.get("notes", ""),
            data.get("text", ""), data.get("size", 0), data.get("pages", 0),
            data.get("image_count", 0), data.get("checksum", ""),
            now_text(), now_text()
        ))
        self.conn.commit()
        return cur.lastrowid

    def update(self, doc_id, **fields):
        if not fields:
            return
        fields["modified_at"] = now_text()
        keys = list(fields)
        sql = (
            "UPDATE documents SET " +
            ", ".join(f"{k}=?" for k in keys) +
            " WHERE id=?"
        )
        self.conn.execute(sql, [fields[k] for k in keys] + [doc_id])
        self.conn.commit()

    def remove(self, doc_id):
        self.conn.execute(
            "DELETE FROM extracted_images WHERE document_id=?", (doc_id,)
        )
        self.conn.execute(
            "DELETE FROM relations WHERE source_id=? OR target_id=?",
            (doc_id, doc_id)
        )
        self.conn.execute("DELETE FROM documents WHERE id=?", (doc_id,))
        self.conn.commit()

    def add_images(self, doc_id, images):
        self.conn.executemany(
            "INSERT INTO extracted_images(document_id,page,image) VALUES(?,?,?)",
            [(doc_id, page, blob) for page, blob in images]
        )
        self.conn.commit()

    def images(self, doc_id):
        return self.conn.execute(
            "SELECT page,image FROM extracted_images "
            "WHERE document_id=? ORDER BY id",
            (doc_id,)
        ).fetchall()

    def add_relation(self, source, target, relation):
        self.conn.execute("""
            INSERT OR IGNORE INTO relations(source_id,target_id,relation)
            VALUES(?,?,?)
        """, (source, target, relation))
        self.conn.commit()

    def relations(self, doc_id):
        return self.conn.execute("""
            SELECT r.relation, d.id, d.filename
            FROM relations r
            JOIN documents d ON d.id =
                CASE WHEN r.source_id=? THEN r.target_id
                     ELSE r.source_id END
            WHERE r.source_id=? OR r.target_id=?
            ORDER BY d.filename COLLATE NOCASE
        """, (doc_id, doc_id, doc_id)).fetchall()

    def stats(self):
        return self.conn.execute("""
            SELECT COUNT(*) documents,
                   COALESCE(SUM(size),0) bytes,
                   COALESCE(SUM(pages),0) pages,
                   COALESCE(SUM(image_count),0) images
            FROM documents
        """).fetchone()

    def collections(self):
        return self.conn.execute(
            "SELECT * FROM collections ORDER BY name COLLATE NOCASE"
        ).fetchall()

    def add_collection(self, name, description=""):
        cur = self.conn.execute(
            "INSERT OR IGNORE INTO collections(name,description,created_at) VALUES(?,?,?)",
            (name.strip(), description.strip(), now_text())
        )
        self.conn.commit()
        return cur.lastrowid

    def get_collection(self, collection_id):
        return self.conn.execute(
            "SELECT * FROM collections WHERE id=?", (collection_id,)
        ).fetchone()

    def delete_collection(self, collection_id):
        self.conn.execute(
            "DELETE FROM collection_documents WHERE collection_id=?",
            (collection_id,)
        )
        self.conn.execute(
            "DELETE FROM collections WHERE id=?", (collection_id,)
        )
        self.conn.commit()

    def collection_documents(self, collection_id):
        return self.conn.execute("""
            SELECT d.* FROM documents d
            JOIN collection_documents cd ON cd.document_id=d.id
            WHERE cd.collection_id=?
            ORDER BY d.filename COLLATE NOCASE
        """, (collection_id,)).fetchall()

    def add_to_collection(self, collection_id, document_id):
        self.conn.execute(
            "INSERT OR IGNORE INTO collection_documents(collection_id,document_id) VALUES(?,?)",
            (collection_id, document_id)
        )
        self.conn.commit()

    def remove_from_collection(self, collection_id, document_id):
        self.conn.execute(
            "DELETE FROM collection_documents WHERE collection_id=? AND document_id=?",
            (collection_id, document_id)
        )
        self.conn.commit()

    def collection_for_document(self, document_id):
        return self.conn.execute("""
            SELECT c.* FROM collections c
            JOIN collection_documents cd ON cd.collection_id=c.id
            WHERE cd.document_id=?
            ORDER BY c.name COLLATE NOCASE
        """, (document_id,)).fetchall()

    def collection_document_ids(self, collection_id):
        rows = self.conn.execute(
            "SELECT document_id FROM collection_documents WHERE collection_id=?",
            (collection_id,)
        ).fetchall()
        return {row["document_id"] for row in rows}

    def is_favorite(self, document_id):
        return self.conn.execute(
            "SELECT 1 FROM favorites WHERE document_id=?", (document_id,)
        ).fetchone() is not None

    def set_favorite(self, document_id, enabled):
        if enabled:
            self.conn.execute(
                "INSERT OR IGNORE INTO favorites(document_id, added_at) VALUES(?,?)",
                (document_id, now_text())
            )
        else:
            self.conn.execute(
                "DELETE FROM favorites WHERE document_id=?", (document_id,)
            )
        self.conn.commit()

    def favorite_documents(self):
        return self.conn.execute("""
            SELECT d.* FROM documents d
            JOIN favorites f ON f.document_id=d.id
            ORDER BY f.added_at DESC
        """).fetchall()

    def cross_document_search(self, query):
        query = query.strip()
        if not query:
            return []

        rows = self.conn.execute("""
            SELECT id, filename, category, text
            FROM documents
            WHERE text LIKE ? OR filename LIKE ? OR tags LIKE ? OR notes LIKE ?
            ORDER BY filename COLLATE NOCASE
        """, (f"%{query}%", f"%{query}%", f"%{query}%", f"%{query}%")).fetchall()

        results = []
        needle = query.lower()

        for row in rows:
            text_value = row["text"] or ""
            lower = text_value.lower()
            positions = []
            start = 0
            while True:
                pos = lower.find(needle, start)
                if pos < 0:
                    break
                positions.append(pos)
                start = pos + max(1, len(needle))

            snippets = []
            for pos in positions[:5]:
                a = max(0, pos - 90)
                b = min(len(text_value), pos + len(query) + 140)
                snippet = text_value[a:b].replace("\n", " ").strip()
                snippets.append(("..." if a else "") + snippet +
                                ("..." if b < len(text_value) else ""))

            results.append({
                "id": row["id"],
                "filename": row["filename"],
                "category": row["category"],
                "matches": len(positions),
                "snippets": snippets,
            })

        return results

    def close(self):
        self.conn.close()


class RelationDialog(QDialog):
    def __init__(self, parent, docs, source_id):
        super().__init__(parent)
        self.setWindowTitle("Add Document Relationship")
        self.setMinimumWidth(420)

        form = QFormLayout(self)
        self.target = QComboBox()
        for d in docs:
            if d["id"] != source_id:
                self.target.addItem(d["filename"], d["id"])

        self.relation = QComboBox()
        self.relation.addItems([
            "Related to", "References", "Derived from",
            "Part of", "Supports", "Uses"
        ])

        form.addRow("Document", self.target)
        form.addRow("Relationship", self.relation)

        buttons = QHBoxLayout()
        ok = QPushButton("Add")
        cancel = QPushButton("Cancel")
        ok.clicked.connect(self.accept)
        cancel.clicked.connect(self.reject)
        buttons.addWidget(ok)
        buttons.addWidget(cancel)
        form.addRow(buttons)

    def values(self):
        return self.target.currentData(), self.relation.currentText()


class ResearchSearchDialog(QDialog):
    def __init__(self, parent, db):
        super().__init__(parent)
        self.parent_window = parent
        self.db = db
        self.setWindowTitle("JASS Research Search")
        self.resize(900, 620)

        layout = QVBoxLayout(self)

        intro = QLabel(
            "<b>Search across your document library</b><br>"
            "Results show matching documents and context snippets."
        )
        intro.setObjectName("muted")
        layout.addWidget(intro)

        row = QHBoxLayout()
        self.query = QLineEdit()
        self.query.setPlaceholderText("Search a research term...")
        search_btn = QPushButton("Search")
        search_btn.clicked.connect(self.run_search)
        self.query.returnPressed.connect(self.run_search)
        row.addWidget(self.query)
        row.addWidget(search_btn)
        layout.addLayout(row)

        self.results = QListWidget()
        self.results.itemDoubleClicked.connect(self.open_result)
        layout.addWidget(self.results, 1)

        self.summary = QLabel("")
        self.summary.setObjectName("muted")
        layout.addWidget(self.summary)

        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn)

        self.result_ids = {}

    def run_search(self):
        query = self.query.text().strip()
        self.results.clear()
        self.result_ids.clear()

        if not query:
            self.summary.setText("Enter a search term.")
            return

        results = self.db.cross_document_search(query)

        for result in results:
            title = (
                f"{result['filename']}  •  "
                f"{result['matches']} text match(es)  •  "
                f"{result['category']}"
            )
            item = QListWidgetItem(title)
            item.setToolTip(
                "\n\n".join(result["snippets"]) or
                "Match found in filename, tags or notes."
            )
            self.results.addItem(item)
            self.result_ids[id(item)] = result["id"]

        self.summary.setText(
            f"{len(results)} document(s) matched “{query}”. "
            "Double-click a result to open it."
        )

    def open_result(self, item):
        doc_id = self.result_ids.get(id(item))
        if doc_id:
            self.parent_window.select_document_by_id(doc_id)
            self.accept()



class CollectionDialog(QDialog):
    def __init__(self, parent, title="Create Collection", name="", description=""):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setMinimumWidth(480)

        form = QFormLayout(self)
        self.name = QLineEdit(name)
        self.description = QPlainTextEdit()
        self.description.setPlainText(description)
        self.description.setMaximumHeight(110)

        form.addRow("Collection name", self.name)
        form.addRow("Description", self.description)

        buttons = QHBoxLayout()
        ok = QPushButton("Save")
        cancel = QPushButton("Cancel")
        ok.clicked.connect(self.accept)
        cancel.clicked.connect(self.reject)
        buttons.addWidget(ok)
        buttons.addWidget(cancel)
        form.addRow(buttons)

    def values(self):
        return self.name.text().strip(), self.description.toPlainText().strip()


class CollectionManagerDialog(QDialog):
    def __init__(self, parent, db, current_document_id=None):
        super().__init__(parent)
        self.setWindowTitle("JASS Research Collections")
        self.resize(760, 520)
        self.db = db
        self.document_id = current_document_id

        layout = QVBoxLayout(self)
        self.list = QListWidget()
        layout.addWidget(self.list, 1)

        buttons = QHBoxLayout()
        self.new_btn = QPushButton("＋ New Collection")
        self.add_btn = QPushButton("Add Current Document")
        self.remove_btn = QPushButton("Remove Current Document")
        self.delete_btn = QPushButton("Delete Collection")
        close_btn = QPushButton("Close")

        self.new_btn.clicked.connect(self.new_collection)
        self.add_btn.clicked.connect(self.add_current)
        self.remove_btn.clicked.connect(self.remove_current)
        self.delete_btn.clicked.connect(self.delete_collection)
        close_btn.clicked.connect(self.accept)

        for b in (self.new_btn, self.add_btn, self.remove_btn, self.delete_btn, close_btn):
            buttons.addWidget(b)
        layout.addLayout(buttons)

        self.refresh()

    def refresh(self):
        self.list.clear()
        for c in self.db.collections():
            count = len(self.db.collection_documents(c["id"]))
            item = QListWidgetItem(f"{c['name']}  •  {count} document(s)")
            item.setData(Qt.UserRole, c["id"])
            item.setToolTip(c["description"] or "No description")
            self.list.addItem(item)

    def selected_id(self):
        item = self.list.currentItem()
        return item.data(Qt.UserRole) if item else None

    def new_collection(self):
        dlg = CollectionDialog(self)
        if dlg.exec():
            name, description = dlg.values()
            if not name:
                return
            try:
                self.db.add_collection(name, description)
                self.refresh()
            except sqlite3.IntegrityError:
                QMessageBox.warning(self, "Collection exists",
                                    "A collection with that name already exists.")

    def add_current(self):
        collection_id = self.selected_id()
        if not collection_id or not self.document_id:
            QMessageBox.information(
                self, "Select collection",
                "Select a collection and a current document first."
            )
            return
        self.db.add_to_collection(collection_id, self.document_id)
        self.refresh()

    def remove_current(self):
        collection_id = self.selected_id()
        if not collection_id or not self.document_id:
            return
        self.db.remove_from_collection(collection_id, self.document_id)
        self.refresh()

    def delete_collection(self):
        collection_id = self.selected_id()
        if not collection_id:
            return
        answer = QMessageBox.question(
            self, "Delete Collection",
            "Delete this collection?\n\nDocuments themselves will not be deleted."
        )
        if answer == QMessageBox.Yes:
            self.db.delete_collection(collection_id)
            self.refresh()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle(f"{APP_NAME} v{VERSION}")
        self.resize(1600, 920)

        base = Path.home() / "JASS" / "DocumentIntelligenceStudio"
        base.mkdir(parents=True, exist_ok=True)
        self.db = Database(base / "jass_documents.db")

        self.current_id = None
        self.current_pdf = None
        self.current_page = 0
        self.current_images = []

        self.build_ui()
        self.apply_style()
        self.refresh_collections()
        self.refresh_library()

        # Show the first existing document immediately.
        if self.library.count():
            self.library.setCurrentRow(0)

    def build_ui(self):
        # ----- Toolbar -----
        toolbar = QToolBar("Main")
        toolbar.setMovable(False)
        self.addToolBar(toolbar)

        import_btn = QPushButton("＋ Import")
        import_btn.clicked.connect(self.import_documents)
        toolbar.addWidget(import_btn)

        refresh_btn = QPushButton("↻ Refresh")
        refresh_btn.clicked.connect(self.refresh_library)
        toolbar.addWidget(refresh_btn)

        collections_btn = QPushButton("Collections")
        collections_btn.clicked.connect(self.open_collections)
        toolbar.addWidget(collections_btn)

        research_btn = QPushButton("Research Search")
        research_btn.clicked.connect(self.open_research_search)
        toolbar.addWidget(research_btn)

        favorites_btn = QPushButton("★ Favorites")
        favorites_btn.clicked.connect(self.show_favorites)
        toolbar.addWidget(favorites_btn)

        toolbar.addSeparator()

        self.search = QLineEdit()
        self.search.setPlaceholderText(
            "Search filenames, document text, tags and notes..."
        )
        self.search.setMinimumWidth(430)
        self.search.textChanged.connect(self.refresh_library)
        toolbar.addWidget(self.search)

        # ----- Left library -----
        left = QWidget()
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(8, 8, 8, 8)

        title = QLabel("DOCUMENT LIBRARY")
        title.setObjectName("sectionTitle")
        left_layout.addWidget(title)

        self.category = QComboBox()
        self.category.addItem("All categories")
        self.category.addItems(CATEGORIES)
        self.category.currentTextChanged.connect(self.refresh_library)
        left_layout.addWidget(self.category)

        self.collection = QComboBox()
        self.collection.addItem("All collections", None)
        self.collection.currentIndexChanged.connect(self.refresh_library)
        left_layout.addWidget(self.collection)

        self.collection_info = QLabel("")
        self.collection_info.setObjectName("muted")
        self.collection_info.setWordWrap(True)
        left_layout.addWidget(self.collection_info)

        self.library_count = QLabel("0 documents")
        self.library_count.setObjectName("muted")
        left_layout.addWidget(self.library_count)

        self.library = QListWidget()
        self.library.setSelectionMode(QAbstractItemView.SingleSelection)
        self.library.currentItemChanged.connect(self.select_document)
        left_layout.addWidget(self.library)

        # ----- Center -----
        center = QWidget()
        center_layout = QVBoxLayout(center)
        center_layout.setContentsMargins(8, 8, 8, 8)

        self.title = QLabel("Select a document")
        self.title.setObjectName("docTitle")
        self.title.setWordWrap(True)
        center_layout.addWidget(self.title)

        self.meta = QLabel("")
        self.meta.setObjectName("muted")
        self.meta.setWordWrap(True)
        center_layout.addWidget(self.meta)

        self.tabs = QTabWidget()
        center_layout.addWidget(self.tabs, 1)

        # Document text
        text_page = QWidget()
        text_layout = QVBoxLayout(text_page)
        text_layout.setContentsMargins(0, 0, 0, 0)

        find_row = QHBoxLayout()
        self.find_box = QLineEdit()
        self.find_box.setPlaceholderText("Find text in this document...")
        self.find_box.textChanged.connect(self.find_text)
        self.find_result = QLabel("")
        find_row.addWidget(self.find_box)
        find_row.addWidget(self.find_result)
        text_layout.addLayout(find_row)

        self.text_view = QTextBrowser()
        text_layout.addWidget(self.text_view)
        self.tabs.addTab(text_page, "Text")

        # PDF visual viewer
        pdf_page = QWidget()
        pdf_layout = QHBoxLayout(pdf_page)
        pdf_layout.setContentsMargins(0, 0, 0, 0)

        self.page_list = QListWidget()
        self.page_list.setIconSize(QSize(130, 170))
        self.page_list.setFixedWidth(165)
        self.page_list.currentRowChanged.connect(self.show_pdf_page)
        pdf_layout.addWidget(self.page_list)

        pdf_view_area = QScrollArea()
        pdf_view_area.setWidgetResizable(True)

        self.pdf_label = QLabel("Select a PDF document")
        self.pdf_label.setAlignment(Qt.AlignCenter)
        self.pdf_label.setTextInteractionFlags(Qt.TextSelectableByMouse)

        pdf_view_area.setWidget(self.pdf_label)
        pdf_layout.addWidget(pdf_view_area, 1)
        self.tabs.addTab(pdf_page, "PDF Pages")

        # Images
        image_page = QWidget()
        image_layout = QVBoxLayout(image_page)

        self.image_gallery = QListWidget()
        self.image_gallery.setViewMode(QListWidget.IconMode)
        self.image_gallery.setIconSize(QSize(180, 130))
        self.image_gallery.setResizeMode(QListWidget.Adjust)
        self.image_gallery.setSpacing(12)
        self.image_gallery.itemDoubleClicked.connect(self.open_image_preview)
        image_layout.addWidget(self.image_gallery)

        self.image_info = QLabel("")
        self.image_info.setObjectName("muted")
        image_layout.addWidget(self.image_info)

        self.tabs.addTab(image_page, "Images")

        # Statistics
        self.stats_view = QTextBrowser()
        self.tabs.addTab(self.stats_view, "Statistics")

        # ----- Right inspector -----
        right = QWidget()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(8, 8, 8, 8)

        inspector = QLabel("DOCUMENT INSPECTOR")
        inspector.setObjectName("sectionTitle")
        right_layout.addWidget(inspector)

        self.type_label = QLabel("No document selected")
        self.type_label.setObjectName("muted")
        self.type_label.setWordWrap(True)
        right_layout.addWidget(self.type_label)

        right_layout.addWidget(QLabel("Category"))
        self.category_edit = QComboBox()
        self.category_edit.addItems(CATEGORIES)
        right_layout.addWidget(self.category_edit)

        right_layout.addWidget(QLabel("Tags"))
        self.tags = QPlainTextEdit()
        self.tags.setPlaceholderText("research, python, athena...")
        self.tags.setMaximumHeight(85)
        right_layout.addWidget(self.tags)

        right_layout.addWidget(QLabel("Notes"))
        self.notes = QPlainTextEdit()
        self.notes.setPlaceholderText("Research notes...")
        right_layout.addWidget(self.notes, 1)

        save_btn = QPushButton("Save Metadata")
        save_btn.clicked.connect(self.save_metadata)
        right_layout.addWidget(save_btn)

        relation_btn = QPushButton("＋ Add Relationship")
        relation_btn.clicked.connect(self.add_relationship)
        right_layout.addWidget(relation_btn)

        self.favorite_btn = QPushButton("☆ Add to Favorites")
        self.favorite_btn.clicked.connect(self.toggle_favorite)
        right_layout.addWidget(self.favorite_btn)

        collection_btn = QPushButton("＋ Add to Current Collection")
        collection_btn.clicked.connect(self.add_to_current_collection)
        right_layout.addWidget(collection_btn)

        remove_collection_btn = QPushButton("− Remove from Current Collection")
        remove_collection_btn.clicked.connect(self.remove_from_current_collection)
        right_layout.addWidget(remove_collection_btn)

        right_layout.addWidget(QLabel("Related Documents"))
        self.related = QListWidget()
        right_layout.addWidget(self.related)

        export_btn = QPushButton("Export Text")
        export_btn.clicked.connect(self.export_text)
        right_layout.addWidget(export_btn)

        delete_btn = QPushButton("Remove from Library")
        delete_btn.clicked.connect(self.delete_document)
        right_layout.addWidget(delete_btn)

        # ----- Main splitter -----
        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(left)
        splitter.addWidget(center)
        splitter.addWidget(right)
        splitter.setSizes([300, 900, 340])
        self.setCentralWidget(splitter)

        self.setStatusBar(QStatusBar())

        # Menu
        file_menu = self.menuBar().addMenu("&File")
        import_action = QAction("Import Documents", self)
        import_action.triggered.connect(self.import_documents)
        file_menu.addAction(import_action)

        export_action = QAction("Export Current Text", self)
        export_action.triggered.connect(self.export_text)
        file_menu.addAction(export_action)

        file_menu.addSeparator()
        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

    def apply_style(self):
        self.setStyleSheet("""
        QMainWindow, QWidget {
            background: #080c16;
            color: #e8edf7;
            font-family: "Segoe UI";
            font-size: 13px;
        }
        QMenuBar, QMenu, QToolBar {
            background: #0d1322;
            color: #e8edf7;
            border: 0;
        }
        QToolBar { padding: 7px; spacing: 7px; }
        QLineEdit, QPlainTextEdit, QTextBrowser, QListWidget,
        QComboBox {
            background: #0d1423;
            border: 1px solid #27334a;
            border-radius: 8px;
            padding: 7px;
            color: #e8edf7;
        }
        QPushButton {
            background: #18243a;
            border: 1px solid #334764;
            border-radius: 8px;
            padding: 8px 13px;
        }
        QPushButton:hover { background: #223451; }
        QListWidget::item {
            padding: 10px;
            border-radius: 6px;
            color: #e8edf7;
        }
        QListWidget::item:selected {
            background: #263b5e;
            color: #ffffff;
        }
        QLabel#sectionTitle {
            font-size: 12px;
            font-weight: bold;
            color: #d9e4f7;
        }
        QLabel#docTitle {
            font-size: 22px;
            font-weight: bold;
            color: #ffffff;
        }
        QLabel#muted { color: #91a0b8; }
        QTabWidget::pane {
            border: 1px solid #27334a;
            border-radius: 8px;
        }
        QTabBar::tab {
            background: #111a2b;
            padding: 9px 16px;
        }
        QTabBar::tab:selected { background: #263b5e; }
        QScrollArea {
            background: #080c16;
            border: 1px solid #27334a;
        }
        """)

    def refresh_collections(self):
        current = self.collection.currentData() if hasattr(self, "collection") else None

        self.collection.blockSignals(True)
        self.collection.clear()
        self.collection.addItem("All collections", None)

        for c in self.db.collections():
            count = len(self.db.collection_documents(c["id"]))
            self.collection.addItem(
                f"{c['name']}  ({count})",
                c["id"]
            )

        if current is not None:
            for i in range(self.collection.count()):
                if self.collection.itemData(i) == current:
                    self.collection.setCurrentIndex(i)
                    break

        self.collection.blockSignals(False)
        self.update_collection_info()

    def update_collection_info(self):
        if not hasattr(self, "collection"):
            return

        collection_id = self.collection.currentData()
        if not collection_id:
            self.collection_info.setText(
                "Showing all documents"
            )
            return

        c = self.db.get_collection(collection_id)
        if not c:
            self.collection_info.setText("")
            return

        count = len(self.db.collection_documents(collection_id))
        description = c["description"] or "No description"
        self.collection_info.setText(
            f"<b>{c['name']}</b><br>"
            f"{count} document(s)<br>"
            f"{description}"
        )

    def refresh_library(self):
        wanted = self.current_id
        docs = self.db.documents(
            self.search.text(),
            self.category.currentText()
        )

        collection_id = self.collection.currentData()
        if collection_id:
            allowed = self.db.collection_document_ids(collection_id)
            docs = [d for d in docs if d["id"] in allowed]

        self.update_collection_info()

        self.library.blockSignals(True)
        self.library.clear()

        for d in docs:
            item = QListWidgetItem(
                f"{d['filename']}\n"
                f"{d['category']}  •  {d['extension'].upper()}  •  "
                f"{d['pages']} page(s)"
            )
            item.setData(Qt.UserRole, d["id"])
            self.library.addItem(item)

        self.library.blockSignals(False)

        self.library_count.setText(
            f"{len(docs)} document{'s' if len(docs) != 1 else ''}"
        )

        selected_row = -1
        for row in range(self.library.count()):
            if self.library.item(row).data(Qt.UserRole) == wanted:
                selected_row = row
                break

        if selected_row >= 0:
            self.library.setCurrentRow(selected_row)
        elif self.library.count():
            self.library.setCurrentRow(0)
        else:
            self.clear_view()

        self.update_status()

    def update_status(self):
        s = self.db.stats()
        mb = s["bytes"] / (1024 * 1024) if s["bytes"] else 0
        collection_id = self.collection.currentData() if hasattr(self, "collection") else None
        if collection_id:
            visible_count = len(self.db.collection_documents(collection_id))
            scope = f"{visible_count} in selected collection"
        else:
            scope = f"{s['documents']} total"

        self.statusBar().showMessage(
            f"{scope}  •  {mb:.2f} MB  •  "
            f"{s['pages']} pages  •  "
            f"{s['images']} extracted images"
        )

    def import_documents(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Import Files",
            str(Path.home()),
            "All Files (*)"
        )
        if not files:
            return

        imported_ids = []
        errors = []

        for filename in files:
            path = Path(filename)
            try:
                checksum = sha1_file(path)
                duplicate = self.db.conn.execute(
                    "SELECT id FROM documents WHERE checksum=?",
                    (checksum,)
                ).fetchone()

                if duplicate:
                    imported_ids.append(duplicate["id"])
                    continue

                text, pages, images = extract_document(path)
                category = self.guess_category(path, text)

                doc_id = self.db.add({
                    "path": str(path.resolve()),
                    "filename": path.name,
                    "extension": path.suffix.lower(),
                    "category": category,
                    "text": text,
                    "size": path.stat().st_size,
                    "pages": pages,
                    "image_count": len(images),
                    "checksum": checksum,
                })

                if images:
                    self.db.add_images(doc_id, images)

                imported_ids.append(doc_id)

            except Exception as exc:
                errors.append(f"{path.name}: {exc}")

        if imported_ids:
            self.current_id = imported_ids[-1]

        self.refresh_library()

        if errors:
            QMessageBox.warning(
                self,
                "Import completed with errors",
                "Some files could not be imported:\n\n" +
                "\n".join(errors)
            )

        self.statusBar().showMessage(
            f"Import complete • {len(imported_ids)} file(s) processed • library contains "
            f"{self.db.stats()['documents']} item(s)"
        )

    def guess_category(self, path, text):
        sample = (path.name + " " + text[:5000]).lower()

        rules = [
            ("Research", ["research", "paper", "study", "journal"]),
            ("Language", ["punjabi", "mizo", "language", "gurbani"]),
            ("Technical", ["python", "sqlite", "database", "api", "software"]),
            ("Project", ["project athena", "jass digital lab"]),
            ("Literature", ["poem", "poetry", "novel", "literature"]),
            ("Data", ["dataset", "csv", "parquet", "corpus"]),
            ("Creative", ["image", "photo", "picture", "video", "audio", "music"]),
        ]

        for category, words in rules:
            if any(word in sample for word in words):
                return category

        return "Other"

    def select_document(self, current, previous=None):
        if not current:
            return

        doc_id = current.data(Qt.UserRole)
        doc = self.db.get(doc_id)
        if not doc:
            return

        self.current_id = doc_id

        self.title.setText(doc["filename"])
        ext = doc["extension"].lower()
        kind = SUPPORTED.get(ext, "Unknown / Binary")
        self.type_label.setText(
            f"{kind}  •  {doc['extension'].upper().replace('.', '') or 'NO EXTENSION'}  •  "
            f"{doc['size'] / 1024:.1f} KB\n"
            f"{doc['pages']} page(s)  •  {doc['image_count']} image(s)"
        )

        memberships = self.db.collection_for_document(doc["id"])
        membership_text = (
            ", ".join(c["name"] for c in memberships)
            if memberships else "No collection"
        )
        self.meta.setText(
            f"{doc['path']}\nCollections: {membership_text}"
        )

        self.category_edit.blockSignals(True)
        self.category_edit.setCurrentText(doc["category"] or "Other")
        self.category_edit.blockSignals(False)

        self.tags.setPlainText(doc["tags"] or "")
        self.notes.setPlainText(doc["notes"] or "")

        self.text_view.setPlainText(doc["text"] or "")
        self.find_box.clear()

        self.update_favorite_button()
        self.load_pdf(doc)
        self.load_images(doc)
        self.load_stats(doc)
        self.load_relations(doc)

    def load_pdf(self, doc):
        self.page_list.clear()
        self.current_pdf = None
        self.current_page = 0

        if doc["extension"].lower() != ".pdf" or pymupdf is None:
            self.page_list.addItem("PDF preview unavailable")
            self.pdf_label.setText(
                "This file is not a PDF.\n\n"
                "Use the Text tab for extracted text, or the Images tab for image files."
            )
            return

        path = Path(doc["path"])
        if not path.exists():
            self.pdf_label.setText("Original PDF file is no longer available.")
            return

        try:
            self.current_pdf = pymupdf.open(str(path))

            for page_no in range(len(self.current_pdf)):
                page = self.current_pdf[page_no]
                pix = page.get_pixmap(
                    matrix=pymupdf.Matrix(0.65, 0.65),
                    alpha=False
                )
                image = QImage(
                    pix.samples, pix.width, pix.height,
                    pix.stride, QImage.Format_RGB888
                ).copy()

                item = QListWidgetItem(f"Page {page_no + 1}")
                item.setIcon(
                    QPixmap.fromImage(image).scaled(
                        125, 165,
                        Qt.KeepAspectRatio,
                        Qt.SmoothTransformation
                    )
                )
                self.page_list.addItem(item)

            if self.page_list.count():
                self.page_list.setCurrentRow(0)

        except Exception as exc:
            self.pdf_label.setText(f"Could not render PDF:\n{exc}")

    def show_pdf_page(self, row):
        if self.current_pdf is None or row < 0:
            return

        try:
            page = self.current_pdf[row]
            pix = page.get_pixmap(matrix=pymupdf.Matrix(1.45, 1.45), alpha=False)
            image = QImage(
                pix.samples, pix.width, pix.height,
                pix.stride, QImage.Format_RGB888
            ).copy()

            self.pdf_label.setPixmap(QPixmap.fromImage(image))
            self.pdf_label.setText("")
            self.current_page = row
            self.tabs.setCurrentIndex(1)

        except Exception as exc:
            self.pdf_label.setText(f"Could not render page:\n{exc}")

    def load_images(self, doc):
        self.image_gallery.clear()
        self.current_images = self.db.images(doc["id"])

        for index, row in enumerate(self.current_images, 1):
            pix = QPixmap()
            pix.loadFromData(row["image"])

            if pix.isNull():
                continue

            item = QListWidgetItem()
            item.setIcon(
                pix.scaled(
                    180, 130,
                    Qt.KeepAspectRatio,
                    Qt.SmoothTransformation
                )
            )
            if row["page"] == 0 and doc["extension"].lower() in IMAGE_EXTENSIONS:
                item.setText(f"Image {index} • Imported file")
                item.setToolTip("Imported image file")
            else:
                item.setText(f"Image {index} • Page {row['page']}")
                item.setToolTip(f"Extracted from page {row['page']}")
            self.image_gallery.addItem(item)

        self.image_info.setText(
            f"{len(self.current_images)} extracted image(s)"
        )

    def open_image_preview(self, item):
        row = self.image_gallery.row(item)
        if row < 0 or row >= len(self.current_images):
            return
        pix = QPixmap()
        pix.loadFromData(self.current_images[row]["image"])
        if pix.isNull():
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("Image Preview")
        dialog.resize(1000, 760)
        layout = QVBoxLayout(dialog)
        area = QScrollArea()
        area.setWidgetResizable(True)
        label = QLabel()
        label.setAlignment(Qt.AlignCenter)
        label.setPixmap(pix)
        area.setWidget(label)
        layout.addWidget(area)
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(dialog.accept)
        layout.addWidget(close_btn)
        dialog.exec()

    def load_stats(self, doc):
        words = len(re.findall(r"\b\w+\b", doc["text"] or ""))
        chars = len(doc["text"] or "")

        collections = [
            c["name"] for c in self.db.collections()
            if any(d["id"] == doc["id"] for d in self.db.collection_documents(c["id"]))
        ]
        collection_text = ", ".join(collections) if collections else "None"

        self.stats_view.setPlainText(
            f"Filename: {doc['filename']}\n"
            f"Type: {SUPPORTED.get(doc['extension'].lower(), 'Unknown / Binary')} ({doc['extension'].upper() or 'no extension'})\n"
            f"MIME: {mimetypes.guess_type(doc['filename'])[0] or 'unknown'}\n"
            f"Category: {doc['category']}\n"
            f"Collections: {collection_text}\n"
            f"Size: {doc['size']:,} bytes\n"
            f"Pages: {doc['pages']:,}\n"
            f"Words: {words:,}\n"
            f"Characters: {chars:,}\n"
            f"Extracted images: {doc['image_count']:,}\n"
            f"Imported: {doc['imported_at']}\n"
            f"Updated: {doc['modified_at']}\n"
            f"SHA-1: {doc['checksum']}\n"
            f"Path: {doc['path']}"
        )

    def find_text(self, query):
        self.text_view.setExtraSelections([])

        if not query:
            self.find_result.setText("")
            return

        selections = []
        cursor = self.text_view.document().find(query)
        count = 0

        while not cursor.isNull():
            fmt = QTextCharFormat()
            fmt.setBackground(QColor("#7a5a00"))
            fmt.setForeground(QColor("#ffffff"))

            selection = type("Selection", (), {})()
            selection.cursor = cursor
            selection.format = fmt
            selections.append(selection)

            count += 1
            cursor = self.text_view.document().find(query, cursor)

        self.text_view.setExtraSelections(selections)
        self.find_result.setText(
            f"{count} match{'es' if count != 1 else ''}"
        )

    def load_relations(self, doc):
        self.related.clear()
        for row in self.db.relations(doc["id"]):
            item = QListWidgetItem(
                f"{row['relation']}  →  {row['filename']}"
            )
            item.setData(Qt.UserRole, row["id"])
            self.related.addItem(item)

    def select_document_by_id(self, doc_id):
        for row in range(self.library.count()):
            item = self.library.item(row)
            if item.data(Qt.UserRole) == doc_id:
                self.library.setCurrentItem(item)
                return

        # If current filters hide it, clear filters and select it.
        self.search.blockSignals(True)
        self.search.clear()
        self.search.blockSignals(False)
        self.category.blockSignals(True)
        self.category.setCurrentIndex(0)
        self.category.blockSignals(False)
        self.collection.blockSignals(True)
        self.collection.setCurrentIndex(0)
        self.collection.blockSignals(False)
        self.refresh_library()

        for row in range(self.library.count()):
            item = self.library.item(row)
            if item.data(Qt.UserRole) == doc_id:
                self.library.setCurrentItem(item)
                return

    def update_favorite_button(self):
        if not hasattr(self, "favorite_btn"):
            return
        if not self.current_id:
            self.favorite_btn.setText("☆ Add to Favorites")
            return
        if self.db.is_favorite(self.current_id):
            self.favorite_btn.setText("★ Remove from Favorites")
        else:
            self.favorite_btn.setText("☆ Add to Favorites")

    def toggle_favorite(self):
        if not self.current_id:
            return
        enabled = not self.db.is_favorite(self.current_id)
        self.db.set_favorite(self.current_id, enabled)
        self.update_favorite_button()
        self.statusBar().showMessage(
            "Added to Favorites." if enabled else "Removed from Favorites."
        )

    def show_favorites(self):
        docs = self.db.favorite_documents()
        self.library.blockSignals(True)
        self.library.clear()

        for d in docs:
            item = QListWidgetItem(
                f"★ {d['filename']}\n"
                f"{d['category']}  •  {d['pages']} page(s)"
            )
            item.setData(Qt.UserRole, d["id"])
            self.library.addItem(item)

        self.library.blockSignals(False)
        self.library_count.setText(
            f"{len(docs)} favorite document(s)"
        )

        if docs:
            self.library.setCurrentRow(0)
        else:
            self.clear_view()

        self.statusBar().showMessage(
            f"Favorites: {len(docs)} document(s)"
        )

    def open_research_search(self):
        dlg = ResearchSearchDialog(self, self.db)
        dlg.exec()

    def add_to_current_collection(self):
        if not self.current_id:
            QMessageBox.information(
                self, "Select a document",
                "Select a document first."
            )
            return

        collection_id = self.collection.currentData()
        if not collection_id:
            QMessageBox.information(
                self, "Select a collection",
                "Choose a collection from the left panel first."
            )
            return

        self.db.add_to_collection(collection_id, self.current_id)
        self.refresh_collections()
        self.refresh_library()
        self.load_stats(self.db.get(self.current_id))
        self.statusBar().showMessage("Document added to collection.")

    def remove_from_current_collection(self):
        if not self.current_id:
            return

        collection_id = self.collection.currentData()
        if not collection_id:
            QMessageBox.information(
                self, "Select a collection",
                "Choose a collection from the left panel first."
            )
            return

        self.db.remove_from_collection(collection_id, self.current_id)
        self.refresh_collections()
        self.refresh_library()
        self.statusBar().showMessage("Document removed from collection.")

    def open_collections(self):
        dlg = CollectionManagerDialog(
            self, self.db, self.current_id
        )
        dlg.exec()
        self.refresh_collections()
        self.refresh_library()
        if self.current_id:
            self.load_stats(self.db.get(self.current_id))

    def save_metadata(self):
        if not self.current_id:
            return

        self.db.update(
            self.current_id,
            category=self.category_edit.currentText(),
            tags=self.tags.toPlainText().strip(),
            notes=self.notes.toPlainText().strip()
        )
        self.refresh_library()
        self.statusBar().showMessage("Metadata saved.")

    def add_relationship(self):
        if not self.current_id:
            QMessageBox.information(
                self, "Select a document",
                "Select a document before adding a relationship."
            )
            return

        docs = self.db.documents()
        if len(docs) < 2:
            QMessageBox.information(
                self, "More documents required",
                "Import at least two documents first."
            )
            return

        dialog = RelationDialog(self, docs, self.current_id)
        if dialog.exec():
            target, relation = dialog.values()
            if target:
                self.db.add_relation(self.current_id, target, relation)
                self.load_relations(self.db.get(self.current_id))

    def export_text(self):
        if not self.current_id:
            QMessageBox.information(
                self, "Select a document", "Select a document first."
            )
            return

        doc = self.db.get(self.current_id)
        if not doc:
            return

        filename, _ = QFileDialog.getSaveFileName(
            self,
            "Export Text",
            str(Path.home() / f"{Path(doc['filename']).stem}_export.txt"),
            "Text Files (*.txt)"
        )
        if not filename:
            return

        Path(filename).write_text(
            doc["text"] or "",
            encoding="utf-8"
        )
        self.statusBar().showMessage(f"Exported: {filename}")

    def delete_document(self):
        if not self.current_id:
            return

        doc = self.db.get(self.current_id)
        if not doc:
            return

        answer = QMessageBox.question(
            self,
            "Remove from Library",
            f"Remove '{doc['filename']}' from the library?\n\n"
            "The original file will NOT be deleted."
        )
        if answer != QMessageBox.Yes:
            return

        self.db.remove(self.current_id)
        self.current_id = None
        self.refresh_library()

    def clear_view(self):
        self.current_id = None
        self.title.setText("Document Library")
        self.meta.setText(
            "Import any file to begin. Supported formats get specialized extraction or preview; unknown formats are preserved as file records."
        )
        self.type_label.setText("No document selected")
        self.text_view.clear()
        self.pdf_label.setText(
            "Select a PDF for page preview, or an image for visual preview in the Images tab."
        )
        self.page_list.clear()
        self.image_gallery.clear()
        self.image_info.setText("")
        self.stats_view.clear()
        self.related.clear()
        self.tags.clear()
        self.notes.clear()
        self.find_box.clear()
        self.update_favorite_button()

    def closeEvent(self, event):
        if self.current_pdf:
            try:
                self.current_pdf.close()
            except Exception:
                pass
        self.db.close()
        event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setOrganizationName("JASS Digital Lab")

    window = MainWindow()
    window.show()
    sys.exit(app.exec())
