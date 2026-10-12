"""Manage SQLite document storage, exact filters, and FTS5 retrieval."""

import json
import os
import re
import sqlite3

from backend import core

KNOWLEDGE_DB_PATH = os.path.join("knowledge", "knowledge.sqlite3")
FULL_DATASET_ROW_LIMIT = 300
SMALL_DOCUMENT_CHUNK_LIMIT = 20
SMALL_DOCUMENT_CHARACTER_LIMIT = 20000


def _connect():
    """Open a connection to the SQLite knowledge index."""

    connection = sqlite3.connect(KNOWLEDGE_DB_PATH, timeout=30)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    return connection


def init_knowledge_db():
    """Create relational document, attribute, and FTS5 search tables."""

    os.makedirs(os.path.dirname(KNOWLEDGE_DB_PATH), exist_ok=True)
    connection = _connect()
    connection.executescript(
        """
        CREATE TABLE IF NOT EXISTS knowledge_documents (
            id TEXT PRIMARY KEY,
            context_id TEXT NOT NULL,
            document_id TEXT NOT NULL,
            filename TEXT NOT NULL,
            file_type TEXT NOT NULL,
            chunk_index INTEGER NOT NULL,
            record_type TEXT NOT NULL,
            row_index INTEGER NOT NULL,
            dataset_type TEXT NOT NULL DEFAULT '',
            column_names TEXT NOT NULL DEFAULT '',
            content TEXT NOT NULL,
            metadata_json TEXT NOT NULL
        );
        CREATE INDEX IF NOT EXISTS idx_knowledge_context_document
            ON knowledge_documents(context_id, document_id);
        CREATE INDEX IF NOT EXISTS idx_knowledge_context_file
            ON knowledge_documents(context_id, filename);
        CREATE TABLE IF NOT EXISTS knowledge_attributes (
            document_row_id TEXT NOT NULL,
            attribute_name TEXT NOT NULL,
            attribute_value TEXT NOT NULL COLLATE NOCASE,
            FOREIGN KEY(document_row_id) REFERENCES knowledge_documents(id)
                ON DELETE CASCADE
        );
        CREATE INDEX IF NOT EXISTS idx_knowledge_attribute_lookup
            ON knowledge_attributes(attribute_name, attribute_value);
        CREATE VIRTUAL TABLE IF NOT EXISTS knowledge_fts USING fts5(
            document_row_id UNINDEXED,
            content,
            tokenize='unicode61 remove_diacritics 2'
        );
        """
    )
    connection.commit()
    connection.close()


init_knowledge_db()


def _format_dataset_summaries(metadata_rows):
    """Summarize tabular schemas and column values across all indexed rows."""

    datasets = {}
    for metadata in metadata_rows:
        if metadata.get("record_type") != "row":
            continue

        filename = metadata.get("filename", "unknown")
        dataset = datasets.setdefault(
            filename,
            {
                "dataset_type": metadata.get("dataset_type", "tabular"),
                "column_names": metadata.get("column_names", "").split("|"),
                "rows": 0,
                "values": {}
            }
        )
        dataset["rows"] += 1
        for key, value in metadata.items():
            if key.startswith("attr_") and isinstance(value, str) and value.strip():
                dataset["values"].setdefault(key, []).append(value.strip())

    summaries = []
    for filename, dataset in sorted(datasets.items()):
        labels = {
            "attr_" + re.sub(r"[^a-z0-9]+", "_", column.lower()).strip("_"): column
            for column in dataset["column_names"]
        }
        summary_lines = [
            f"[Dataset summary: {filename}]",
            f"Dataset type: {dataset['dataset_type']}",
            f"Rows indexed: {dataset['rows']}",
            "Column values:"
        ]

        for key, raw_values in sorted(dataset["values"].items()):
            counts = {}
            display_values = {}
            for value in raw_values:
                normalized = value.casefold()
                counts[normalized] = counts.get(normalized, 0) + 1
                display_values.setdefault(normalized, value)

            distinct_count = len(counts)
            label = labels.get(key, key[5:].replace("_", " ").title())
            if distinct_count <= 25:
                value_summary = ", ".join(
                    f"{display_values[value]} ({count})"
                    for value, count in sorted(
                        counts.items(),
                        key=lambda item: (-item[1], item[0])
                    )
                )
            else:
                common = sorted(
                    counts.items(),
                    key=lambda item: (-item[1], item[0])
                )[:5]
                value_summary = ", ".join(
                    f"{display_values[value]} ({count})"
                    for value, count in common
                )

            summary_lines.append(
                f"- {label}: {distinct_count} distinct values; {value_summary}"
            )

        summaries.append("\n".join(summary_lines))

    return "\n\n".join(summaries)


def restore_persisted_session():
    """Restore the most recently indexed context from SQLite."""

    if core.SESSION_FILES:
        return

    connection = _connect()
    rows = connection.execute(
        """
        SELECT context_id, filename, MAX(rowid) AS last_indexed
        FROM knowledge_documents
        GROUP BY context_id, filename
        ORDER BY last_indexed DESC
        """
    ).fetchall()
    connection.close()
    if not rows:
        return

    context_id = rows[0]["context_id"]
    filenames = sorted({
        row["filename"]
        for row in rows
        if row["context_id"] == context_id
    })
    existing_files = [
        os.path.join(core.UPLOAD_DIR, filename)
        for filename in filenames
        if os.path.exists(os.path.join(core.UPLOAD_DIR, filename))
    ]

    if existing_files:
        core.ACTIVE_CONTEXT_ID = context_id
        core.SESSION_FILES.extend(existing_files)
        print(
            "RESTORED PERSISTED KNOWLEDGE:",
            {
                "context_id": context_id,
                "files": filenames
            }
        )


def clear_collection():
    """Delete all indexed documents, attributes, and FTS entries."""

    connection = _connect()
    connection.execute("DELETE FROM knowledge_fts")
    connection.execute("DELETE FROM knowledge_attributes")
    connection.execute("DELETE FROM knowledge_documents")
    connection.commit()
    connection.close()


def index_file(
    filename,
    text,
    context_id=None,
    documents=None,
    document_metadata=None
):
    """Store text chunks or normalized rows and attributes in SQLite/FTS5."""

    context_id = context_id or core.ACTIVE_CONTEXT_ID
    if not text or not text.strip():
        return

    chunks = documents if documents is not None else [
        text[index:index + 1000]
        for index in range(0, len(text), 1000)
    ]

    if not chunks:
        return

    file_type = os.path.splitext(filename)[1].lower().lstrip(".") or "unknown"
    is_row_index = documents is not None
    connection = _connect()
    existing_ids = connection.execute(
        "SELECT id FROM knowledge_documents WHERE filename = ?",
        (filename,)
    ).fetchall()
    existing_ids = [row["id"] for row in existing_ids]
    if existing_ids:
        placeholders = ",".join("?" for _ in existing_ids)
        connection.execute(
            f"DELETE FROM knowledge_fts WHERE document_row_id IN ({placeholders})",
            existing_ids
        )
        connection.execute(
            f"DELETE FROM knowledge_attributes WHERE document_row_id IN ({placeholders})",
            existing_ids
        )
        connection.execute(
            f"DELETE FROM knowledge_documents WHERE id IN ({placeholders})",
            existing_ids
        )

    for index, content in enumerate(chunks):
        document_row_id = f"{context_id}:{filename}:{index}"
        metadata = {
            "context_id": context_id,
            "document_id": filename,
            "filename": filename,
            "file_type": file_type,
            "chunk_index": index,
            "record_type": "row" if is_row_index else "chunk",
            "row_index": index if is_row_index else -1
        }
        if document_metadata and index < len(document_metadata):
            metadata.update(document_metadata[index])
        connection.execute(
            """
            INSERT INTO knowledge_documents (
                id, context_id, document_id, filename, file_type, chunk_index,
                record_type, row_index, dataset_type, column_names, content,
                metadata_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                document_row_id,
                context_id,
                filename,
                filename,
                file_type,
                index,
                metadata["record_type"],
                metadata["row_index"],
                metadata.get("dataset_type", ""),
                metadata.get("column_names", ""),
                content,
                json.dumps(metadata, ensure_ascii=False)
            )
        )
        connection.execute(
            "INSERT INTO knowledge_fts(document_row_id, content) VALUES (?, ?)",
            (document_row_id, content)
        )
        for key, value in metadata.items():
            if key.startswith("attr_") and value is not None and str(value).strip():
                connection.execute(
                    """
                    INSERT INTO knowledge_attributes (
                        document_row_id, attribute_name, attribute_value
                    ) VALUES (?, ?, ?)
                    """,
                    (document_row_id, key[5:], str(value).strip())
                )

    connection.commit()
    connection.close()

    print(
        "INDEXED METADATA:",
        {
            "context_id": context_id,
            "filename": filename,
            "document_id": filename,
            "file_type": file_type,
            "chunks": len(chunks),
            "characters": len(text)
        }
    )


def search_context(query, limit=5):

    """Use SQL metadata filters, then FTS5-ranked text retrieval."""

    restore_persisted_session()
    session_documents = [
        os.path.basename(filepath)
        for filepath in core.SESSION_FILES
        if filepath and os.path.exists(filepath)
    ]
    if not session_documents:
        return ""

    placeholders = ",".join("?" for _ in session_documents)
    session_params = [core.ACTIVE_CONTEXT_ID, *session_documents]
    session_where = (
        f"context_id = ? AND document_id IN ({placeholders})"
    )
    connection = _connect()
    records = connection.execute(
        f"SELECT * FROM knowledge_documents WHERE {session_where}",
        session_params
    ).fetchall()
    metadata_rows = [json.loads(row["metadata_json"]) for row in records]
    indexed_documents = [row["content"] for row in records]
    query_text = re.sub(r"[^a-z0-9]+", " ", query.lower()).strip()

    exact_candidates = []
    for metadata in metadata_rows:
        for key, value in metadata.items():
            if not key.startswith("attr_") or not isinstance(value, str):
                continue
            normalized_value = re.sub(r"[^a-z0-9]+", " ", value.lower()).strip()
            if len(normalized_value) < 3 or normalized_value not in query_text:
                continue
            attribute_name = key[5:]
            column_mentioned = re.sub(r"_+", " ", attribute_name) in query_text
            exact_candidates.append((column_mentioned, len(normalized_value), attribute_name, value))

    exact_records = []
    if exact_candidates:
        _, _, attribute_name, attribute_value = max(exact_candidates)
        exact_records = connection.execute(
            f"""
            SELECT d.* FROM knowledge_documents AS d
            JOIN knowledge_attributes AS a ON a.document_row_id = d.id
            WHERE {session_where}
              AND a.attribute_name = ?
              AND a.attribute_value = ? COLLATE NOCASE
            ORDER BY d.filename, d.chunk_index
            """,
            [*session_params, attribute_name, attribute_value]
        ).fetchall()

    row_records = connection.execute(
        f"""SELECT * FROM knowledge_documents
            WHERE {session_where} AND record_type = 'row'
            ORDER BY filename, row_index""",
        session_params
    ).fetchall()
    dataset_summary = _format_dataset_summaries(metadata_rows)
    full_dataset_request = bool(re.search(
        r"\b(full dataset|whole dataset|entire dataset|all rows|every row|"
        r"entire spreadsheet|whole spreadsheet)\b",
        query_text
    ))
    small_non_tabular_document = (
        bool(records)
        and len(records) <= SMALL_DOCUMENT_CHUNK_LIMIT
        and sum(len(row["content"] or "") for row in records)
        <= SMALL_DOCUMENT_CHARACTER_LIMIT
        and all(
            row["record_type"] != "row"
            and row["file_type"] not in {"csv", "xlsx", "xlsm", "xls"}
            for row in records
        )
    )

    if full_dataset_request and 0 < len(row_records) <= FULL_DATASET_ROW_LIMIT:
        selected_records = row_records
    elif small_non_tabular_document:
        selected_records = records
    elif exact_records:
        selected_records = exact_records
    else:
        stop_words = {
            "the", "and", "for", "with", "what", "when", "where", "which",
            "who", "why", "how", "many", "does", "this", "that", "are",
            "was", "were", "can", "could", "would", "should", "from", "have",
            "has", "had", "into", "about", "there", "their", "them", "please"
        }
        query_tokens = [
            token
            for token in dict.fromkeys(re.findall(r"[a-z0-9]+", query_text))
            if len(token) > 1 and token not in stop_words
        ]
        if not query_tokens:
            connection.close()
            return dataset_summary

        fts_query = " OR ".join(f'"{token.replace(chr(34), chr(34) * 2)}"' for token in query_tokens)
        selected_records = connection.execute(
            f"""
            SELECT d.* FROM knowledge_fts AS f
            JOIN knowledge_documents AS d ON d.id = f.document_row_id
            WHERE knowledge_fts MATCH ? AND {session_where}
            ORDER BY bm25(knowledge_fts)
            LIMIT ?
            """,
            [fts_query, *session_params, limit]
        ).fetchall()

    print(
        "SQLITE RETRIEVED SOURCES:",
        [
            {"filename": row["filename"], "chunk_index": row["chunk_index"]}
            for row in selected_records
        ]
    )
    connection.close()

    retrieved_context = "\n\n".join(
        f"[Source: {row['filename']}, chunk {row['chunk_index']}]\n{row['content']}"
        for row in selected_records
    )
    return "\n\n".join(
        section for section in (dataset_summary, retrieved_context) if section
    )


def get_knowledge_statistics():
    """Return file and indexed-record statistics from SQLite."""

    restore_persisted_session()
    connection = _connect()
    rows = connection.execute(
        """
        SELECT filename, file_type, content, record_type
        FROM knowledge_documents
        WHERE context_id = ?
        """,
        (core.ACTIVE_CONTEXT_ID,)
    ).fetchall()
    connection.close()
    files = {}

    for row in rows:
        filename = row["filename"]
        file_data = files.setdefault(
            filename,
            {
                "filename": filename,
                "file_type": row["file_type"],
                "chunks": 0,
                "characters": 0
            }
        )
        file_data["chunks"] += 1
        file_data["characters"] += len(row["content"] or "")

    statistics = {
        "context_id": core.ACTIVE_CONTEXT_ID,
        "files": len(files),
        "chunks": len(rows),
        "characters": sum(item["characters"] for item in files.values()),
        "file_types": sorted({item["file_type"] for item in files.values()}),
        "documents": sorted(files.values(), key=lambda item: item["filename"].lower())
    }
    print("KNOWLEDGE STATISTICS:", statistics)
    return statistics


def get_knowledge_explorer():
    """List indexed SQLite records for the knowledge explorer endpoint."""

    restore_persisted_session()
    connection = _connect()
    rows = connection.execute(
        """
        SELECT filename, file_type, chunk_index, content
        FROM knowledge_documents
        WHERE context_id = ?
        ORDER BY filename, chunk_index
        """,
        (core.ACTIVE_CONTEXT_ID,)
    ).fetchall()
    connection.close()
    explorer_records = []

    for row in rows:
        explorer_records.append({
            "filename": row["filename"],
            "file_type": row["file_type"],
            "chunk_index": row["chunk_index"],
            "characters": len(row["content"] or ""),
            "preview": (row["content"] or "")[:240]
        })

    print("KNOWLEDGE EXPLORER RECORDS:", explorer_records)
    return explorer_records