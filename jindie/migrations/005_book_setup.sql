CREATE TABLE book_setup_versions (
    book_id TEXT NOT NULL REFERENCES books(id),
    version INTEGER NOT NULL CHECK (version > 0),
    config_json TEXT NOT NULL,
    template BLOB NOT NULL,
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    PRIMARY KEY (book_id, version)
);
