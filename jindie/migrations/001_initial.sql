CREATE TABLE books (
    id TEXT PRIMARY KEY NOT NULL
);
CREATE TABLE accounts (
    book_id TEXT NOT NULL REFERENCES books(id),
    id TEXT NOT NULL,
    number TEXT NOT NULL,
    PRIMARY KEY (book_id, id),
    UNIQUE (book_id, number)
);
CREATE TABLE bank_profile_versions (
    id TEXT NOT NULL,
    version INTEGER NOT NULL CHECK (version > 0),
    config_json TEXT NOT NULL,
    PRIMARY KEY (id, version)
);
CREATE TABLE import_batches (
    id INTEGER PRIMARY KEY,
    book_id TEXT NOT NULL,
    account_id TEXT NOT NULL,
    period TEXT NOT NULL,
    profile_id TEXT NOT NULL,
    profile_version INTEGER NOT NULL,
    file_hash TEXT NOT NULL,
    filename TEXT NOT NULL,
    sheet TEXT NOT NULL,
    header_row INTEGER NOT NULL CHECK (header_row > 0),
    options_json TEXT NOT NULL,
    warnings_json TEXT NOT NULL,
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    UNIQUE (book_id, account_id, period, profile_id, profile_version, file_hash, sheet, options_json),
    FOREIGN KEY (book_id, account_id) REFERENCES accounts(book_id, id),
    FOREIGN KEY (profile_id, profile_version) REFERENCES bank_profile_versions(id, version)
);
CREATE TABLE bank_transactions (
    id INTEGER PRIMARY KEY,
    book_id TEXT NOT NULL,
    account_id TEXT NOT NULL,
    reference TEXT,
    origin_key TEXT NOT NULL,
    file_hash TEXT NOT NULL,
    content_hash TEXT NOT NULL,
    candidate_hash TEXT NOT NULL,
    data_json TEXT NOT NULL,
    FOREIGN KEY (book_id, account_id) REFERENCES accounts(book_id, id),
    UNIQUE (book_id, account_id, id),
    UNIQUE (book_id, account_id, reference),
    UNIQUE (book_id, account_id, origin_key)
);
CREATE INDEX bank_transaction_candidates ON bank_transactions (book_id, account_id, candidate_hash);
CREATE TABLE source_rows (
    batch_id INTEGER NOT NULL REFERENCES import_batches(id),
    row_number INTEGER NOT NULL CHECK (row_number > 0),
    content_hash TEXT NOT NULL,
    raw_json TEXT NOT NULL,
    parsed_json TEXT,
    issues_json TEXT NOT NULL,
    match_status TEXT NOT NULL CHECK (match_status IN ('new', 'invalid', 'duplicate', 'conflict', 'possible_duplicate')),
    transaction_id INTEGER REFERENCES bank_transactions(id),
    PRIMARY KEY (batch_id, row_number)
);
