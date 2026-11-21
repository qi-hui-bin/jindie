CREATE TABLE review_batches (
    batch_id INTEGER PRIMARY KEY REFERENCES import_batches(id),
    book_config_json TEXT NOT NULL,
    exclusions_json TEXT NOT NULL,
    revision INTEGER NOT NULL DEFAULT 0 CHECK (revision >= 0)
);
CREATE TABLE voucher_drafts (
    id INTEGER PRIMARY KEY,
    batch_id INTEGER NOT NULL REFERENCES review_batches(batch_id),
    payload_json TEXT NOT NULL,
    review_digest TEXT
);
CREATE INDEX voucher_drafts_batch ON voucher_drafts(batch_id, id);
CREATE TABLE voucher_exports (
    id INTEGER PRIMARY KEY,
    batch_id INTEGER NOT NULL REFERENCES review_batches(batch_id),
    request_id TEXT NOT NULL,
    book_id TEXT NOT NULL REFERENCES books(id),
    period TEXT NOT NULL,
    voucher_word TEXT NOT NULL,
    template_hash TEXT NOT NULL,
    options_json TEXT NOT NULL,
    exclusions_json TEXT NOT NULL,
    workbook BLOB NOT NULL,
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    UNIQUE (batch_id, request_id),
    UNIQUE (id, book_id, period, voucher_word)
);
CREATE TABLE exported_vouchers (
    export_id INTEGER NOT NULL,
    draft_id INTEGER NOT NULL UNIQUE REFERENCES voucher_drafts(id),
    book_id TEXT NOT NULL,
    period TEXT NOT NULL,
    voucher_word TEXT NOT NULL,
    number INTEGER NOT NULL CHECK (number > 0),
    snapshot_json TEXT NOT NULL,
    import_status TEXT NOT NULL DEFAULT 'generated' CHECK (import_status IN ('generated', 'imported', 'failed', 'not_imported')),
    import_note TEXT NOT NULL DEFAULT '',
    PRIMARY KEY (export_id, number),
    UNIQUE (book_id, period, voucher_word, number),
    FOREIGN KEY (export_id, book_id, period, voucher_word) REFERENCES voucher_exports(id, book_id, period, voucher_word)
);
CREATE TABLE exported_sources (
    transaction_id INTEGER PRIMARY KEY REFERENCES bank_transactions(id),
    export_id INTEGER NOT NULL,
    number INTEGER NOT NULL,
    FOREIGN KEY (export_id, number) REFERENCES exported_vouchers(export_id, number)
);
