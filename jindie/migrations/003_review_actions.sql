ALTER TABLE review_batches ADD COLUMN skips_json TEXT NOT NULL DEFAULT '{}';
ALTER TABLE voucher_drafts ADD COLUMN superseded INTEGER NOT NULL DEFAULT 0 CHECK (superseded IN (0, 1));
CREATE TABLE review_actions (
    id INTEGER PRIMARY KEY,
    batch_id INTEGER NOT NULL REFERENCES review_batches(batch_id),
    kind TEXT NOT NULL,
    details_json TEXT NOT NULL,
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE voucher_retries (
    draft_id INTEGER PRIMARY KEY REFERENCES voucher_drafts(id),
    original_export_id INTEGER NOT NULL,
    original_number INTEGER NOT NULL,
    UNIQUE (original_export_id, original_number),
    FOREIGN KEY (original_export_id, original_number) REFERENCES exported_vouchers(export_id, number)
);
