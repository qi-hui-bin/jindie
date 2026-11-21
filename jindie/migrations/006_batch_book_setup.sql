ALTER TABLE import_batches ADD COLUMN setup_version INTEGER CHECK (setup_version > 0);
UPDATE import_batches
SET setup_version = 1
WHERE book_id IN (SELECT book_id FROM book_setup_versions WHERE version = 1);
