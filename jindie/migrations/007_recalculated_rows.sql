CREATE TABLE recalculated_rows (
    origin_batch_id INTEGER NOT NULL,
    origin_row_number INTEGER NOT NULL,
    target_batch_id INTEGER NOT NULL UNIQUE REFERENCES review_batches(batch_id),
    PRIMARY KEY (origin_batch_id, origin_row_number),
    FOREIGN KEY (origin_batch_id, origin_row_number) REFERENCES source_rows(batch_id, row_number)
);
