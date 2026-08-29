-- BookStats v0.5: user-defined shelves are synchronized as first-class records.
ALTER TABLE library_records
  ADD COLUMN IF NOT EXISTS record_type text NOT NULL DEFAULT 'book';

CREATE INDEX IF NOT EXISTS idx_library_records_type
  ON library_records(user_id, record_type, updated_at);
