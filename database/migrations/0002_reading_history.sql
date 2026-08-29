-- BookStats v0.2 cloud-model alignment.
-- Cloud persistence is not active in the v0.2 client yet; this migration keeps
-- the PostgreSQL foundation aligned with the local fields introduced now.

ALTER TABLE works ADD COLUMN IF NOT EXISTS series_name text;
ALTER TABLE works ADD COLUMN IF NOT EXISTS series_volume text;

CREATE TABLE IF NOT EXISTS reading_sessions (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  user_id uuid NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  work_id uuid NOT NULL REFERENCES works(id) ON DELETE CASCADE,
  completed_on date,
  notes text,
  revision bigint NOT NULL DEFAULT 1,
  deleted_at timestamptz,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_reading_sessions_user_completed
  ON reading_sessions(user_id, completed_on);
