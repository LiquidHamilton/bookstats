-- BookStats cloud schema foundation.
-- The first client iteration is local-first; these tables establish the cloud model
-- that accounts and synchronization will use in the next milestone.

CREATE EXTENSION IF NOT EXISTS pgcrypto;

CREATE TABLE IF NOT EXISTS users (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  email text NOT NULL UNIQUE,
  display_name text NOT NULL,
  password_hash text NOT NULL,
  email_verified_at timestamptz,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS works (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  canonical_title text NOT NULL,
  subtitle text,
  description text,
  first_publication_date date,
  original_language text,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS editions (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  work_id uuid NOT NULL REFERENCES works(id) ON DELETE CASCADE,
  title text NOT NULL,
  isbn_10 text,
  isbn_13 text,
  publisher text,
  publication_date date,
  language text,
  page_count integer CHECK (page_count IS NULL OR page_count >= 0),
  format text,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS user_books (
  id uuid PRIMARY KEY,
  user_id uuid NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  work_id uuid NOT NULL REFERENCES works(id) ON DELETE CASCADE,
  reading_status text NOT NULL DEFAULT 'not_started',
  rating numeric(2,1) CHECK (rating IS NULL OR (rating >= 0 AND rating <= 5)),
  review text,
  private_notes text,
  favorite boolean NOT NULL DEFAULT false,
  revision bigint NOT NULL DEFAULT 1,
  deleted_at timestamptz,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now(),
  UNIQUE(user_id, work_id)
);

CREATE INDEX IF NOT EXISTS idx_editions_isbn13 ON editions(isbn_13);
CREATE INDEX IF NOT EXISTS idx_user_books_sync ON user_books(user_id, updated_at);
