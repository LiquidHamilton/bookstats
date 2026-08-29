-- BookStats v1.0.1: durable user-selected cover assets.
-- Image bytes live on the server filesystem; PostgreSQL stores ownership, integrity,
-- provenance, and an opaque access token used by clients to render the asset.

CREATE TABLE IF NOT EXISTS cover_assets (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  user_id uuid NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  content_sha256 text NOT NULL,
  access_token text NOT NULL UNIQUE,
  mime_type text NOT NULL,
  byte_size integer NOT NULL CHECK (byte_size > 0),
  storage_path text NOT NULL,
  source_url text,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now(),
  UNIQUE(user_id, content_sha256)
);

CREATE INDEX IF NOT EXISTS idx_cover_assets_user_id ON cover_assets(user_id);
CREATE INDEX IF NOT EXISTS idx_cover_assets_content_hash ON cover_assets(content_sha256);
