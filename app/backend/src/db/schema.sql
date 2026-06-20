-- Users table
CREATE TABLE IF NOT EXISTS users (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  username VARCHAR(50) UNIQUE NOT NULL,
  email VARCHAR(100) UNIQUE NOT NULL,
  password_hash VARCHAR(255) NOT NULL,
  role VARCHAR(20) NOT NULL DEFAULT 'reader' CHECK (role IN ('admin', 'editor', 'reader')),
  created_by UUID REFERENCES users(id),
  created_at TIMESTAMPTZ DEFAULT NOW(),
  is_active BOOLEAN DEFAULT TRUE
);

-- Tables metadata
CREATE TABLE IF NOT EXISTS spreadsheets (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  name VARCHAR(255) NOT NULL,
  created_by UUID NOT NULL REFERENCES users(id),
  created_at TIMESTAMPTZ DEFAULT NOW(),
  updated_at TIMESTAMPTZ DEFAULT NOW(),
  is_locked BOOLEAN DEFAULT FALSE,
  backup_enabled BOOLEAN DEFAULT FALSE
);

-- Table permissions (who can access which table)
CREATE TABLE IF NOT EXISTS spreadsheet_permissions (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  spreadsheet_id UUID NOT NULL REFERENCES spreadsheets(id) ON DELETE CASCADE,
  user_id UUID NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  role VARCHAR(20) NOT NULL CHECK (role IN ('editor', 'reader')),
  UNIQUE(spreadsheet_id, user_id)
);

-- Spreadsheet data (full state stored as JSON)
CREATE TABLE IF NOT EXISTS spreadsheet_data (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  spreadsheet_id UUID NOT NULL REFERENCES spreadsheets(id) ON DELETE CASCADE,
  sheet_index INT NOT NULL DEFAULT 0,
  data JSONB NOT NULL DEFAULT '{}',
  updated_at TIMESTAMPTZ DEFAULT NOW(),
  UNIQUE(spreadsheet_id, sheet_index)
);

-- Backups
CREATE TABLE IF NOT EXISTS backups (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  spreadsheet_id UUID REFERENCES spreadsheets(id) ON DELETE SET NULL,
  filename VARCHAR(255) NOT NULL,
  created_at TIMESTAMPTZ DEFAULT NOW(),
  created_by UUID REFERENCES users(id),
  size_bytes BIGINT
);

-- Admin transfer log
CREATE TABLE IF NOT EXISTS admin_transfers (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  from_user_id UUID NOT NULL REFERENCES users(id),
  to_user_id UUID NOT NULL REFERENCES users(id),
  transferred_at TIMESTAMPTZ DEFAULT NOW()
);
