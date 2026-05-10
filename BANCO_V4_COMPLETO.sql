-- ============================================================
-- CAIXAPRO v4 · Tirzepatida — Script completo do banco
-- Execute no SQL Editor do Supabase
-- ============================================================

-- 1. PRODUTOS (com controle de estoque)
CREATE TABLE IF NOT EXISTS products (
  id              TEXT PRIMARY KEY,
  code            TEXT,
  name            TEXT NOT NULL,
  description     TEXT,
  cost_per_unit   NUMERIC NOT NULL DEFAULT 0,
  price_per_unit  NUMERIC NOT NULL DEFAULT 0,
  units_per_box   INTEGER NOT NULL DEFAULT 4,
  batch           TEXT,
  expiry          DATE,
  stock_qty       INTEGER NOT NULL DEFAULT 0,
  markup          NUMERIC DEFAULT 0,
  margin          NUMERIC DEFAULT 0,
  profit          NUMERIC DEFAULT 0,
  added_by        TEXT,
  created_at      TIMESTAMPTZ DEFAULT NOW()
);

-- 2. VENDAS
CREATE TABLE IF NOT EXISTS sales (
  id              TEXT PRIMARY KEY,
  product_id      TEXT NOT NULL,
  product_name    TEXT NOT NULL,
  client_id       TEXT,
  client_name     TEXT,
  quantity        INTEGER NOT NULL DEFAULT 1,
  unit_type       TEXT DEFAULT 'ampola',
  unit_price      NUMERIC NOT NULL DEFAULT 0,
  total_price     NUMERIC NOT NULL DEFAULT 0,
  payment_method  TEXT DEFAULT 'dinheiro',
  notes           TEXT,
  added_by        TEXT,
  date            TEXT,
  created_at      TIMESTAMPTZ DEFAULT NOW()
);

-- 3. MOVIMENTAÇÃO DE CAIXA
CREATE TABLE IF NOT EXISTS cash_transactions (
  id              TEXT PRIMARY KEY,
  description     TEXT NOT NULL,
  value           NUMERIC NOT NULL,
  type            TEXT NOT NULL CHECK (type IN ('entrada','saida')),
  category        TEXT,
  sale_id         TEXT,
  product_name    TEXT,
  added_by        TEXT,
  date            TEXT,
  created_at      TIMESTAMPTZ DEFAULT NOW()
);

-- 4. CLIENTES
CREATE TABLE IF NOT EXISTS clients (
  id              TEXT PRIMARY KEY,
  name            TEXT NOT NULL,
  email           TEXT,
  phone           TEXT,
  notes           TEXT,
  added_by        TEXT,
  created_at      TIMESTAMPTZ DEFAULT NOW()
);

-- 5. USUÁRIOS DO APP
CREATE TABLE IF NOT EXISTS app_users (
  id              TEXT PRIMARY KEY,
  username        TEXT UNIQUE NOT NULL,
  display_name    TEXT NOT NULL,
  role            TEXT NOT NULL DEFAULT 'operator' CHECK (role IN ('admin','operator','viewer')),
  password_hash   TEXT NOT NULL,
  active          BOOLEAN DEFAULT TRUE,
  last_login      TIMESTAMPTZ,
  created_at      TIMESTAMPTZ DEFAULT NOW()
);

-- 6. Desabilitar RLS em todas as tabelas
ALTER TABLE products          DISABLE ROW LEVEL SECURITY;
ALTER TABLE sales             DISABLE ROW LEVEL SECURITY;
ALTER TABLE cash_transactions DISABLE ROW LEVEL SECURITY;
ALTER TABLE clients           DISABLE ROW LEVEL SECURITY;
ALTER TABLE app_users         DISABLE ROW LEVEL SECURITY;

-- 7. Habilitar Realtime
ALTER PUBLICATION supabase_realtime ADD TABLE products;
ALTER PUBLICATION supabase_realtime ADD TABLE sales;
ALTER PUBLICATION supabase_realtime ADD TABLE cash_transactions;
ALTER PUBLICATION supabase_realtime ADD TABLE clients;
ALTER PUBLICATION supabase_realtime ADD TABLE app_users;

-- 8. Usuários iniciais
-- William Saldanha (admin) — senha: Samuel@2026
INSERT INTO app_users (id, username, display_name, role, password_hash, active)
VALUES ('usr-william-001','williamsaldanha','William Saldanha','admin','U2FtdWVsQDIwMjZ8Y2FpeGFwcm8yMDI2',true)
ON CONFLICT (username) DO NOTHING;

-- Tatyanne (operator) — senha: Samuel@271026
INSERT INTO app_users (id, username, display_name, role, password_hash, active)
VALUES ('usr-tatyanne-002','tatyanne','Tatyanne','operator','U2FtdWVsQDI3MTAyNnxjYWl4YXBybzIwMjY=',true)
ON CONFLICT (username) DO NOTHING;

-- 9. Confirmar
SELECT 'TABELAS CRIADAS COM SUCESSO!' as status;
SELECT table_name FROM information_schema.tables WHERE table_schema = 'public' ORDER BY table_name;
SELECT id, username, display_name, role FROM app_users;
