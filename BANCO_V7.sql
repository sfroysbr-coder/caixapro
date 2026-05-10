-- ============================================================
-- CAIXAPRO v7 — Execute no SQL Editor do Supabase
-- SEGURO: não apaga dados existentes
-- ============================================================

-- TABELAS
CREATE TABLE IF NOT EXISTS products (
  id TEXT PRIMARY KEY, code TEXT, name TEXT NOT NULL,
  description TEXT, category TEXT DEFAULT 'outro', unit TEXT DEFAULT 'un',
  cost_per_unit NUMERIC DEFAULT 0, price_per_unit NUMERIC DEFAULT 0,
  units_per_pack INTEGER DEFAULT 1, batch TEXT, expiry DATE,
  stock_qty INTEGER DEFAULT 0, min_stock INTEGER DEFAULT 5,
  supplier_id TEXT, supplier_name TEXT,
  markup NUMERIC DEFAULT 0, margin NUMERIC DEFAULT 0, profit NUMERIC DEFAULT 0,
  added_by TEXT, created_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS sales (
  id TEXT PRIMARY KEY, product_id TEXT, product_name TEXT,
  client_id TEXT, client_name TEXT, quantity INTEGER DEFAULT 1,
  unit_price NUMERIC DEFAULT 0, total_price NUMERIC DEFAULT 0,
  payment_method TEXT DEFAULT 'PIX', notes TEXT,
  added_by TEXT, date TEXT, created_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS cash_transactions (
  id TEXT PRIMARY KEY, description TEXT NOT NULL,
  value NUMERIC NOT NULL, type TEXT NOT NULL CHECK (type IN ('entrada','saida')),
  category TEXT, sale_id TEXT, product_name TEXT,
  added_by TEXT, date TEXT, created_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS clients (
  id TEXT PRIMARY KEY, name TEXT NOT NULL,
  email TEXT, phone TEXT, notes TEXT,
  added_by TEXT, created_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS suppliers (
  id TEXT PRIMARY KEY, name TEXT NOT NULL,
  contact TEXT, phone TEXT, email TEXT, notes TEXT,
  added_by TEXT, created_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS app_users (
  id TEXT PRIMARY KEY, username TEXT UNIQUE NOT NULL,
  display_name TEXT NOT NULL,
  role TEXT DEFAULT 'operator' CHECK (role IN ('admin','operator','viewer')),
  password_hash TEXT NOT NULL, active BOOLEAN DEFAULT TRUE,
  last_login TIMESTAMPTZ, created_at TIMESTAMPTZ DEFAULT NOW()
);

-- COLUNAS NOVAS (seguro)
ALTER TABLE products ADD COLUMN IF NOT EXISTS category TEXT DEFAULT 'outro';
ALTER TABLE products ADD COLUMN IF NOT EXISTS unit TEXT DEFAULT 'un';
ALTER TABLE products ADD COLUMN IF NOT EXISTS min_stock INTEGER DEFAULT 5;
ALTER TABLE products ADD COLUMN IF NOT EXISTS supplier_id TEXT;
ALTER TABLE products ADD COLUMN IF NOT EXISTS supplier_name TEXT;
ALTER TABLE clients  ADD COLUMN IF NOT EXISTS notes TEXT;
ALTER TABLE sales    ADD COLUMN IF NOT EXISTS unit_price NUMERIC DEFAULT 0;
ALTER TABLE cash_transactions ADD COLUMN IF NOT EXISTS sale_id TEXT;

-- DESABILITAR RLS
ALTER TABLE products          DISABLE ROW LEVEL SECURITY;
ALTER TABLE sales             DISABLE ROW LEVEL SECURITY;
ALTER TABLE cash_transactions DISABLE ROW LEVEL SECURITY;
ALTER TABLE clients           DISABLE ROW LEVEL SECURITY;
ALTER TABLE suppliers         DISABLE ROW LEVEL SECURITY;
ALTER TABLE app_users         DISABLE ROW LEVEL SECURITY;

-- REALTIME (adiciona apenas o que não existe)
DO $$
DECLARE tbl TEXT;
BEGIN
  FOREACH tbl IN ARRAY ARRAY['products','sales','cash_transactions','clients','suppliers','app_users'] LOOP
    IF NOT EXISTS (
      SELECT 1 FROM pg_publication_tables
      WHERE pubname = 'supabase_realtime' AND tablename = tbl
    ) THEN
      EXECUTE format('ALTER PUBLICATION supabase_realtime ADD TABLE %I', tbl);
    END IF;
  END LOOP;
END $$;

-- USUÁRIOS INICIAIS
INSERT INTO app_users (id, username, display_name, role, password_hash, active)
VALUES ('usr-william-001','williamsaldanha','William Saldanha','admin','U2FtdWVsQDIwMjZ8Y2FpeGFwcm8yMDI2',true)
ON CONFLICT (username) DO NOTHING;

INSERT INTO app_users (id, username, display_name, role, password_hash, active)
VALUES ('usr-tatyanne-002','tatyanne','Tatyanne','operator','U2FtdWVsQDI3MTAyNnxjYWl4YXBybzIwMjY=',true)
ON CONFLICT (username) DO NOTHING;

-- FORNECEDORES PRÉ-CADASTRADOS
INSERT INTO suppliers (id, name, contact, phone, notes, added_by) VALUES
('sup-indufar',   'Indufar',   'Comercial', '+595 21 000-0000', 'Laboratório Paraguai · T.G. (Tirzepatida)', 'sistema'),
('sup-eticos',    'Éticos',    'Comercial', '+595 21 000-0001', 'Laboratório Paraguai · Lipoless', 'sistema'),
('sup-landerlan', 'Landerlan', 'Comercial', '+595 21 000-0002', 'Laboratório Paraguai · Lipoland', 'sistema'),
('sup-quimfa',    'Quimfa',    'Comercial', '+595 21 000-0003', 'Laboratório Paraguai · Tirzec', 'sistema')
ON CONFLICT (id) DO NOTHING;

-- VERIFICAÇÃO FINAL
SELECT 'SUCESSO v7!' as resultado;
SELECT table_name FROM information_schema.tables WHERE table_schema = 'public' ORDER BY table_name;
SELECT username, display_name, role, active FROM app_users;
SELECT name, notes FROM suppliers;
