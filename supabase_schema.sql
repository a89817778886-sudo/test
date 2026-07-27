-- Схема CRM для Supabase PostgreSQL
-- Все AUTOINCREMENT → SERIAL, INTEGER PRIMARY KEY → BIGSERIAL PRIMARY KEY

CREATE TABLE IF NOT EXISTS customers (
    id BIGSERIAL PRIMARY KEY,
    inn TEXT UNIQUE,
    kpp TEXT, ogrn TEXT, name_full TEXT, name_short TEXT,
    phone TEXT, email TEXT, address TEXT,
    director_position TEXT, director_fio TEXT,
    bank TEXT, bik TEXT, rs TEXT, ks TEXT,
    risk_flag TEXT, risk_checked_at TEXT,
    created_at TEXT NOT NULL, updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS quotes (
    id BIGSERIAL PRIMARY KEY,
    kp_number TEXT NOT NULL,
    customer_id BIGINT REFERENCES customers(id) ON DELETE SET NULL,
    product_type TEXT NOT NULL,
    product_model TEXT,
    include_montage INTEGER NOT NULL DEFAULT 0,
    delivery_city TEXT,
    base_total DOUBLE PRECISION NOT NULL,
    discount_pct DOUBLE PRECISION NOT NULL DEFAULT 0,
    discount_amount DOUBLE PRECISION NOT NULL DEFAULT 0,
    final_total DOUBLE PRECISION NOT NULL,
    status TEXT NOT NULL DEFAULT 'draft',
    currency TEXT NOT NULL DEFAULT 'RUB',
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    sold_at TEXT,
    notes TEXT,
    contact_fio TEXT,
    last_contact_at TEXT,
    request_summary TEXT,
    loss_reason TEXT,
    owner_id BIGINT,
    owner_name TEXT,
    probability_pct DOUBLE PRECISION DEFAULT 50,
    pdf_path TEXT,
    docx_path TEXT
);

CREATE TABLE IF NOT EXISTS quote_items (
    id BIGSERIAL PRIMARY KEY,
    quote_id BIGINT NOT NULL REFERENCES quotes(id) ON DELETE CASCADE,
    code TEXT, name TEXT NOT NULL, unit TEXT DEFAULT 'шт',
    qty DOUBLE PRECISION NOT NULL DEFAULT 1,
    price DOUBLE PRECISION NOT NULL DEFAULT 0,
    total DOUBLE PRECISION NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS sales (
    id BIGSERIAL PRIMARY KEY,
    quote_id BIGINT NOT NULL UNIQUE REFERENCES quotes(id) ON DELETE CASCADE,
    customer_id BIGINT REFERENCES customers(id) ON DELETE SET NULL,
    product_type TEXT, product_model TEXT,
    sale_date TEXT NOT NULL, price DOUBLE PRECISION NOT NULL,
    discount_pct DOUBLE PRECISION NOT NULL DEFAULT 0,
    delivery_city TEXT
);

CREATE TABLE IF NOT EXISTS users (
    id BIGSERIAL PRIMARY KEY,
    username TEXT UNIQUE NOT NULL,
    password_hash TEXT NOT NULL,
    full_name TEXT, role TEXT NOT NULL DEFAULT 'manager',
    telegram_chat_id TEXT, is_active INTEGER NOT NULL DEFAULT 1,
    created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS quote_history (
    id BIGSERIAL PRIMARY KEY,
    quote_id BIGINT NOT NULL REFERENCES quotes(id) ON DELETE CASCADE,
    user_id BIGINT, username TEXT,
    field_changed TEXT NOT NULL, old_value TEXT, new_value TEXT,
    changed_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS customer_notes (
    id BIGSERIAL PRIMARY KEY,
    customer_id BIGINT NOT NULL REFERENCES customers(id) ON DELETE CASCADE,
    user_id BIGINT, username TEXT,
    note TEXT NOT NULL, created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS reminders (
    id BIGSERIAL PRIMARY KEY,
    quote_id BIGINT REFERENCES quotes(id) ON DELETE CASCADE,
    customer_id BIGINT REFERENCES customers(id) ON DELETE CASCADE,
    user_id BIGINT,
    due_at TEXT NOT NULL, message TEXT NOT NULL,
    is_done INTEGER NOT NULL DEFAULT 0, created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS quote_files (
    id BIGSERIAL PRIMARY KEY,
    quote_id BIGINT NOT NULL REFERENCES quotes(id) ON DELETE CASCADE,
    file_type TEXT NOT NULL,
    file_path TEXT NOT NULL,
    filename TEXT,
    file_size INTEGER,
    created_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_quote_files_quote ON quote_files(quote_id);

-- Настройки/секреты (для DaData токена и т.п.)
CREATE TABLE IF NOT EXISTS app_settings (
    key TEXT PRIMARY KEY,
    value TEXT,
    updated_at TEXT NOT NULL
);

-- Индексы для быстрого поиска
CREATE INDEX IF NOT EXISTS idx_customers_inn ON customers(inn);
CREATE INDEX IF NOT EXISTS idx_customers_name ON customers(name_short);
CREATE INDEX IF NOT EXISTS idx_quotes_kp_number ON quotes(kp_number);
CREATE INDEX IF NOT EXISTS idx_quotes_customer ON quotes(customer_id);
CREATE INDEX IF NOT EXISTS idx_quote_items_quote ON quote_items(quote_id);
