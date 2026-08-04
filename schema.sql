-- Holly Hill Surplus — shop database schema (Cloudflare D1 / SQLite)
--
-- Apply with:
--   npx wrangler d1 execute holly-hill-shop --file=schema.sql
--
-- Design notes:
--   - `lots` is the source of truth for what's for sale. status flips to
--     'sold' the instant Stripe confirms payment (via webhook) so two
--     buyers can never both "win" the same one-of-a-kind item.
--   - `orders` links a Clerk user_id (from Clerk, not stored here in full)
--     to the lot they bought and the Stripe payment record.
--   - Amounts are stored in cents (integers) to avoid floating-point
--     rounding bugs with money -- standard practice, and what Stripe
--     itself uses.

CREATE TABLE IF NOT EXISTS lots (
    id              TEXT PRIMARY KEY,          -- Lot # / SKU, e.g. "A1234"
    name            TEXT NOT NULL,
    category        TEXT,
    description     TEXT,
    website_price_cents INTEGER NOT NULL,       -- price in cents, e.g. $70.00 -> 7000
    status          TEXT NOT NULL DEFAULT 'available'
                        CHECK (status IN ('available', 'reserved', 'sold', 'shipped')),
    stripe_price_id TEXT,                       -- optional: pre-created Stripe Price object
    image_urls      TEXT,                       -- JSON array of image URLs, e.g. ["https://.../a.jpg", ...]
    created_at      TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at      TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS orders (
    id                  TEXT PRIMARY KEY,       -- our own order id (uuid)
    clerk_user_id       TEXT NOT NULL,           -- who bought it (from Clerk auth)
    lot_id              TEXT NOT NULL REFERENCES lots(id),
    amount_cents        INTEGER NOT NULL,
    status              TEXT NOT NULL DEFAULT 'pending'
                            CHECK (status IN ('pending', 'paid', 'failed', 'refunded')),
    stripe_checkout_session_id TEXT,
    stripe_payment_intent_id   TEXT,
    created_at          TEXT NOT NULL DEFAULT (datetime('now')),
    paid_at             TEXT
);

CREATE INDEX IF NOT EXISTS idx_orders_user ON orders(clerk_user_id);
CREATE INDEX IF NOT EXISTS idx_orders_lot ON orders(lot_id);
CREATE INDEX IF NOT EXISTS idx_lots_status ON lots(status);
