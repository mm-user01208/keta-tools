-- ============================================================
-- Stripe Dashboard用 RPC関数（public schemaに作成）
-- GASからSupabase REST API経由でautomationデータを取得するため
-- ============================================================

-- 全チケット一覧を返す
CREATE OR REPLACE FUNCTION get_stripe_dashboard()
RETURNS TABLE (
  id BIGINT,
  payment_intent_id TEXT,
  amount INTEGER,
  currency TEXT,
  status TEXT,
  customer_email TEXT,
  card_brand TEXT,
  card_last4 TEXT,
  card_country TEXT,
  payment_method_type TEXT,
  three_d_secure_result TEXT,
  stripe_created_at TIMESTAMPTZ,
  reconciliation_status TEXT,
  matched_at TIMESTAMPTZ,
  synced_at TIMESTAMPTZ
)
LANGUAGE sql
SECURITY DEFINER
AS $$
  SELECT
    id,
    payment_intent_id,
    amount,
    currency,
    status,
    customer_email,
    card_brand,
    card_last4,
    card_country,
    payment_method_type,
    three_d_secure_result,
    stripe_created_at,
    reconciliation_status,
    matched_at,
    synced_at
  FROM automation.stripe_transactions
  ORDER BY stripe_created_at DESC;
$$;

-- 日次サマリーを返す
CREATE OR REPLACE FUNCTION get_stripe_daily_summary()
RETURNS TABLE (
  report_date DATE,
  total_succeeded BIGINT,
  matched BIGINT,
  unmatched BIGINT,
  pending BIGINT,
  held BIGINT,
  frictionless BIGINT,
  challenge BIGINT,
  three_ds_other BIGINT
)
LANGUAGE sql
SECURITY DEFINER
AS $$
  SELECT
    (stripe_created_at AT TIME ZONE 'Asia/Tokyo')::DATE AS report_date,
    COUNT(*) FILTER (WHERE status = 'succeeded') AS total_succeeded,
    COUNT(*) FILTER (WHERE reconciliation_status = 'matched') AS matched,
    COUNT(*) FILTER (WHERE reconciliation_status = 'unmatched') AS unmatched,
    COUNT(*) FILTER (WHERE reconciliation_status = 'pending' AND status = 'succeeded') AS pending,
    COUNT(*) FILTER (WHERE status IN ('requires_payment_method', 'requires_action', 'requires_capture', 'processing')) AS held,
    COUNT(*) FILTER (WHERE three_d_secure_result IS NULL OR three_d_secure_result = '' OR three_d_secure_result = 'authenticated') AS frictionless,
    COUNT(*) FILTER (WHERE three_d_secure_result = 'challenge') AS challenge,
    COUNT(*) FILTER (WHERE three_d_secure_result IS NOT NULL AND three_d_secure_result != '' AND three_d_secure_result != 'authenticated' AND three_d_secure_result != 'challenge') AS three_ds_other
  FROM automation.stripe_transactions
  GROUP BY (stripe_created_at AT TIME ZONE 'Asia/Tokyo')::DATE
  ORDER BY report_date DESC;
$$;
