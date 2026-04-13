-- ============================================================
-- Stripe Dashboard用 ビュー（public schemaに作成）
-- GASからSupabase REST API経由でautomationデータを取得するため
-- ビューはREST APIで直接クエリ可能
-- ============================================================

-- 既存のRPC関数を削除（不要になるため）
DROP FUNCTION IF EXISTS public.get_stripe_dashboard();
DROP FUNCTION IF EXISTS public.get_stripe_daily_summary();

-- 全チケットビュー
CREATE OR REPLACE VIEW public.stripe_dashboard AS
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

-- 日次サマリービュー
CREATE OR REPLACE VIEW public.stripe_daily_summary AS
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

-- ビューへのアクセス権限
GRANT SELECT ON public.stripe_dashboard TO anon, authenticated, service_role;
GRANT SELECT ON public.stripe_daily_summary TO anon, authenticated, service_role;

-- スキーマキャッシュ更新
NOTIFY pgrst, 'reload schema';
