-- ============================================================
-- Step 2e: 日次レポート関数
-- ============================================================

CREATE OR REPLACE FUNCTION automation.generate_daily_report()
RETURNS TEXT
LANGUAGE plpgsql
SECURITY DEFINER
AS $$
DECLARE
  v_date DATE := (NOW() AT TIME ZONE 'Asia/Tokyo' - INTERVAL '1 day')::DATE;
  v_total_stripe INTEGER;
  v_total_matched INTEGER;
  v_total_unmatched INTEGER;
  v_total_pending INTEGER;
  v_held_payments INTEGER;
  v_prev_hash TEXT;
  v_report_id BIGINT;
BEGIN
  -- 前日のStripe決済数（JSTベース）
  SELECT COUNT(*) INTO v_total_stripe
  FROM automation.stripe_transactions
  WHERE (stripe_created_at AT TIME ZONE 'Asia/Tokyo')::DATE = v_date
    AND status = 'succeeded';

  -- 突合ステータス別カウント
  SELECT COUNT(*) INTO v_total_matched
  FROM automation.stripe_transactions
  WHERE (stripe_created_at AT TIME ZONE 'Asia/Tokyo')::DATE = v_date
    AND reconciliation_status = 'matched';

  SELECT COUNT(*) INTO v_total_unmatched
  FROM automation.stripe_transactions
  WHERE (stripe_created_at AT TIME ZONE 'Asia/Tokyo')::DATE = v_date
    AND reconciliation_status = 'unmatched';

  SELECT COUNT(*) INTO v_total_pending
  FROM automation.stripe_transactions
  WHERE (stripe_created_at AT TIME ZONE 'Asia/Tokyo')::DATE = v_date
    AND reconciliation_status = 'pending';

  -- 決済保留（requires_payment_method, requires_action等）
  SELECT COUNT(*) INTO v_held_payments
  FROM automation.stripe_transactions
  WHERE (stripe_created_at AT TIME ZONE 'Asia/Tokyo')::DATE = v_date
    AND status IN ('requires_payment_method', 'requires_action', 'requires_capture', 'processing');

  -- 日次レポート保存
  INSERT INTO automation.daily_reports (
    report_date,
    total_stripe_payments,
    matched_count,
    unmatched_count,
    pending_count,
    held_payments_count,
    report_data
  ) VALUES (
    v_date,
    v_total_stripe,
    v_total_matched,
    v_total_unmatched,
    v_total_pending,
    v_held_payments,
    jsonb_build_object(
      'report_date', v_date,
      'total_succeeded', v_total_stripe,
      'matched', v_total_matched,
      'unmatched', v_total_unmatched,
      'pending', v_total_pending,
      'held', v_held_payments,
      'generated_at', NOW()
    )
  )
  ON CONFLICT (report_date) DO UPDATE SET
    total_stripe_payments = EXCLUDED.total_stripe_payments,
    matched_count = EXCLUDED.matched_count,
    unmatched_count = EXCLUDED.unmatched_count,
    pending_count = EXCLUDED.pending_count,
    held_payments_count = EXCLUDED.held_payments_count,
    report_data = EXCLUDED.report_data,
    updated_at = NOW()
  RETURNING id INTO v_report_id;

  RETURN 'OK: report for ' || v_date || ' — succeeded=' || v_total_stripe
    || ' matched=' || v_total_matched
    || ' unmatched=' || v_total_unmatched
    || ' pending=' || v_total_pending
    || ' held=' || v_held_payments;
END;
$$;

-- ============================================================
-- daily_reports テーブルにカラム追加（不足分）
-- ============================================================

DO $$
BEGIN
  -- report_date にユニーク制約が無ければ追加
  IF NOT EXISTS (
    SELECT 1 FROM pg_constraint
    WHERE conrelid = 'automation.daily_reports'::regclass
    AND contype = 'u'
  ) THEN
    ALTER TABLE automation.daily_reports ADD CONSTRAINT daily_reports_date_unique UNIQUE (report_date);
  END IF;
EXCEPTION WHEN OTHERS THEN
  NULL;
END $$;

-- カラムが無ければ追加
ALTER TABLE automation.daily_reports ADD COLUMN IF NOT EXISTS total_stripe_payments INTEGER DEFAULT 0;
ALTER TABLE automation.daily_reports ADD COLUMN IF NOT EXISTS matched_count INTEGER DEFAULT 0;
ALTER TABLE automation.daily_reports ADD COLUMN IF NOT EXISTS unmatched_count INTEGER DEFAULT 0;
ALTER TABLE automation.daily_reports ADD COLUMN IF NOT EXISTS pending_count INTEGER DEFAULT 0;
ALTER TABLE automation.daily_reports ADD COLUMN IF NOT EXISTS held_payments_count INTEGER DEFAULT 0;
ALTER TABLE automation.daily_reports ADD COLUMN IF NOT EXISTS report_data JSONB;
ALTER TABLE automation.daily_reports ADD COLUMN IF NOT EXISTS updated_at TIMESTAMPTZ DEFAULT NOW();

-- ============================================================
-- Step 3: 30分おき自動実行（cron設定）
-- ============================================================

-- 30分おきにStripe同期
SELECT cron.schedule(
  'stripe-sync-every-30min',
  '*/30 * * * *',
  $$SELECT automation.sync_stripe_payments()$$
);

-- 30分おきに突合チェック（同期の1分後）
SELECT cron.schedule(
  'reconcile-every-30min',
  '1-59/30 * * * *',
  $$SELECT automation.reconcile_payments()$$
);

-- 毎日深夜2時に日次レポート生成
SELECT cron.schedule(
  'daily-report-2am',
  '0 17 * * *',
  $$SELECT automation.generate_daily_report()$$
);
-- 注: UTC 17:00 = JST 翌2:00
