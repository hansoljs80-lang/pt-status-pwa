-- 물리치료 현황 PWA Supabase 데이터베이스 테이블 스키마
-- Supabase 대시보드 -> SQL Editor에 붙여넣고 [Run]을 누르면 즉시 테이블이 생성됩니다.

CREATE TABLE IF NOT EXISTS pt_daily_records (
    date TEXT PRIMARY KEY,               -- 일자 (예: '2026-09-29')
    rows_data JSONB NOT NULL DEFAULT '[]'::jsonb, -- 해당 일자의 150개 행 전체 데이터
    total_count INTEGER DEFAULT 0,       -- 환자 수 요약
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL
);

-- RLS (Row Level Security) 활성화
ALTER TABLE pt_daily_records ENABLE ROW LEVEL SECURITY;

-- 누구나 읽기/쓰기 허용 (클리닉 내부 PWA 편의용)
CREATE POLICY "Allow public read pt_daily_records" 
ON pt_daily_records FOR SELECT 
TO anon, authenticated 
USING (true);

CREATE POLICY "Allow public insert/update pt_daily_records" 
ON pt_daily_records FOR ALL 
TO anon, authenticated 
USING (true) 
WITH CHECK (true);
