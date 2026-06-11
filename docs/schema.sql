-- Sample schema for Weekly Performance Analyzer

CREATE TABLE sample_transactions (
   f0103 TEXT,        -- household id
   f0105 TEXT,
   f0122 TEXT,        -- EAN / SKU
   prod_group TEXT,   -- category
   date DATE
);

CREATE INDEX idx_sample_date ON sample_transactions(date);
CREATE INDEX idx_sample_hh ON sample_transactions(f0103);

