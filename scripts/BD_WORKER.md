# Bigdata.com enrichment worker

Goal: for each ticker in the map, pull a full Bigdata tearsheet and extract a
fixed set of fields into a JSONL file. Copy values exactly via jq — never invent.

## Step 1 — Load the tool schema (deferred)
Run ToolSearch: `select:mcp__c77e2712-2438-4471-b302-0ae988b3ebdb__bigdata_company_tearsheet`

## Step 2 — Read the map
Read `/home/user/stoksanalysis/data/rpmap.txt`. Each line is: `TICKER RP_ENTITY_ID`.

## Step 3 — For EACH line, call the tearsheet tool
Call `bigdata_company_tearsheet` with:
- rp_entity_id = the 6-char ID
- company_type = "Public"
- sections = ["company_overview","financial_ratios","analyst_ratings","key_metrics","dividends","latest_earnings"]

Because many sections are requested, the result will almost always be **saved to a
file** and the tool response will contain that file PATH (e.g.
`/root/.claude/projects/.../tool-results/mcp-...-bigdata_company_tearsheet-<ts>.txt`).
Capture that path. (If instead the result comes back inline, write it to a temp file
yourself first, e.g. `/tmp/ts.json`, then use that path.)

## Step 4 — Extract with jq and append
Run this Bash command, substituting THE_TICKER and THE_PATH:
```
jq -c --arg sym "THE_TICKER" '{symbol:$sym, bd:{
  company_name:.company_overview.company_name, sector:.company_overview.sector,
  industry:.company_overview.industry, ceo:.company_overview.ceo,
  market_cap:.company_overview.market_cap, currency:.company_overview.currency,
  exchange:.company_overview.exchange, country:.company_overview.country,
  employees:.company_overview.full_time_employees, ipo_date:.company_overview.ipo_date,
  beta:.company_overview.beta, price:.company_overview.price,
  change:.company_overview.change, change_pct:.company_overview.change_percentage,
  volume:.company_overview.volume, avg_volume:.company_overview.average_volume,
  year_high:.price_performance.current_market.year_high,
  year_low:.price_performance.current_market.year_low,
  price_avg_50:.price_performance.current_market.price_avg_50,
  price_avg_200:.price_performance.current_market.price_avg_200,
  pe:.key_financial_highlights.price_to_earnings_ratio_ttm,
  peg:.key_financial_highlights.price_to_earnings_growth_ratio_ttm,
  pb:.key_financial_highlights.price_to_book_ratio_ttm,
  ps:.key_financial_highlights.price_to_sales_ratio_ttm,
  pfcf:.key_financial_highlights.price_to_free_cash_flow_ratio_ttm,
  div_yield:.key_financial_highlights.dividend_yield_ttm,
  div_per_share:.key_financial_highlights.dividend_per_share_ttm,
  payout:.key_financial_highlights.dividend_payout_ratio_ttm,
  eps:.key_financial_highlights.net_income_per_share_ttm,
  rev_per_share:.key_financial_highlights.revenue_per_share_ttm,
  gross_margin:.key_financial_highlights.gross_profit_margin_ttm,
  oper_margin:.key_financial_highlights.operating_profit_margin_ttm,
  net_margin:.key_financial_highlights.net_profit_margin_ttm,
  debt_to_equity:.key_financial_highlights.debt_to_equity_ratio_ttm,
  current_ratio:.key_financial_highlights.current_ratio_ttm,
  target_consensus:.analyst_data.price_targets.target_consensus,
  target_median:.analyst_data.price_targets.target_median,
  target_high:.analyst_data.price_targets.target_high,
  target_low:.analyst_data.price_targets.target_low,
  strong_buy:.analyst_data.ratings.strong_buy, buy:.analyst_data.ratings.buy,
  hold:.analyst_data.ratings.hold, sell:.analyst_data.ratings.sell,
  strong_sell:.analyst_data.ratings.strong_sell, consensus:.analyst_data.ratings.consensus
}}' "THE_PATH" >> /home/user/stoksanalysis/data/out/bd_enrich.jsonl
```
This appends exactly one line per ticker. Do them one at a time (sequential) to be
gentle on the API. If a tearsheet call errors with a rate-limit/quota message, wait
~10s and retry up to 2 times, then skip that ticker and continue.

## Step 5 — Report
Reply with: how many tickers processed, how many lines written to bd_enrich.jsonl,
and any tickers skipped due to errors/quota.
