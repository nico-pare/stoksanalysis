# Bigdata.com self-resolving enrichment worker

Goal: for each ticker in your chunk, resolve it to a Bigdata entity ID, pull the
full tearsheet, and extract a fixed set of fields into your output JSONL. Copy
values exactly via jq — never invent. Process EVERY ticker; do not stop early.

## Step 1 — Load the two deferred tools
ToolSearch: `select:mcp__c77e2712-2438-4471-b302-0ae988b3ebdb__find_securities,mcp__c77e2712-2438-4471-b302-0ae988b3ebdb__bigdata_company_tearsheet`

## Step 2 — Read your tickers
Read `/home/user/stoksanalysis/data/chunks/chunk_NN.txt` (NN = your chunk number),
a comma-separated list of tickers.

## Step 3 — For EACH ticker
(a) Resolve: call `find_securities` with query=TICKER. In the returned
`result[0].results` array, choose the entity ID ("id") of the FIRST entry that is
`security_type=="COMPANY"` AND `listing_type=="PUBLIC"` AND has a listing in
`listing_values` ending in `:TICKER` (e.g. "XNYS:TICKER"/"XNAS:TICKER"); prefer
`country=="US"`. If none match that exactly, use the first COMPANY+PUBLIC entry.
If there is no COMPANY result at all, skip the ticker (note it).

(b) Tearsheet: call `bigdata_company_tearsheet` with rp_entity_id = that id,
company_type="Public", sections=["company_overview","financial_ratios",
"analyst_ratings","key_metrics","dividends","latest_earnings"]. The large result is
saved to a file; capture the file PATH from the response. (If returned inline, write
it to /tmp/ts_NN.json and use that path.)

(c) Extract + append — run this Bash (substitute THE_TICKER and THE_PATH):
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
}}' "THE_PATH" >> /home/user/stoksanalysis/data/out/bd_part_NN.jsonl
```
Go ticker-by-ticker (sequential). On a rate-limit/quota error wait ~10s and retry up
to 2x, then skip. Save progress incrementally.

## Step 4 — Report
Reply with: tickers processed, lines written to bd_part_NN.jsonl, and any skipped
(with reason: no match / error).
