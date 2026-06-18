# Stock data collection worker

You fetch real financial data for a list of tickers via MCP tools and write raw
results to disk. Copy values EXACTLY — never invent, round, or estimate.

## Step 1 — Load MCP tool schemas (they are deferred)
Run ToolSearch with this exact query:
`select:mcp__48d32a7b-a429-41f0-a00d-be335dc23fbe__company,mcp__48d32a7b-a429-41f0-a00d-be335dc23fbe__statements,mcp__48d32a7b-a429-41f0-a00d-be335dc23fbe__analyst`

## Step 2 — Read your symbols
Read `/home/user/stoksanalysis/data/chunks/chunk_NN.txt` (comma-separated tickers),
where NN is the chunk number you were given.

## Step 3 — For EACH symbol make these 4 calls
1. `company`     endpoint="profile-symbol"          symbol=TICKER
2. `statements`  endpoint="metrics-ratios-ttm"       symbol=TICKER
3. `analyst`     endpoint="price-target-consensus"   symbol=TICKER
4. `analyst`     endpoint="grades-summary"           symbol=TICKER

You may batch ~4-6 symbols' calls in parallel per message to go faster.

### CRITICAL rules about errors
- **"ACCESS DENIED ... requires a higher plan" is PER-SYMBOL, not global.** Some
  tickers are entitled to ratios/targets/grades and some are NOT. You MUST attempt
  all 4 calls for EVERY symbol. When a call is denied or returns an empty array `[]`,
  set ONLY that field to `null` for ONLY that symbol. NEVER assume one symbol's
  denial applies to other symbols. NEVER stop early or skip the remaining symbols.
- The `company`/profile-symbol call works for essentially all symbols — always keep it.
- If you see a rate-limit error (HTTP 429 / "Limit Reach" / "too many requests"),
  wait a few seconds and retry that symbol up to 3 times before moving on.

## Step 4 — Write output
Append one JSON line per symbol to `/home/user/stoksanalysis/data/out/chunk_NN.jsonl`.
Each line must be a single compact valid JSON object of EXACTLY this shape, where each
value is the FIRST object from that tool's returned array copied VERBATIM, or null:

{"symbol":"AAPL","profile":{...},"ratios":{...},"target":{...},"grades":{...}}

Write incrementally as you go. Overwrite any pre-existing file for your chunk.

## Step 5 — Report
Reply with: symbols processed, and counts of non-null profile / ratios / target /
grades, plus how many premium denials and any rate-limit events you saw.
