---
name: china-market-data
description: Query A-share market data using the AkShare MCP server. Provides real-time quotes, historical prices, financial statements, industry data, and index data for Chinese stocks and funds. Use whenever the agent needs Chinese financial data — primary data source for A-share analysis. Compatible with 沪深 A 股, 科创板, 创业板, 北交所, and 港股通 stocks.
---

# china-market-data

## Data sources

All data is served by the **AkShare** MCP server, which wraps the open-source [AkShare](https://github.com/akfamily/akshare) library (data from East Money / 东方财富, with additional sources).

## Available MCP tools

| Tool | Purpose |
|------|---------|
| `search_stock` | Search A-share stocks by code or name |
| `get_quote` | Real-time quote (price, PE, PB, market cap) |
| `get_historical_data` | OHLCV history (daily/weekly/monthly), forward-adjusted |
| `get_financials` | 利润表/资产负债表/现金流量表 |
| `get_industry_stocks` | 行业分类 & 成分股 (东方财富行业) |
| `get_index_data` | 指数行情 (上证/深证/创业板/科创50) |
| `get_stock_info` | 公司基本信息 & 业务范围 |
| `get_market_overview` | 涨幅榜/跌幅榜/成交额榜 |
| `get_fund_data` | 公募基金/ETF 行情 |

The **china-news** MCP server provides:

| Tool | Purpose |
|------|---------|
| `get_stock_news` | 个股新闻 |
| `get_market_headlines` | 市场头条 |

## Workflow

### 1. Identify the stock/security

Use `search_stock(keyword)` to find the stock code if the user provides a company name in Chinese or English. AkShare codes are:
- A-shares: 6-digit code (e.g., "{{TICKER}}" for {{COMPANY_NAME}})
- SH main board codes start with 6, SZ main board with 00, 创业板 with 30, 科创板 with 688

### 2. Pull market data

```python
# Real-time quote
get_quote(ticker="{{TICKER}}")

# Historical prices (adjusted forward)
get_historical_data(ticker="{{TICKER}}", start_date="{{START_DATE}}", end_date="{{END_DATE}}", frequency="daily")

# Financial statements
get_financials(ticker="{{TICKER}}", statement_type="income", period="annual")
# statement_type: "income" | "balance" | "cashflow"
# period: "annual" | "quarterly"
```

### 3. Industry / peer context

```python
# List all industry sectors
get_industry_stocks()

# Get constituents of a specific industry
get_industry_stocks(industry="白酒")

# Index data
get_index_data(index_code="000001")  # 上证指数
```

### 4. News and sentiment

```python
get_stock_news(ticker="{{TICKER}}")
get_market_headlines(top_n=20)
```

## Note on data coverage

- **A-shares**: Full coverage (SH, SZ, BJ, STAR, ChiNext)
- **Hong Kong stocks (港股通)**: Partial coverage via `ak.stock_hk_spot_em()`
- **US-listed Chinese ADRs**: Not covered — use web search for these
- **Funds**: Public funds and ETFs listed on Chinese exchanges are covered
- **Bonds**: Chinese exchange-traded bonds available via `ak.bond_*` family — extend the MCP server if needed
