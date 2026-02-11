# OpenAlgo Excel Add-In

## Overview
OpenAlgo is an Excel Add-In that provides seamless integration with the OpenAlgo API for algorithmic trading. This add-in allows users to fetch market data, place and manage orders, retrieve historical data, and stream real-time market data directly from Excel.

## Features
- **Account Management**: Retrieve funds, order books, trade books, and position books.
- **Market Data**: Fetch real-time market quotes, depth, historical data, and available time intervals.
- **Order Management**: Place, modify, cancel, and retrieve order statuses.
- **Smart & Basket Orders**: Execute split, smart, and bulk orders.
- **Risk Management**: Close all open positions for a given strategy.
- **WebSocket Streaming**: Real-time market data streaming with support for LTP, Quote, and Depth modes.
- **Persistent Configuration**: API key and settings are saved to disk and auto-loaded on Excel restart.

## Prerequisites
- .NET 8.0 Desktop Runtime installed
- Excel-DNA Add-In (included in the project dependencies)
- Microsoft Excel (Office 365 recommended)

## Install the OpenAlgo Excel Add-In

Before installing, ensure you are selecting the correct version based on your Excel installation.

### Steps to Check Your Excel Version

1. Open Microsoft Excel
2. Click **File** > **Account**
3. Click **About Excel**
4. Look for **32-bit** or **64-bit** in the version details.

### Which Version Should You Install?

- If your Excel version is **64-bit** > Install the 64-bit add-in (Recommended)
- If your Excel version is **32-bit** > Install the 32-bit add-in

**Download the OpenAlgo Excel Add-In**: [GitHub Releases](https://github.com/marketcalls/OpenAlgo-Excel/releases)

### .NET 8 Desktop Runtime is Required

OpenAlgo Excel Add-In is built using **Excel-DNA**, which requires the **.NET 8 Desktop Runtime** to run.

If the add-in is not working or Excel does not recognize it, install the .NET 8 Desktop Runtime from:
[Download .NET 8 Desktop Runtime](https://dotnet.microsoft.com/en-us/download/dotnet/8.0)

After installing the runtime, restart your system and try loading the add-in again.

---

## Configuration

### Setting API Key, Version, and Host URL

**Function:** `oa_api(api_key, [version], [host_url])`

This function must be called once to configure the API connection. The configuration is **persisted to disk** at `%LOCALAPPDATA%\OpenAlgo\config.json`, so you only need to call it once. On subsequent Excel sessions, the saved API key is automatically loaded.

All other functions use these stored credentials.

**Parameters:**

| Parameter  | Required | Default                   | Description                |
| ---------- | -------- | ------------------------- | -------------------------- |
| `api_key`  | Yes      | -                         | API key for authentication |
| `version`  | No       | `"v1"`                    | API version                |
| `host_url` | No       | `"http://127.0.0.1:5000"` | OpenAlgo server URL        |

**Example Usage:**

```
=oa_api("your_api_key")
=oa_api("your_api_key", "v1", "http://127.0.0.1:5000")
```

---

## Account Functions

### Retrieve Funds

**Function:** `oa_funds()`

Returns a table with available funds and margin details (available cash, collateral, M2M realized/unrealized, utilised debits).

```
=oa_funds()
```

---

### Retrieve Order Book

**Function:** `oa_orderbook()`

Returns a table with all orders placed during the session (Symbol, Action, Exchange, Quantity, Order Status, Order ID, Price, Price Type, Trigger Price, Product, Timestamp).

```
=oa_orderbook()
```

---

### Retrieve Trade Book

**Function:** `oa_tradebook()`

Returns a table with all executed trades (Symbol, Exchange, Action, Quantity, Product, Timestamp, Trade Value, Average Price, Order ID).

```
=oa_tradebook()
```

---

### Retrieve Position Book

**Function:** `oa_positionbook()`

Returns a table with all open positions (Symbol, Exchange, Quantity, Product, Average Price).

```
=oa_positionbook()
```

---

### Retrieve Holdings

**Function:** `oa_holdings()`

Returns a table with all holdings in the demat account (Symbol, Exchange, Quantity, Product, Pnl, Pnl Percent).

```
=oa_holdings()
```

---

## Market Data Functions

### Get Market Quotes

**Function:** `oa_quotes(symbol, exchange)`

| Parameter  | Required | Description                         |
| ---------- | -------- | ----------------------------------- |
| `symbol`   | Yes      | Trading symbol                      |
| `exchange` | Yes      | Exchange (e.g., NSE, BSE, NFO, MCX) |

Returns market price details for the given symbol (LTP, open, high, low, close, volume, etc.).

```
=oa_quotes("RELIANCE", "NSE")
```

---

### Get Market Depth

**Function:** `oa_depth(symbol, exchange)`

| Parameter  | Required | Description    |
| ---------- | -------- | -------------- |
| `symbol`   | Yes      | Trading symbol |
| `exchange` | Yes      | Exchange       |

Returns order book depth showing LTP, Volume, Ask Price/Quantity, and Bid Price/Quantity levels.

```
=oa_depth("RELIANCE", "NSE")
```

---

### Fetch Historical Data

**Function:** `oa_history(symbol, exchange, interval, start_date, end_date)`

| Parameter    | Required | Description                                              |
| ------------ | -------- | -------------------------------------------------------- |
| `symbol`     | Yes      | Trading symbol                                           |
| `exchange`   | Yes      | Exchange                                                 |
| `interval`   | Yes      | Candle interval (e.g., `"1m"`, `"5m"`, `"15m"`, `"1d"`) |
| `start_date` | Yes      | Start date in `YYYY-MM-DD` format                        |
| `end_date`   | Yes      | End date in `YYYY-MM-DD` format                          |

Returns historical OHLCV data in a table (Ticker, Date, Time IST, Open, High, Low, Close, Volume).

```
=oa_history("RELIANCE", "NSE", "1m", "2024-12-01", "2024-12-31")
```

---

### Get Supported Intervals

**Function:** `oa_intervals()`

Returns a list of all supported candle intervals grouped by category (Seconds, Minutes, Hours, Days, Weeks, Months).

```
=oa_intervals()
```

---

## Order Functions

### Place an Order

**Function:** `oa_placeorder(strategy, symbol, action, exchange, pricetype, product, [quantity], [price], [trigger_price], [disclosed_quantity])`

| Parameter            | Required | Description                              |
| -------------------- | -------- | ---------------------------------------- |
| `strategy`           | Yes      | Strategy name                            |
| `symbol`             | Yes      | Trading symbol                           |
| `action`             | Yes      | `"BUY"` or `"SELL"`                      |
| `exchange`           | Yes      | Exchange (e.g., NSE, BSE, NFO, MCX)      |
| `pricetype`          | Yes      | `"MARKET"`, `"LIMIT"`, `"SL"`, `"SL-M"` |
| `product`            | Yes      | `"MIS"`, `"CNC"`, `"NRML"`              |
| `quantity`           | No       | Number of shares/lots (default: 0)       |
| `price`              | No       | Limit price (default: 0)                 |
| `trigger_price`      | No       | Trigger price (default: 0)               |
| `disclosed_quantity` | No       | Disclosed quantity (default: 0)          |

```
=oa_placeorder("MyStrategy", "INFY", "BUY", "NSE", "LIMIT", "MIS", 10, 1500, 0, 0)
```

---

### Place a Smart Order

**Function:** `oa_placesmartorder(strategy, symbol, action, exchange, pricetype, product, [quantity], [position_size], [price], [trigger_price], [disclosed_quantity])`

Smart orders automatically manage position sizing based on the current position.

| Parameter            | Required | Description                              |
| -------------------- | -------- | ---------------------------------------- |
| `strategy`           | Yes      | Strategy name                            |
| `symbol`             | Yes      | Trading symbol                           |
| `action`             | Yes      | `"BUY"` or `"SELL"`                      |
| `exchange`           | Yes      | Exchange                                 |
| `pricetype`          | Yes      | `"MARKET"`, `"LIMIT"`, `"SL"`, `"SL-M"` |
| `product`            | Yes      | `"MIS"`, `"CNC"`, `"NRML"`              |
| `quantity`           | No       | Number of shares/lots (default: 0)       |
| `position_size`      | No       | Target position size (default: 0)        |
| `price`              | No       | Limit price (default: 0)                 |
| `trigger_price`      | No       | Trigger price (default: 0)               |
| `disclosed_quantity` | No       | Disclosed quantity (default: 0)          |

```
=oa_placesmartorder("SmartStrat", "INFY", "BUY", "NSE", "MARKET", "MIS", 10, 0, 0, 0, 0)
```

---

### Place a Basket Order

**Function:** `oa_basketorder(strategy, orders)`

Place multiple orders at once by referencing an Excel range containing order details.

| Parameter  | Required | Description                                 |
| ---------- | -------- | ------------------------------------------- |
| `strategy` | Yes      | Strategy name                               |
| `orders`   | Yes      | Excel range reference containing order rows |

The referenced range must contain 9 columns in this exact order:

| Column | 1      | 2        | 3      | 4        | 5         | 6       | 7     | 8             | 9                  |
| ------ | ------ | -------- | ------ | -------- | --------- | ------- | ----- | ------------- | ------------------ |
| Field  | Symbol | Exchange | Action | Quantity | PriceType | Product | Price | Trigger Price | Disclosed Quantity |

```
=oa_basketorder("MyStrategy", A2:I5)
```

---

### Place a Split Order

**Function:** `oa_splitorder(strategy, symbol, action, exchange, [quantity], [splitsize], [pricetype], [product], [price], [trigger_price], [disclosed_quantity])`

Splits a large order into smaller chunks to reduce market impact.

| Parameter            | Required | Default    | Description              |
| -------------------- | -------- | ---------- | ------------------------ |
| `strategy`           | Yes      | -          | Strategy name            |
| `symbol`             | Yes      | -          | Trading symbol           |
| `action`             | Yes      | -          | `"BUY"` or `"SELL"`     |
| `exchange`           | Yes      | -          | Exchange                 |
| `quantity`           | No       | 0          | Total order quantity     |
| `splitsize`          | No       | 0          | Size of each split chunk |
| `pricetype`          | No       | `"MARKET"` | Order type               |
| `product`            | No       | `"MIS"`    | Product type             |
| `price`              | No       | 0          | Limit price              |
| `trigger_price`      | No       | 0          | Trigger price            |
| `disclosed_quantity` | No       | 0          | Disclosed quantity       |

```
=oa_splitorder("MyStrategy", "RELIANCE", "BUY", "NSE", 100, 10, "MARKET", "MIS", 0, 0, 0)
```

---

### Modify an Order

**Function:** `oa_modifyorder(strategy, orderid, symbol, action, exchange, [quantity], [pricetype], [product], [price], [trigger_price], [disclosed_quantity])`

| Parameter            | Required | Default    | Description            |
| -------------------- | -------- | ---------- | ---------------------- |
| `strategy`           | Yes      | -          | Strategy name          |
| `orderid`            | Yes      | -          | Order ID to modify     |
| `symbol`             | Yes      | -          | Trading symbol         |
| `action`             | Yes      | -          | `"BUY"` or `"SELL"`   |
| `exchange`           | Yes      | -          | Exchange               |
| `quantity`           | No       | 0          | New quantity           |
| `pricetype`          | No       | `"MARKET"` | New order type         |
| `product`            | No       | `"MIS"`    | New product type       |
| `price`              | No       | 0          | New limit price        |
| `trigger_price`      | No       | 0          | New trigger price      |
| `disclosed_quantity` | No       | 0          | New disclosed quantity |

```
=oa_modifyorder("MyStrategy", "241700000023457", "RELIANCE", "BUY", "NSE", 1, "LIMIT", "MIS", 2500, 0, 0)
```

---

### Cancel an Order

**Function:** `oa_cancelorder(strategy, orderid)`

| Parameter  | Required | Description        |
| ---------- | -------- | ------------------ |
| `strategy` | Yes      | Strategy name      |
| `orderid`  | Yes      | Order ID to cancel |

```
=oa_cancelorder("MyStrategy", "241700000023457")
```

---

### Cancel All Orders

**Function:** `oa_cancelallorder(strategy)`

Cancels all open orders for the given strategy.

| Parameter  | Required | Description   |
| ---------- | -------- | ------------- |
| `strategy` | Yes      | Strategy name |

```
=oa_cancelallorder("MyStrategy")
```

---

### Close All Open Positions

**Function:** `oa_closeposition(strategy)`

| Parameter  | Required | Description   |
| ---------- | -------- | ------------- |
| `strategy` | Yes      | Strategy name |

```
=oa_closeposition("MyStrategy")
```

---

### Get Order Status

**Function:** `oa_orderstatus(strategy, orderid)`

| Parameter  | Required | Description       |
| ---------- | -------- | ----------------- |
| `strategy` | Yes      | Strategy name     |
| `orderid`  | Yes      | Order ID to check |

```
=oa_orderstatus("MyStrategy", "241700000023457")
```

---

### Get Open Position

**Function:** `oa_openposition(strategy, symbol, exchange, product)`

| Parameter  | Required | Description                   |
| ---------- | -------- | ----------------------------- |
| `strategy` | Yes      | Strategy name                 |
| `symbol`   | Yes      | Trading symbol                |
| `exchange` | Yes      | Exchange                      |
| `product`  | Yes      | Product type (MIS, CNC, NRML) |

```
=oa_openposition("MyStrategy", "INFY", "NSE", "MIS")
```

---

## WebSocket Functions (Real-Time Streaming)

WebSocket functions provide real-time streaming market data directly in Excel cells. The add-in handles connection, authentication, and subscription automatically.

### How It Works

1. Call `=oa_api("your_key")` once to set your API key (persisted to disk for future sessions)
2. Use any WebSocket data function - it auto-connects and auto-subscribes
3. Cells update continuously in real-time

### Connection Management

#### Connect to WebSocket

**Function:** `oa_ws_connect([url])`

Connects to the WebSocket server and authenticates. If the API key was previously set via `oa_api()`, the connection authenticates automatically. The function waits for the API key if `oa_api()` hasn't been evaluated yet during recalculation.

| Parameter | Required | Default                  | Description          |
| --------- | -------- | ------------------------ | -------------------- |
| `url`     | No       | `"ws://127.0.0.1:8765"` | WebSocket server URL |

```
=oa_ws_connect()
=oa_ws_connect("ws://127.0.0.1:8765")
=oa_ws_connect("wss://yourdomain.com/ws")
```

---

#### Get Connection Status

**Function:** `oa_ws_status()`

Returns the current WebSocket connection state (Open, Closed, Connecting, Disconnected).

```
=oa_ws_status()
```

---

### Data Functions (Auto-Subscribe)

These volatile functions update continuously in real-time. They automatically connect to the WebSocket server and subscribe to the required data feed if not already subscribed.

#### LTP (Last Traded Price) - Streaming

**Function:** `oa_ws_ltp(symbol, exchange)`

| Parameter  | Required | Description    |
| ---------- | -------- | -------------- |
| `symbol`   | Yes      | Trading symbol |
| `exchange` | Yes      | Exchange       |

Returns a real-time last traded price value that updates automatically.

```
=oa_ws_ltp("RELIANCE", "NSE")
```

---

#### Quote - Streaming

**Function:** `oa_ws_quote(symbol, exchange)`

| Parameter  | Required | Description    |
| ---------- | -------- | -------------- |
| `symbol`   | Yes      | Trading symbol |
| `exchange` | Yes      | Exchange       |

Returns a real-time market quote table (LTP, open, high, low, close, volume, etc.) that updates automatically.

```
=oa_ws_quote("RELIANCE", "NSE")
```

---

#### Depth - Streaming

**Function:** `oa_ws_depth(symbol, exchange, [depth_level])`

| Parameter     | Required | Default | Description                        |
| ------------- | -------- | ------- | ---------------------------------- |
| `symbol`      | Yes      | -       | Trading symbol                     |
| `exchange`    | Yes      | -       | Exchange                           |
| `depth_level` | No       | 5       | Number of depth levels to display  |

Returns a real-time order book depth table showing Bid Orders, Bid Qty, Bid Price, LTP, Ask Price, Ask Qty, Ask Orders.

```
=oa_ws_depth("RELIANCE", "NSE")
=oa_ws_depth("RELIANCE", "NSE", 5)
```

---

#### Get Specific Field - Streaming

**Function:** `oa_ws_field(symbol, exchange, field, [mode])`

Retrieve a specific data field from the WebSocket stream.

| Parameter  | Required | Default | Description                                                   |
| ---------- | -------- | ------- | ------------------------------------------------------------- |
| `symbol`   | Yes      | -       | Trading symbol                                                |
| `exchange` | Yes      | -       | Exchange                                                      |
| `field`    | Yes      | -       | Field name (e.g., `"ltp"`, `"open"`, `"high"`, `"low"`, `"close"`, `"volume"`, `"change_percent"`) |
| `mode`     | No       | 2       | Data mode: 1=LTP, 2=Quote, 3=Depth                           |

```
=oa_ws_field("RELIANCE", "NSE", "ltp")
=oa_ws_field("RELIANCE", "NSE", "volume", 2)
```

---

### Subscription Management

Data functions auto-subscribe, so manual subscription management is optional. Use these when you need fine-grained control.

#### Subscribe to WebSocket Feed

**Function:** `oa_ws_subscribe(symbol, exchange, mode, [depth_level])`

| Parameter     | Required | Description                                  |
| ------------- | -------- | -------------------------------------------- |
| `symbol`      | Yes      | Trading symbol                               |
| `exchange`    | Yes      | Exchange                                     |
| `mode`        | Yes      | Subscription mode (1=LTP, 2=Quote, 3=Depth)  |
| `depth_level` | No       | Number of depth levels (only for mode 3)     |

```
=oa_ws_subscribe("RELIANCE", "NSE", 1)
=oa_ws_subscribe("RELIANCE", "NSE", 3, 5)
```

---

#### Unsubscribe from WebSocket Feed

**Functions:**

| Function                                       | Description                         |
| ---------------------------------------------- | ----------------------------------- |
| `oa_ws_unsubscribe(symbol, exchange, mode)`    | Unsubscribe from a specific mode    |
| `oa_ws_unsubscribe_ltp(symbol, exchange)`      | Unsubscribe from LTP data           |
| `oa_ws_unsubscribe_quote(symbol, exchange)`    | Unsubscribe from Quote data         |
| `oa_ws_unsubscribe_depth(symbol, exchange)`    | Unsubscribe from Depth data         |
| `oa_ws_unsubscribe_all()`                      | Unsubscribe from all feeds          |

```
=oa_ws_unsubscribe("RELIANCE", "NSE", 1)
=oa_ws_unsubscribe_ltp("RELIANCE", "NSE")
=oa_ws_unsubscribe_quote("RELIANCE", "NSE")
=oa_ws_unsubscribe_depth("RELIANCE", "NSE")
=oa_ws_unsubscribe_all()
```

---

#### View Active Subscriptions

**Function:** `oa_ws_subscriptions()`

Returns a table of all active WebSocket subscriptions.

```
=oa_ws_subscriptions()
```

---

#### Debug WebSocket Data

**Function:** `oa_ws_debug(symbol, exchange, mode)`

Shows subscription status and cached data for debugging purposes.

| Parameter  | Required | Description                          |
| ---------- | -------- | ------------------------------------ |
| `symbol`   | Yes      | Trading symbol                       |
| `exchange` | Yes      | Exchange                             |
| `mode`     | Yes      | Data mode (1=LTP, 2=Quote, 3=Depth) |

```
=oa_ws_debug("RELIANCE", "NSE", 2)
```

---

### WebSocket Data Modes

| Mode | Name  | Data Includes                                            |
| ---- | ----- | -------------------------------------------------------- |
| 1    | LTP   | Last traded price, timestamp                             |
| 2    | Quote | OHLC, LTP, volume, change %, open interest               |
| 3    | Depth | Full order book with bid/ask price, quantity, and orders |

---

### WebSocket Quick Start Example

```
' Step 1: Set API key (only needed once, persisted to disk)
=oa_api("your_api_key")

' Step 2: Connect to WebSocket
=oa_ws_connect()

' Step 3: Use data functions - they auto-subscribe
Cell A1: =oa_ws_ltp("RELIANCE", "NSE")
Cell A2: =oa_ws_ltp("TCS", "NSE")
Cell A3: =oa_ws_ltp("INFY", "NSE")
Cell B1: =oa_ws_field("RELIANCE", "NSE", "volume", 2)
Cell B2: =oa_ws_field("TCS", "NSE", "volume", 2)
Cell B3: =oa_ws_field("INFY", "NSE", "volume", 2)

' All cells update continuously in real-time

' Step 4: Clean up when done
=oa_ws_unsubscribe_all()
```

---

## Debugging and Logs

- If functions return `#VALUE!`, check if API credentials are set correctly using `=oa_api()`.
- Ensure the OpenAlgo backend is running and accessible at the configured Host URL.
- Ensure the OpenAlgo WebSocket server is running (default: `ws://127.0.0.1:8765`).
- WebSocket logs are written to `%LOCALAPPDATA%\OpenAlgo\websocket.log`.
- Configuration is saved at `%LOCALAPPDATA%\OpenAlgo\config.json`.

## Notes

- Call `=oa_api("your_api_key")` once to configure the API connection. The key is persisted and auto-loaded on future Excel sessions.
- Test in **OpenAlgo Analyzer Mode** before using in live markets.
- All order functions use `strategy` as the first parameter for consistency.
- By default, missing optional parameters in order functions default to `0`.
- WebSocket data functions (`oa_ws_ltp`, `oa_ws_quote`, `oa_ws_depth`, `oa_ws_field`) are volatile and recalculate automatically.

## Support and Contributions
- **Issues**: Report issues at the repository's issue tracker.
- **Contributions**: PRs are welcome to improve features, documentation, or bug fixes.
- **License**: OpenAlgo is open-source and distributed under the **AGPL-3.0 License**.

## References
- [OpenAlgo API Docs](https://docs.openalgo.in/api-documentation/v1/)
- [Excel-DNA Documentation](https://excel-dna.net/)

## Disclaimer

The creators of this add-in are not responsible for any issues, losses, or damages that may arise from its use. It is strongly recommended to test all functionalities in OpenAlgo Analyzer Mode before applying them to live trading. Always verify API responses and exercise caution while executing trades.
