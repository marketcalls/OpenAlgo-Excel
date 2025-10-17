# OpenAlgo Excel Add-In

## Overview
OpenAlgo is an Excel Add-In that provides seamless integration with the OpenAlgo API for algorithmic trading. This add-in allows users to fetch market data, place and manage orders, retrieve historical data, and interact with OpenAlgo's trading infrastructure directly from Excel.

## Features
- **Account Management**: Retrieve funds, order books, trade books, and position books.
- **Market Data**: Fetch real-time market quotes, depth, historical data, and available time intervals.
- **Order Management**: Place, modify, cancel, and retrieve order statuses.
- **Smart & Basket Orders**: Execute split, smart, and bulk orders.
- **Risk Management**: Close all open positions for a given strategy.
- **WebSocket Streaming**: Real-time market data streaming with support for LTP, Quote, and Depth modes.

## Prerequisites
- .NET 8.0 SDK installed
- Excel-DNA Add-In (included in the project dependencies)
- Microsoft Excel (Office 365 recommended)

## Installation
1. Clone the repository or download the files.
2. Open the project in Visual Studio.
3. Build the project to generate the `.xll` add-in file.
4. Open Excel and load the add-in (`OpenAlgo-AddIn64.xll`).
5. Ensure Excel-DNA dependencies are installed.

## Configuration
To configure API credentials, use the following function:
```excel
=oa_api("YOUR_API_KEY", "v1", "http://127.0.0.1:5000")
```
This sets the API key, API version, and base URL for all further API requests.

## Available Excel Functions

### 📌 Account Management
| Function | Description |
|----------|-------------|
| `=oa_funds()` | Retrieve available funds |
| `=oa_orderbook()` | Fetch open order book |
| `=oa_tradebook()` | Fetch trade book |
| `=oa_positionbook()` | Fetch position book |
| `=oa_holdings()` | Fetch holdings data |

### 📌 Market Data
| Function | Description |
|----------|-------------|
| `=oa_quotes("SYMBOL", "EXCHANGE")` | Retrieve market quotes |
| `=oa_depth("SYMBOL", "EXCHANGE")` | Retrieve bid/ask depth |
| `=oa_history("SYMBOL", "EXCHANGE", "1m", "START_DATE", "END_DATE")` | Fetch historical data |
| `=oa_intervals()` | Retrieve available time intervals |

### 📌 Order Management
| Function | Description |
|----------|-------------|
| `=oa_placeorder("Strategy", "SYMBOL", "BUY/SELL", "EXCHANGE", "LIMIT", "MIS", "10", "100", "0", "0")` | Place an order |
| `=oa_placesmartorder("Strategy", "SYMBOL", "BUY/SELL", "EXCHANGE", "LIMIT", "MIS", "10", "100", "0", "0", "0")` | Place a smart order |
| `=oa_basketorder("Strategy", A1:A10)` | Place multiple orders in a basket |
| `=oa_splitorder("Strategy", "SYMBOL", "BUY/SELL", "EXCHANGE", "100", "10", "LIMIT", "MIS", "100", "0", "0")` | Place split order |
| `=oa_modifyorder("Strategy", "241700000023457", "RELIANCE", "BUY", "NSE", 1, "LIMIT", "MIS", 2500, 0, 0)` | Modify an order |
| `=oa_cancelorder("Strategy", "241700000023457")` | Cancel a specific order |
| `=oa_cancelallorder("Strategy")` | Cancel all orders for a strategy |
| `=oa_closeposition("Strategy")` | Close all open positions for a strategy |
| `=oa_orderstatus("MyStrategy", "241700000023457")` | Retrieve order status |
| `=oa_openposition("Strategy", "SYMBOL", "EXCHANGE", "MIS")` | Fetch open positions |

### 📌 WebSocket Streaming (Real-Time Data)

#### Connection Management
| Function | Description |
|----------|-------------|
| `=oa_ws_connect()` or `=oa_ws_connect([url])` | Connect to WebSocket (merged config+connect). Optional URL, default: ws://127.0.0.1:8765 |
| `=oa_ws_status()` | Get connection status |

#### Data Functions (Auto-Subscribe!)
| Function | Mode | Description |
|----------|------|-------------|
| `=oa_ws_ltp("RELIANCE", "NSE")` | 1 | Get real-time Last Traded Price (auto-subscribes) |
| `=oa_ws_quote("RELIANCE", "NSE")` | 2 | Get real-time quote - OHLC, volume, etc. (auto-subscribes) |
| `=oa_ws_depth("RELIANCE", "NSE")` | 3 | Get real-time market depth (auto-subscribes, default 5 levels) |


#### Subscription Management (Optional - Now Automatic!)
| Function | Description |
|----------|-------------|
| `=oa_ws_subscribe("RELIANCE", "NSE", 1, [level])` | Manual subscribe (LTP=1, Quote=2, Depth=3) |
| `=oa_ws_unsubscribe_ltp("RELIANCE", "NSE")` | **NEW!** Unsubscribe from LTP data |
| `=oa_ws_unsubscribe_quote("RELIANCE", "NSE")` | **NEW!** Unsubscribe from Quote data |
| `=oa_ws_unsubscribe_depth("RELIANCE", "NSE")` | **NEW!** Unsubscribe from Depth data |
| `=oa_ws_unsubscribe("RELIANCE", "NSE", 1)` | Unsubscribe from specific symbol/mode (legacy) |
| `=oa_ws_unsubscribe_all()` | **NEW!** Unsubscribe from all active subscriptions |
| `=oa_ws_subscriptions()` | List all active subscriptions |

#### Debug/Utility
| Function | Description |
|----------|-------------|
| `=oa_ws_debug("RELIANCE", "NSE", 1)` | Debug subscription status and cached data |

## Handling API Responses
- Most functions return structured tabular data directly into Excel.
- If an error occurs, the function returns a descriptive error message.
- Market data timestamps are converted to **IST (Indian Standard Time)** for better readability.

## WebSocket Usage Guide

### ⚡ Quick Start (Simplified Workflow)

**Step 1: Set API Key (once)**
```excel
=oa_api("YOUR_API_KEY", "v1", "http://127.0.0.1:5000")
```

**Step 2: Connect to WebSocket**
```excel
=oa_ws_connect()                              ' Default: ws://127.0.0.1:8765
=oa_ws_connect("ws://127.0.0.1:8765")         ' Custom URL
=oa_ws_connect("wss://yourdomain.com/ws")     ' Production with HTTPS
```

**Step 3: Just use the data functions - they auto-subscribe!**
```excel
=oa_ws_ltp("RELIANCE", "NSE")                  ' LTP - Auto-subscribes to Mode 1
=oa_ws_quote("RELIANCE", "NSE")                ' Quote - Auto-subscribes to Mode 2
=oa_ws_depth("RELIANCE", "NSE")                 ' Depth - Auto-subscribes to Mode 3 (5 levels)
```

**Step 4: When done**
```excel
=oa_ws_unsubscribe_all()                       ' Clean up all subscriptions
```

### 📊 Example: Monitor Multiple Stocks

```excel
' Connect once
=oa_ws_connect()

' Monitor multiple stocks in different cells - all auto-subscribe!
Cell A1: =oa_ws_ltp("RELIANCE", "NSE")
Cell A2: =oa_ws_ltp("TCS", "NSE")
Cell A3: =oa_ws_ltp("INFY", "NSE")

Cell B1: =oa_ws_quote("RELIANCE", "NSE", "volume")
Cell B2: =oa_ws_depth("TCS", "NSE", "volume")
Cell B3: =oa_ws_quote("INFY", "NSE", "volume")

' All cells update continuously in real-time!

' Clean up when done
=oa_ws_unsubscribe_all()
```

### 🔧 Advanced Usage

**Manual Subscription (Optional):**
```excel
=oa_ws_subscribe("RELIANCE", "NSE", 1)         ' Mode 1: LTP
=oa_ws_subscribe("RELIANCE", "NSE", 2)         ' Mode 2: Quote
=oa_ws_subscribe("RELIANCE", "NSE", 3, 5)      ' Mode 3: Depth (5 levels)
```

**Managing Subscriptions:**
```excel
=oa_ws_status()                                ' Check connection status
=oa_ws_subscriptions()                         ' List active subscriptions
=oa_ws_unsubscribe_ltp("RELIANCE", "NSE")       ' Unsubscribe from LTP
=oa_ws_unsubscribe_quote("RELIANCE", "NSE")   ' Unsubscribe from Quote
=oa_ws_unsubscribe_depth("RELIANCE", "NSE")    ' Unsubscribe from Depth
=oa_ws_unsubscribe_all()                       ' Unsubscribe from everything
```

### WebSocket Data Modes

| Mode | Name | Data Includes |
|------|------|---------------|
| 1 | LTP | Last traded price, timestamp |
| 2 | Quote | OHLC, LTP, volume, change %, open interest |
| 3 | Depth | Full order book with bid/ask levels (5-50 levels) |


## Debugging & Logs
- If functions return `#VALUE!`, check if API credentials are set correctly using `=oa_api()`.
- Ensure the OpenAlgo backend is running and accessible at the configured `Host URL`.
- Logs for failed API calls can be checked in the response messages returned to Excel.

## Notes
- Ensure all parameter values are passed as **strings** to match OpenAlgo API specifications.
- By default, missing parameters in order functions are set to `"0"` or reasonable defaults.

## Support & Contributions
- **Issues**: If you find any issues, report them in the repository's issue tracker.
- **Contributions**: PRs are welcome to improve features, documentation, or bug fixes.
- **License**: OpenAlgo is open-source and distributed under the **AGPL-3.0 License**.

## References
- [OpenAlgo API Docs](https://docs.openalgo.in/api-documentation/v1/)
- [Excel-DNA Documentation](https://excel-dna.net/)

## Disclaimer

The creators of this add-in are not responsible for any issues, losses, or damages that may arise from its use. It is strongly recommended to test all functionalities in OpenAlgo Analyzer Mode before applying them to live trading. Always verify API responses and exercise caution while executing trades.

🚀 Happy Trading!

