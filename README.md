# OpenAlgo Excel Add-In

## Overview
OpenAlgo is an Excel Add-In that provides seamless integration with the OpenAlgo API for algorithmic trading. This add-in allows users to fetch market data, place and manage orders, retrieve historical data, and interact with OpenAlgo's trading infrastructure directly from Excel.

## Features
- **Account Management**: Retrieve funds, order books, trade books, and position books.
- **Market Data**: Fetch real-time market quotes, depth, historical data, and available time intervals.
- **Order Management**: Place, modify, cancel, and retrieve order statuses.
- **Smart & Basket Orders**: Execute split, smart, and bulk orders.
- **Risk Management**: Close all open positions for a given strategy.

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
| `=oa_modifyorder("ORDER_ID", "Strategy", "SYMBOL", "BUY/SELL", "EXCHANGE", "10", "LIMIT", "MIS", "100", "0", "0")` | Modify an order |
| `=oa_cancelorder("ORDER_ID", "Strategy")` | Cancel a specific order |
| `=oa_cancelallorder("Strategy")` | Cancel all orders for a strategy |
| `=oa_closeposition("Strategy")` | Close all open positions for a strategy |
| `=oa_orderstatus("ORDER_ID", "Strategy")` | Retrieve order status |
| `=oa_openposition("Strategy", "SYMBOL", "EXCHANGE", "MIS")` | Fetch open positions |

## Handling API Responses
- Most functions return structured tabular data directly into Excel.
- If an error occurs, the function returns a descriptive error message.
- Market data timestamps are converted to **IST (Indian Standard Time)** for better readability.

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

