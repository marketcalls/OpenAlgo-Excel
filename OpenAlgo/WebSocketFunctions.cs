using System;
using ExcelDna.Integration;
using ExcelDna.Registration.Utils;
using Newtonsoft.Json.Linq;

namespace OpenAlgo
{
    /// <summary>
    /// Excel functions for WebSocket operations
    /// </summary>
    public static class WebSocketFunctions
    {
        /// <summary>
        /// Connects to the WebSocket server with optional URL configuration
        /// </summary>
        [ExcelFunction(Name = "oa_ws_connect", Description = "Connect to OpenAlgo WebSocket server (optionally specify URL, default: ws://127.0.0.1:8765)")]
        public static object oa_ws_connect(
            [ExcelArgument(Name = "WebSocket URL", Description = "Optional: WebSocket URL (e.g., ws://127.0.0.1:8765 or wss://yourdomain.com/ws)")] object wsUrlOptional)
        {
            if (string.IsNullOrWhiteSpace(OpenAlgoConfig.ApiKey))
                return "Error: API Key not set. Use oa_api() first.";

            // Set WebSocket URL if provided
            if (!(wsUrlOptional is ExcelMissing || wsUrlOptional == null))
            {
                string wsUrl = wsUrlOptional.ToString()!;
                WebSocketManager.Instance.SetWebSocketUrl(wsUrl);
            }

            return AsyncTaskUtil.RunTask(nameof(oa_ws_connect), new object[] { wsUrlOptional ?? "" }, async () =>
            {
                return await WebSocketManager.Instance.ConnectAsync();
            })!;
        }

        /// <summary>
        /// Gets the current WebSocket connection status
        /// </summary>
        [ExcelFunction(Name = "oa_ws_status", Description = "Get WebSocket connection status")]
        public static string oa_ws_status()
        {
            return WebSocketManager.Instance.GetConnectionState();
        }

        /// <summary>
        /// Subscribes to market data for a symbol
        /// </summary>
        [ExcelFunction(Name = "oa_ws_subscribe", Description = "Subscribe to real-time market data (Mode: 1=LTP, 2=Quote, 3=Depth)")]
        public static object oa_ws_subscribe(
            [ExcelArgument(Name = "Symbol", Description = "Trading symbol (e.g., RELIANCE)")] string symbol,
            [ExcelArgument(Name = "Exchange", Description = "Exchange (e.g., NSE, BSE)")] string exchange,
            [ExcelArgument(Name = "Mode", Description = "Data mode: 1=LTP, 2=Quote, 3=Depth")] int mode,
            [ExcelArgument(Name = "Depth Level", Description = "Depth level (only for mode 3, default: 5)")] object depthLevelOptional)
        {
            if (string.IsNullOrWhiteSpace(symbol) || string.IsNullOrWhiteSpace(exchange))
                return "Error: Symbol and Exchange are required";

            if (mode < 1 || mode > 3)
                return "Error: Mode must be 1 (LTP), 2 (Quote), or 3 (Depth)";

            int? depthLevel = null;
            if (mode == 3 && !(depthLevelOptional is ExcelMissing))
            {
                if (int.TryParse(depthLevelOptional?.ToString(), out int dl))
                    depthLevel = dl;
                else
                    depthLevel = 5; // Default depth level
            }

            return AsyncTaskUtil.RunTask(nameof(oa_ws_subscribe), new object[] { symbol, exchange, mode, depthLevel ?? 0 }, async () =>
            {
                return await WebSocketManager.Instance.SubscribeAsync(symbol, exchange, mode, depthLevel);
            })!;
        }

        /// <summary>
        /// Unsubscribes from market data
        /// </summary>
        [ExcelFunction(Name = "oa_ws_unsubscribe", Description = "Unsubscribe from real-time market data")]
        public static object oa_ws_unsubscribe(
            [ExcelArgument(Name = "Symbol", Description = "Trading symbol")] string symbol,
            [ExcelArgument(Name = "Exchange", Description = "Exchange")] string exchange,
            [ExcelArgument(Name = "Mode", Description = "Data mode: 1=LTP, 2=Quote, 3=Depth")] int mode)
        {
            if (string.IsNullOrWhiteSpace(symbol) || string.IsNullOrWhiteSpace(exchange))
                return "Error: Symbol and Exchange are required";

            if (mode < 1 || mode > 3)
                return "Error: Mode must be 1 (LTP), 2 (Quote), or 3 (Depth)";

            return AsyncTaskUtil.RunTask(nameof(oa_ws_unsubscribe), new object[] { symbol, exchange, mode }, async () =>
            {
                return await WebSocketManager.Instance.UnsubscribeAsync(symbol, exchange, mode);
            })!;
        }

        /// <summary>
        /// Unsubscribes from LTP (Mode 1) market data
        /// </summary>
        [ExcelFunction(Name = "oa_ws_unsubscribe_ltp", Description = "Unsubscribe from LTP (Last Traded Price) data")]
        public static object oa_ws_unsubscribe_ltp(
            [ExcelArgument(Name = "Symbol", Description = "Trading symbol")] string symbol,
            [ExcelArgument(Name = "Exchange", Description = "Exchange")] string exchange)
        {
            if (string.IsNullOrWhiteSpace(symbol) || string.IsNullOrWhiteSpace(exchange))
                return "Error: Symbol and Exchange are required";

            return AsyncTaskUtil.RunTask(nameof(oa_ws_unsubscribe_ltp), new object[] { symbol, exchange }, async () =>
            {
                return await WebSocketManager.Instance.UnsubscribeAsync(symbol, exchange, 1);
            })!;
        }

        /// <summary>
        /// Unsubscribes from Quote (Mode 2) market data
        /// </summary>
        [ExcelFunction(Name = "oa_ws_unsubscribe_quote", Description = "Unsubscribe from Quote data")]
        public static object oa_ws_unsubscribe_quote(
            [ExcelArgument(Name = "Symbol", Description = "Trading symbol")] string symbol,
            [ExcelArgument(Name = "Exchange", Description = "Exchange")] string exchange)
        {
            if (string.IsNullOrWhiteSpace(symbol) || string.IsNullOrWhiteSpace(exchange))
                return "Error: Symbol and Exchange are required";

            return AsyncTaskUtil.RunTask(nameof(oa_ws_unsubscribe_quote), new object[] { symbol, exchange }, async () =>
            {
                return await WebSocketManager.Instance.UnsubscribeAsync(symbol, exchange, 2);
            })!;
        }

        /// <summary>
        /// Unsubscribes from Depth (Mode 3) market data
        /// </summary>
        [ExcelFunction(Name = "oa_ws_unsubscribe_depth", Description = "Unsubscribe from Depth (order book) data")]
        public static object oa_ws_unsubscribe_depth(
            [ExcelArgument(Name = "Symbol", Description = "Trading symbol")] string symbol,
            [ExcelArgument(Name = "Exchange", Description = "Exchange")] string exchange)
        {
            if (string.IsNullOrWhiteSpace(symbol) || string.IsNullOrWhiteSpace(exchange))
                return "Error: Symbol and Exchange are required";

            return AsyncTaskUtil.RunTask(nameof(oa_ws_unsubscribe_depth), new object[] { symbol, exchange }, async () =>
            {
                return await WebSocketManager.Instance.UnsubscribeAsync(symbol, exchange, 3);
            })!;
        }

        /// <summary>
        /// Gets real-time LTP (Last Traded Price) data - auto-subscribes if needed
        /// </summary>
        [ExcelFunction(Name = "oa_ws_ltp", Description = "Get real-time Last Traded Price (auto-subscribes to Mode 1)", IsVolatile = true)]
        public static object oa_ws_ltp(
            [ExcelArgument(Name = "Symbol", Description = "Trading symbol")] string symbol,
            [ExcelArgument(Name = "Exchange", Description = "Exchange")] string exchange)
        {
            try
            {
                // Check if manually unsubscribed - don't auto-resubscribe
                if (WebSocketManager.Instance.WasManuallyUnsubscribed(symbol, exchange, 1))
                {
                    return "Unsubscribed";
                }

                // Check if subscribed - if not, auto-subscribe
                if (!WebSocketManager.Instance.IsSubscribed(symbol, exchange, 1))
                {
                    // Not subscribed - auto-subscribe
                    var subscribeTask = WebSocketManager.Instance.SubscribeAsync(symbol, exchange, 1);
                    subscribeTask.Wait(); // Wait for subscription to complete
                    return "Subscribing...";
                }

                // Subscribed - get data
                var data = WebSocketManager.Instance.GetMarketData(symbol, exchange, 1);
                if (data == null)
                    return "Waiting for data...";

                var marketData = data["data"];
                if (marketData == null)
                    return "No data available";

                double? ltp = marketData["ltp"]?.ToObject<double?>();
                if (ltp.HasValue)
                    return ltp.Value;
                return "N/A";
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }

        /// <summary>
        /// Gets real-time quote data - auto-subscribes if needed
        /// </summary>
        [ExcelFunction(Name = "oa_ws_quote", Description = "Get real-time quote data (OHLC, volume, etc.) - auto-subscribes to Mode 2", IsVolatile = true)]
        public static object[,] oa_ws_quote(
            [ExcelArgument(Name = "Symbol", Description = "Trading symbol")] string symbol,
            [ExcelArgument(Name = "Exchange", Description = "Exchange")] string exchange)
        {
            try
            {
                // Check if manually unsubscribed - don't auto-resubscribe
                if (WebSocketManager.Instance.WasManuallyUnsubscribed(symbol, exchange, 2))
                {
                    return new object[,] { { "Unsubscribed" } };
                }

                // Auto-subscribe if not already subscribed
                if (!WebSocketManager.Instance.IsSubscribed(symbol, exchange, 2))
                {
                    var subscribeTask = WebSocketManager.Instance.SubscribeAsync(symbol, exchange, 2);
                    subscribeTask.Wait();
                    return new object[,] { { "Subscribing..." } };
                }

                var data = WebSocketManager.Instance.GetMarketData(symbol, exchange, 2);
                if (data == null)
                    return new object[,] { { "Waiting for data..." } };

                var marketData = data["data"] as JObject;
                if (marketData == null)
                    return new object[,] { { "No data available" } };

                // Create result array with key-value pairs
                object[,] result = new object[marketData.Count + 1, 2];
                result[0, 0] = $"{symbol} ({exchange})";
                result[0, 1] = "Value";

                int row = 1;
                foreach (var prop in marketData.Properties())
                {
                    result[row, 0] = prop.Name;
                    result[row, 1] = prop.Value?.ToString() ?? "N/A";
                    row++;
                }

                return result;
            }
            catch (Exception ex)
            {
                return new object[,] { { $"Error: {ex.Message}" } };
            }
        }

        /// <summary>
        /// Gets real-time market depth data - auto-subscribes if needed
        /// </summary>
        [ExcelFunction(Name = "oa_ws_depth", Description = "Get real-time market depth (order book) - auto-subscribes to Mode 3", IsVolatile = true)]
        public static object[,] oa_ws_depth(
            [ExcelArgument(Name = "Symbol", Description = "Trading symbol")] string symbol,
            [ExcelArgument(Name = "Exchange", Description = "Exchange")] string exchange,
            [ExcelArgument(Name = "Depth Level", Description = "Optional: Depth level (default: 5)")] object depthLevelOptional)
        {
            try
            {
                int? depthLevel = null;
                if (!(depthLevelOptional is ExcelMissing || depthLevelOptional == null))
                {
                    if (int.TryParse(depthLevelOptional?.ToString(), out int dl))
                        depthLevel = dl;
                }

                // Check if manually unsubscribed - don't auto-resubscribe
                if (WebSocketManager.Instance.WasManuallyUnsubscribed(symbol, exchange, 3))
                {
                    return new object[,] { { "Unsubscribed" } };
                }

                // Auto-subscribe if not already subscribed
                if (!WebSocketManager.Instance.IsSubscribed(symbol, exchange, 3))
                {
                    var subscribeTask = WebSocketManager.Instance.SubscribeAsync(symbol, exchange, 3, depthLevel ?? 5);
                    subscribeTask.Wait();
                    return new object[,] { { "Subscribing..." } };
                }

                var data = WebSocketManager.Instance.GetMarketData(symbol, exchange, 3);
                if (data == null)
                    return new object[,] { { "Waiting for data..." } };

                var marketData = data["data"] as JObject;
                if (marketData == null)
                    return new object[,] { { "No data available" } };

                double ltp = marketData["ltp"]?.ToObject<double?>() ?? 0;

                // OpenAlgo format: depth.buy and depth.sell arrays
                var depthObject = marketData["depth"] as JObject;
                if (depthObject == null)
                    return new object[,] { { "No depth data available" } };

                JArray buyOrders = depthObject["buy"] as JArray ?? new JArray();
                JArray sellOrders = depthObject["sell"] as JArray ?? new JArray();
                int rowCount = Math.Max(buyOrders.Count, sellOrders.Count);

                // Create result array: header + column headers + data rows
                // Columns: Bid Orders | Bid Qty | Bid Price | LTP | Ask Price | Ask Qty | Ask Orders
                object[,] resultArray = new object[rowCount + 2, 7];

                // Header row with symbol and LTP
                resultArray[0, 0] = $"{symbol} ({exchange})";
                resultArray[0, 1] = "";
                resultArray[0, 2] = "";
                resultArray[0, 3] = "LTP";
                resultArray[0, 4] = ltp;
                resultArray[0, 5] = "";
                resultArray[0, 6] = "";

                // Column headers
                resultArray[1, 0] = "Bid Orders";
                resultArray[1, 1] = "Bid Qty";
                resultArray[1, 2] = "Bid Price";
                resultArray[1, 3] = "";
                resultArray[1, 4] = "Ask Price";
                resultArray[1, 5] = "Ask Qty";
                resultArray[1, 6] = "Ask Orders";

                // Fill depth data
                for (int i = 0; i < rowCount; i++)
                {
                    // Buy side (bids)
                    if (i < buyOrders.Count)
                    {
                        resultArray[i + 2, 0] = buyOrders[i]["orders"]?.ToObject<int?>() ?? 0;
                        resultArray[i + 2, 1] = buyOrders[i]["quantity"]?.ToObject<int?>() ?? 0;
                        resultArray[i + 2, 2] = buyOrders[i]["price"]?.ToObject<double?>() ?? 0;
                    }
                    else
                    {
                        resultArray[i + 2, 0] = 0;
                        resultArray[i + 2, 1] = 0;
                        resultArray[i + 2, 2] = 0;
                    }

                    // Middle column (empty separator)
                    resultArray[i + 2, 3] = "";

                    // Sell side (asks)
                    if (i < sellOrders.Count)
                    {
                        resultArray[i + 2, 4] = sellOrders[i]["price"]?.ToObject<double?>() ?? 0;
                        resultArray[i + 2, 5] = sellOrders[i]["quantity"]?.ToObject<int?>() ?? 0;
                        resultArray[i + 2, 6] = sellOrders[i]["orders"]?.ToObject<int?>() ?? 0;
                    }
                    else
                    {
                        resultArray[i + 2, 4] = 0;
                        resultArray[i + 2, 5] = 0;
                        resultArray[i + 2, 6] = 0;
                    }
                }

                return resultArray;
            }
            catch (Exception ex)
            {
                return new object[,] { { $"Error: {ex.Message}" } };
            }
        }

        /// <summary>
        /// Lists all active subscriptions
        /// </summary>
        [ExcelFunction(Name = "oa_ws_subscriptions", Description = "List all active WebSocket subscriptions")]
        public static object[,] oa_ws_subscriptions()
        {
            try
            {
                var subscriptions = WebSocketManager.Instance.GetActiveSubscriptions();

                if (subscriptions.Length == 0)
                    return new object[,] { { "No active subscriptions" } };

                object[,] result = new object[subscriptions.Length + 1, 1];
                result[0, 0] = "Active Subscriptions";

                for (int i = 0; i < subscriptions.Length; i++)
                {
                    result[i + 1, 0] = subscriptions[i];
                }

                return result;
            }
            catch (Exception ex)
            {
                return new object[,] { { $"Error: {ex.Message}" } };
            }
        }

        /// <summary>
        /// Unsubscribes from all active subscriptions
        /// </summary>
        [ExcelFunction(Name = "oa_ws_unsubscribe_all", Description = "Unsubscribe from all active subscriptions")]
        public static object oa_ws_unsubscribe_all()
        {
            return AsyncTaskUtil.RunTask(nameof(oa_ws_unsubscribe_all), new object[] { }, async () =>
            {
                return await WebSocketManager.Instance.UnsubscribeAllAsync();
            })!;
        }

        /// <summary>
        /// Debug function to check subscription status and cached data
        /// </summary>
        [ExcelFunction(Name = "oa_ws_debug", Description = "Debug: Check subscription status and cached data keys")]
        public static object[,] oa_ws_debug(
            [ExcelArgument(Name = "Symbol", Description = "Trading symbol")] string symbol,
            [ExcelArgument(Name = "Exchange", Description = "Exchange")] string exchange,
            [ExcelArgument(Name = "Mode", Description = "Mode (1, 2, or 3)")] int mode)
        {
            try
            {
                var result = new System.Collections.Generic.List<object[]>();
                result.Add(new object[] { "Debug Information", "Value" });
                result.Add(new object[] { "Symbol", symbol });
                result.Add(new object[] { "Exchange", exchange });
                result.Add(new object[] { "Mode", mode });
                result.Add(new object[] { "Expected Key", $"{symbol}|{exchange}|{mode}" });

                // Check if subscription exists
                var subscriptions = WebSocketManager.Instance.GetActiveSubscriptions();
                bool isSubscribed = subscriptions.Any(s => s.Contains(symbol) && s.Contains(exchange) && s.Contains(mode.ToString()));
                result.Add(new object[] { "Is Subscribed", isSubscribed });

                // Try to get market data
                var data = WebSocketManager.Instance.GetMarketData(symbol, exchange, mode);
                result.Add(new object[] { "Has Data", data != null });

                if (data != null)
                {
                    result.Add(new object[] { "Data Type", data["type"]?.ToString() ?? "N/A" });
                    result.Add(new object[] { "Data Mode", data["mode"]?.ToString() ?? "N/A" });
                    result.Add(new object[] { "Data Topic", data["topic"]?.ToString() ?? "N/A" });
                }

                result.Add(new object[] { "", "" });
                result.Add(new object[] { "Active Subscriptions:", "" });
                foreach (var sub in subscriptions)
                {
                    result.Add(new object[] { sub, "" });
                }

                object[,] resultArray = new object[result.Count, 2];
                for (int i = 0; i < result.Count; i++)
                {
                    resultArray[i, 0] = result[i][0];
                    resultArray[i, 1] = result[i][1];
                }

                return resultArray;
            }
            catch (Exception ex)
            {
                return new object[,] { { "Error", ex.Message } };
            }
        }

        /// <summary>
        /// Gets a specific field from real-time quote data - auto-subscribes if needed
        /// </summary>
        [ExcelFunction(Name = "oa_ws_field", Description = "Get a specific field from real-time data (e.g., 'ltp', 'open', 'high', 'low', 'close', 'volume') - auto-subscribes", IsVolatile = true)]
        public static object oa_ws_field(
            [ExcelArgument(Name = "Symbol", Description = "Trading symbol")] string symbol,
            [ExcelArgument(Name = "Exchange", Description = "Exchange")] string exchange,
            [ExcelArgument(Name = "Field", Description = "Field name (e.g., 'ltp', 'open', 'high', 'low', 'close', 'volume', 'change_percent')")] string field,
            [ExcelArgument(Name = "Mode", Description = "Data mode: 1=LTP, 2=Quote, 3=Depth (default: 2)")] object modeOptional)
        {
            try
            {
                int mode = 2; // Default to Quote mode
                if (!(modeOptional is ExcelMissing) && int.TryParse(modeOptional?.ToString(), out int m))
                    mode = m;

                // Check if manually unsubscribed - don't auto-resubscribe
                if (WebSocketManager.Instance.WasManuallyUnsubscribed(symbol, exchange, mode))
                {
                    return "Unsubscribed";
                }

                // Auto-subscribe if not already subscribed
                if (!WebSocketManager.Instance.IsSubscribed(symbol, exchange, mode))
                {
                    var subscribeTask = WebSocketManager.Instance.SubscribeAsync(symbol, exchange, mode);
                    subscribeTask.Wait();
                    return "Subscribing...";
                }

                var data = WebSocketManager.Instance.GetMarketData(symbol, exchange, mode);
                if (data == null)
                    return "Waiting for data...";

                var marketData = data["data"];
                if (marketData == null)
                    return "No data";

                var fieldValue = marketData[field];
                if (fieldValue == null)
                    return "Field not found";

                // Try to parse as number, otherwise return as string
                if (double.TryParse(fieldValue.ToString(), out double numValue))
                    return numValue;

                return fieldValue.ToString();
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }
    }
}
