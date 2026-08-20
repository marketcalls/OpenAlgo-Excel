using System;
using ExcelDna.Integration;
using ExcelDna.Registration.Utils;
using Newtonsoft.Json.Linq;

namespace OpenAlgo
{
    /// <summary>
    /// Excel functions for WebSocket operations.
    ///
    /// The streaming functions (oa_ws_ltp, oa_ws_quote, oa_ws_depth, oa_ws_field and
    /// oa_ws_orders) are RTD sources, not volatile formulas. Each cell is woken only
    /// when a tick for its own symbol arrives, which is why the workbook stays editable
    /// while data streams.
    /// </summary>
    public static class WebSocketFunctions
    {
        /// <summary>
        /// Largest number of order rows oa_ws_orders will spill by default.
        /// </summary>
        private const int DefaultOrderRows = 50;

        /// <summary>
        /// Connects to the WebSocket server with optional URL configuration
        /// </summary>
        [ExcelFunction(Name = "oa_ws_connect", Description = "Connect to OpenAlgo WebSocket server (optionally specify URL, default: the saved ws_url)")]
        public static object oa_ws_connect(
            [ExcelArgument(Name = "WebSocket URL", Description = "Optional: WebSocket URL (e.g., ws://127.0.0.1:8765 or wss://yourdomain.com/ws)")] object wsUrlOptional)
        {
            // Set WebSocket URL if provided (synchronous, no API key dependency)
            string wsUrl = Arg.Str(wsUrlOptional, OpenAlgoConfig.WebSocketUrl);
            if (!string.IsNullOrWhiteSpace(wsUrl))
            {
                if (!string.Equals(wsUrl, OpenAlgoConfig.WebSocketUrl, StringComparison.Ordinal))
                {
                    OpenAlgoConfig.WebSocketUrl = wsUrl;
                    OpenAlgoConfig.Save();
                }
                WebSocketManager.Instance.SetWebSocketUrl(wsUrl);
            }

            // Excel's RTD collection interval defaults to 2000 ms, which caps every
            // streaming cell at one update every two seconds. AutoOpen already lowers it,
            // but reapply here: the COM object is not always reachable that early, and
            // Excel restores its own value on some transitions.
            ExcelRtdSettings.Apply();

            // API key check moved into ConnectAsync (which waits for oa_api to set the key)
            // This prevents the race condition where oa_ws_connect runs before oa_api during recalculation
            return AsyncTaskUtil.RunTask(nameof(oa_ws_connect), new object[] { wsUrl }, async () =>
            {
                return await WebSocketManager.Instance.ConnectAsync();
            })!;
        }

        /// <summary>
        /// Closes the WebSocket connection and clears all subscriptions
        /// </summary>
        [ExcelFunction(Name = "oa_ws_disconnect", Description = "Unsubscribe everything and close the OpenAlgo WebSocket connection")]
        public static object oa_ws_disconnect()
        {
            return AsyncTaskUtil.RunTask(nameof(oa_ws_disconnect), new object[] { }, async () =>
            {
                return await WebSocketManager.Instance.DisconnectAsync();
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
        /// Reads or sets the minimum gap between two pushed updates for the same cell
        /// </summary>
        /// <summary>
        /// Reads or sets Excel's own RealTimeData collection interval.
        /// </summary>
        [ExcelFunction(Name = "oa_rtd_interval", Description = "Set Excel's RTD collection interval in milliseconds. Excel defaults to 2000, which caps every streaming cell at one update per two seconds. 0 means update as soon as data arrives. Omit the argument to read the current value.", Category = "OpenAlgo Stream")]
        public static object oa_rtd_interval(
            [ExcelArgument(Name = "Milliseconds", Description = "Excel RTD collection interval. 0 updates as soon as data arrives, 2000 is the Excel default, -1 freezes streaming until a manual recalculation.")] object millisecondsOptional)
        {
            if (Arg.IsMissing(millisecondsOptional))
            {
                int live = ExcelRtdSettings.Read();
                return live == -2 ? (object)(double)OpenAlgoConfig.RtdThrottleMs : (double)live;
            }

            int milliseconds = Arg.Int(millisecondsOptional, OpenAlgoConfig.RtdThrottleMs);
            if (milliseconds < -1)
                return "Error: Milliseconds must be -1 or greater";

            OpenAlgoConfig.RtdThrottleMs = milliseconds;
            OpenAlgoConfig.Save();
            ExcelRtdSettings.Apply(milliseconds);

            if (milliseconds == 0)
                return "Excel RTD interval set to 0 ms (updates as soon as data arrives)";
            if (milliseconds == -1)
                return "Excel RTD interval set to -1 (streaming frozen until manual recalculation)";
            return $"Excel RTD interval set to {milliseconds} ms";
        }

        [ExcelFunction(Name = "oa_ws_throttle", Description = "Set the minimum gap in milliseconds between two pushed updates for one streaming cell. 0 pushes every tick. Omit the argument to read the current value.")]
        public static object oa_ws_throttle(
            [ExcelArgument(Name = "Milliseconds", Description = "Minimum gap between pushed updates, in milliseconds. 0 means push every tick.")] object millisecondsOptional)
        {
            if (Arg.IsMissing(millisecondsOptional))
                return (double)OpenAlgoConfig.StreamThrottleMs;

            int milliseconds = Arg.Int(millisecondsOptional, OpenAlgoConfig.StreamThrottleMs);
            if (milliseconds < 0)
                return "Error: Milliseconds must be 0 or greater";

            OpenAlgoConfig.StreamThrottleMs = milliseconds;
            OpenAlgoConfig.Save();
            return (double)OpenAlgoConfig.StreamThrottleMs;
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
            if (mode == 3)
                depthLevel = Arg.Int(depthLevelOptional, 5);

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
        /// Streams the real-time last traded price. Auto-subscribes to Mode 1 on first use.
        /// </summary>
        [ExcelFunction(Name = "oa_ws_ltp", Description = "Stream the real-time Last Traded Price (auto-subscribes to Mode 1)")]
        public static object oa_ws_ltp(
            [ExcelArgument(Name = "Symbol", Description = "Trading symbol")] string symbol,
            [ExcelArgument(Name = "Exchange", Description = "Exchange")] string exchange)
        {
            if (string.IsNullOrWhiteSpace(symbol) || string.IsNullOrWhiteSpace(exchange))
                return "Error: Symbol and Exchange are required";

            return ExcelAsyncUtil.Observe(
                nameof(oa_ws_ltp),
                new object[] { symbol, exchange },
                () => new MarketDataObservable(symbol, exchange, 1, null, MarketDataFormat.Ltp));
        }

        /// <summary>
        /// Streams the real-time quote as a two column key/value table.
        /// Auto-subscribes to Mode 2 on first use.
        /// </summary>
        [ExcelFunction(Name = "oa_ws_quote", Description = "Stream real-time quote data (OHLC, volume, etc.) - auto-subscribes to Mode 2")]
        public static object oa_ws_quote(
            [ExcelArgument(Name = "Symbol", Description = "Trading symbol")] string symbol,
            [ExcelArgument(Name = "Exchange", Description = "Exchange")] string exchange)
        {
            if (string.IsNullOrWhiteSpace(symbol) || string.IsNullOrWhiteSpace(exchange))
                return ExcelTable.Error("Symbol and Exchange are required");

            return ExcelAsyncUtil.Observe(
                nameof(oa_ws_quote),
                new object[] { symbol, exchange },
                () => new MarketDataObservable(symbol, exchange, 2, null,
                    message => MarketDataFormat.Quote(message, symbol, exchange)));
        }

        /// <summary>
        /// Streams the real-time order book as a seven column table.
        /// Auto-subscribes to Mode 3 on first use.
        /// </summary>
        [ExcelFunction(Name = "oa_ws_depth", Description = "Stream real-time market depth (order book) - auto-subscribes to Mode 3")]
        public static object oa_ws_depth(
            [ExcelArgument(Name = "Symbol", Description = "Trading symbol")] string symbol,
            [ExcelArgument(Name = "Exchange", Description = "Exchange")] string exchange,
            [ExcelArgument(Name = "Depth Level", Description = "Optional: Depth level (default: 5)")] object depthLevelOptional)
        {
            if (string.IsNullOrWhiteSpace(symbol) || string.IsNullOrWhiteSpace(exchange))
                return ExcelTable.Error("Symbol and Exchange are required");

            int depthLevel = Arg.Int(depthLevelOptional, 5);
            if (depthLevel < 1)
                depthLevel = 5;

            return ExcelAsyncUtil.Observe(
                nameof(oa_ws_depth),
                new object[] { symbol, exchange, depthLevel },
                () => new MarketDataObservable(symbol, exchange, 3, depthLevel,
                    message => MarketDataFormat.Depth(message, symbol, exchange)));
        }

        /// <summary>
        /// Streams one named field from the real-time payload.
        /// </summary>
        [ExcelFunction(Name = "oa_ws_field", Description = "Stream a single field from real-time data (e.g., 'ltp', 'open', 'high', 'low', 'close', 'volume') - auto-subscribes")]
        public static object oa_ws_field(
            [ExcelArgument(Name = "Symbol", Description = "Trading symbol")] string symbol,
            [ExcelArgument(Name = "Exchange", Description = "Exchange")] string exchange,
            [ExcelArgument(Name = "Field", Description = "Field name (e.g., 'ltp', 'open', 'high', 'low', 'close', 'volume', 'change_percent')")] string field,
            [ExcelArgument(Name = "Mode", Description = "Data mode: 1=LTP, 2=Quote, 3=Depth (default: 2)")] object modeOptional)
        {
            if (string.IsNullOrWhiteSpace(symbol) || string.IsNullOrWhiteSpace(exchange))
                return "Error: Symbol and Exchange are required";

            if (string.IsNullOrWhiteSpace(field))
                return "Error: Field is required";

            int mode = Arg.Int(modeOptional, 2);
            if (mode < 1 || mode > 3)
                return "Error: Mode must be 1 (LTP), 2 (Quote), or 3 (Depth)";

            return ExcelAsyncUtil.Observe(
                nameof(oa_ws_field),
                new object[] { symbol, exchange, field, mode },
                () => new MarketDataObservable(symbol, exchange, mode, mode == 3 ? 5 : (int?)null,
                    message => MarketDataFormat.Field(message, field)));
        }

        /// <summary>
        /// Streams order status updates for the connected account
        /// </summary>
        [ExcelFunction(Name = "oa_ws_orders", Description = "Stream real-time order updates for the account (auto-subscribes via subscribe_orders)")]
        public static object oa_ws_orders(
            [ExcelArgument(Name = "Max Rows", Description = "Optional: largest number of order rows to show (default: 50, cap: 200)")] object maxRowsOptional)
        {
            int maxRows = Arg.Int(maxRowsOptional, DefaultOrderRows);
            if (maxRows < 1)
                maxRows = DefaultOrderRows;
            if (maxRows > 200)
                maxRows = 200;

            return ExcelAsyncUtil.Observe(
                nameof(oa_ws_orders),
                new object[] { maxRows },
                () => new OrderUpdatesObservable(maxRows));
        }

        /// <summary>
        /// Stops the account order update stream
        /// </summary>
        [ExcelFunction(Name = "oa_ws_unsubscribe_orders", Description = "Unsubscribe from real-time order updates")]
        public static object oa_ws_unsubscribe_orders()
        {
            return AsyncTaskUtil.RunTask(nameof(oa_ws_unsubscribe_orders), new object[] { }, async () =>
            {
                return await WebSocketManager.Instance.UnsubscribeOrdersAsync();
            })!;
        }

        /// <summary>
        /// Lists the brokers the OpenAlgo server supports
        /// </summary>
        [ExcelFunction(Name = "oa_ws_brokers", Description = "List the brokers supported by the connected OpenAlgo server")]
        public static object oa_ws_brokers()
        {
            return AsyncTaskUtil.RunTask(nameof(oa_ws_brokers), new object[] { }, async () =>
            {
                var response = await WebSocketManager.Instance.GetSupportedBrokersAsync();

                if (response["status"]?.ToString() == "error")
                    return ExcelTable.Error(response["message"]?.ToString() ?? "Unknown error");

                if (response["brokers"] is not JArray brokers || brokers.Count == 0)
                    return ExcelTable.Single("No brokers reported");

                var result = new object[brokers.Count + 1, 1];
                result[0, 0] = "Broker";
                for (int i = 0; i < brokers.Count; i++)
                    result[i + 1, 0] = ExcelTable.Cell(brokers[i]);

                return result;
            })!;
        }

        /// <summary>
        /// Shows the broker the current session is connected through
        /// </summary>
        [ExcelFunction(Name = "oa_ws_brokerinfo", Description = "Show the broker and adapter status for the authenticated WebSocket session")]
        public static object oa_ws_brokerinfo()
        {
            return AsyncTaskUtil.RunTask(nameof(oa_ws_brokerinfo), new object[] { }, async () =>
            {
                var response = await WebSocketManager.Instance.GetBrokerInfoAsync();

                if (response["status"]?.ToString() == "error")
                    return ExcelTable.Error(response["message"]?.ToString() ?? "Unknown error");

                var display = new JObject(response);
                display.Remove("type");
                return ExcelTable.KeyValue(display);
            })!;
        }

        /// <summary>
        /// Checks the connection with a ping and reports the round trip
        /// </summary>
        [ExcelFunction(Name = "oa_ws_ping", Description = "Ping the WebSocket server and report the round trip in milliseconds")]
        public static object oa_ws_ping()
        {
            return AsyncTaskUtil.RunTask(nameof(oa_ws_ping), new object[] { }, async () =>
            {
                var response = await WebSocketManager.Instance.PingAsync();

                if (response["status"]?.ToString() == "error")
                    return ExcelTable.Error(response["message"]?.ToString() ?? "Unknown error");

                var display = new JObject(response);
                display.Remove("type");
                return ExcelTable.KeyValue(display);
            })!;
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
                    return ExcelTable.Single("No active subscriptions");

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
                return ExcelTable.Error(ex.Message);
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
                var manager = WebSocketManager.Instance;
                var result = new System.Collections.Generic.List<object[]>();
                result.Add(new object[] { "Debug Information", "Value" });
                result.Add(new object[] { "Symbol", symbol });
                result.Add(new object[] { "Exchange", exchange });
                result.Add(new object[] { "Mode", mode });
                result.Add(new object[] { "Expected Key", WebSocketManager.TopicFor(symbol, exchange, mode) });
                result.Add(new object[] { "Connection State", manager.GetConnectionState() });
                result.Add(new object[] { "Authenticated", manager.IsAuthenticated });
                result.Add(new object[] { "WebSocket URL", manager.GetWebSocketUrl() });
                result.Add(new object[] { "Stream Throttle (ms)", OpenAlgoConfig.StreamThrottleMs });
                result.Add(new object[] { "Order Stream", manager.IsSubscribedToOrders });
                result.Add(new object[] { "Buffered Order Updates", manager.GetOrderUpdates().Length });

                // Check if subscription exists
                var subscriptions = manager.GetActiveSubscriptions();
                result.Add(new object[] { "Is Subscribed", manager.IsSubscribed(symbol, exchange, mode) });

                // Try to get market data
                var data = manager.GetMarketData(symbol, exchange, mode);
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

                return ExcelTable.FromRows(result);
            }
            catch (Exception ex)
            {
                return new object[,] { { "Error", ex.Message } };
            }
        }
    }
}
