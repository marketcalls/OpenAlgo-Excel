using System;
using System.Collections.Generic;
using ExcelDna.Integration;
using Newtonsoft.Json.Linq;

namespace OpenAlgo
{
    /// <summary>
    /// RTD source for one streaming cell.
    ///
    /// Excel-DNA gives every distinct (function, arguments) pair its own topic and calls
    /// Subscribe once for it. The observable attaches a listener to WebSocketManager for
    /// the matching market data topic and pushes a new value only when a tick for that
    /// topic arrives. No formula is volatile and Excel is never asked to recalculate, so
    /// the UI stays free for editing, copy and paste.
    /// </summary>
    public sealed class MarketDataObservable : IExcelObservable
    {
        private readonly string _symbol;
        private readonly string _exchange;
        private readonly int _mode;
        private readonly int? _depthLevel;
        private readonly Func<JObject, object> _format;

        /// <summary>
        /// Creates an observable for one symbol, exchange and mode.
        /// </summary>
        /// <param name="symbol">Trading symbol, for example RELIANCE.</param>
        /// <param name="exchange">Exchange code, for example NSE.</param>
        /// <param name="mode">1 for LTP, 2 for Quote, 3 for Depth.</param>
        /// <param name="depthLevel">Depth level requested when mode is 3.</param>
        /// <param name="format">Turns a market data message into the value for the cell.</param>
        public MarketDataObservable(string symbol, string exchange, int mode, int? depthLevel, Func<JObject, object> format)
        {
            _symbol = symbol;
            _exchange = exchange;
            _mode = mode;
            _depthLevel = depthLevel;
            _format = format;
        }

        /// <summary>
        /// Attaches the cell to its topic. Disposing the result detaches it.
        /// </summary>
        public IDisposable Subscribe(IExcelObserver observer)
        {
            var manager = WebSocketManager.Instance;
            string topic = WebSocketManager.TopicFor(_symbol, _exchange, _mode);

            // Attach first so no tick that arrives during start up is missed.
            IDisposable handle = manager.AddListener(topic, message => Push(observer, message));

            // The cell must never sit blank while the first tick is on its way.
            Emit(observer, InitialValue(manager));

            return handle;
        }

        private void Push(IExcelObserver observer, JObject message)
        {
            try
            {
                if (message["type"]?.ToString() == WebSocketManager.StatusMessageType)
                {
                    observer.OnNext(message["message"]?.ToString() ?? "");
                    return;
                }

                observer.OnNext(_format(message));
            }
            catch (Exception ex)
            {
                Emit(observer, $"Error: {ex.Message}");
            }
        }

        private object InitialValue(WebSocketManager manager)
        {
            try
            {
                // A manual unsubscribe is respected: report it and do not resubscribe.
                if (manager.WasManuallyUnsubscribed(_symbol, _exchange, _mode))
                    return "Unsubscribed";

                var cached = manager.GetMarketData(_symbol, _exchange, _mode);
                if (cached != null)
                    return _format(cached);

                if (!manager.IsSubscribed(_symbol, _exchange, _mode))
                {
                    // Must not block the calculation thread: oa_api() may still be waiting
                    // to run and set the API key that ConnectAsync needs.
                    manager.StartAutoSubscribe(_symbol, _exchange, _mode, _depthLevel);
                    return "Subscribing...";
                }

                return "Waiting for data...";
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }

        private static void Emit(IExcelObserver observer, object value)
        {
            try
            {
                observer.OnNext(value);
            }
            catch
            {
                // Excel has torn the topic down. Nothing useful can be done here.
            }
        }
    }

    /// <summary>
    /// RTD source for the account wide order update stream. Pushed the same way as
    /// market data: one listener on the orders topic, no volatile formula.
    /// </summary>
    public sealed class OrderUpdatesObservable : IExcelObservable
    {
        private readonly int _maxRows;

        /// <summary>
        /// Creates the observable.
        /// </summary>
        /// <param name="maxRows">Largest number of order rows to spill.</param>
        public OrderUpdatesObservable(int maxRows)
        {
            _maxRows = maxRows;
        }

        /// <summary>
        /// Attaches the cell to the order update topic. Disposing the result detaches it.
        /// </summary>
        public IDisposable Subscribe(IExcelObserver observer)
        {
            var manager = WebSocketManager.Instance;

            IDisposable handle = manager.AddListener(WebSocketManager.OrdersTopic, message =>
            {
                try
                {
                    if (message["type"]?.ToString() == WebSocketManager.StatusMessageType)
                    {
                        string text = message["message"]?.ToString() ?? "";
                        // Keep showing the table once rows exist, otherwise show the status.
                        var rows = manager.GetOrderUpdates();
                        observer.OnNext(rows.Length > 0
                            ? MarketDataFormat.OrderTable(rows, _maxRows)
                            : ExcelTable.Single(text));
                        return;
                    }

                    observer.OnNext(MarketDataFormat.OrderTable(manager.GetOrderUpdates(), _maxRows));
                }
                catch (Exception ex)
                {
                    try { observer.OnNext(ExcelTable.Error(ex.Message)); } catch { }
                }
            });

            try
            {
                if (manager.OrdersManuallyUnsubscribed)
                {
                    observer.OnNext(ExcelTable.Single("Unsubscribed"));
                }
                else
                {
                    var rows = manager.GetOrderUpdates();
                    if (rows.Length > 0)
                    {
                        observer.OnNext(MarketDataFormat.OrderTable(rows, _maxRows));
                    }
                    else
                    {
                        if (!manager.IsSubscribedToOrders)
                            manager.StartAutoSubscribeOrders();
                        observer.OnNext(ExcelTable.Single("Waiting for order updates..."));
                    }
                }
            }
            catch (Exception ex)
            {
                try { observer.OnNext(ExcelTable.Error(ex.Message)); } catch { }
            }

            return handle;
        }
    }

    /// <summary>
    /// Turns streaming payloads into the exact table shapes the oa_ws_* functions have
    /// always produced, so sheets built against earlier versions keep working.
    /// </summary>
    public static class MarketDataFormat
    {
        /// <summary>
        /// Last traded price as a single number.
        /// </summary>
        public static object Ltp(JObject message)
        {
            var marketData = message["data"];
            if (marketData == null)
                return "No data available";

            double? ltp = marketData["ltp"]?.ToObject<double?>();
            return ltp.HasValue ? ltp.Value : (object)"N/A";
        }

        /// <summary>
        /// Two column key/value table of every field in the quote payload.
        /// </summary>
        public static object Quote(JObject message, string symbol, string exchange)
        {
            if (message["data"] is not JObject marketData)
                return new object[,] { { "No data available" } };

            object[,] result = new object[marketData.Count + 1, 2];
            result[0, 0] = $"{symbol} ({exchange})";
            result[0, 1] = "Value";

            int row = 1;
            foreach (var property in marketData.Properties())
            {
                result[row, 0] = property.Name;
                result[row, 1] = property.Value?.ToString() ?? "N/A";
                row++;
            }

            return result;
        }

        /// <summary>
        /// Seven column order book: bid orders, bid qty, bid price, LTP separator,
        /// ask price, ask qty, ask orders.
        /// </summary>
        public static object Depth(JObject message, string symbol, string exchange)
        {
            if (message["data"] is not JObject marketData)
                return new object[,] { { "No data available" } };

            double ltp = marketData["ltp"]?.ToObject<double?>() ?? 0;

            // OpenAlgo format: depth.buy and depth.sell arrays
            if (marketData["depth"] is not JObject depthObject)
                return new object[,] { { "No depth data available" } };

            JArray buyOrders = depthObject["buy"] as JArray ?? new JArray();
            JArray sellOrders = depthObject["sell"] as JArray ?? new JArray();
            int rowCount = Math.Max(buyOrders.Count, sellOrders.Count);

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

        /// <summary>
        /// A single named field, returned as a number where possible.
        /// </summary>
        public static object Field(JObject message, string field)
        {
            var marketData = message["data"];
            if (marketData == null)
                return "No data";

            var fieldValue = marketData[field];
            if (fieldValue == null)
                return "Field not found";

            // Route through the shared converter rather than parsing here. A bare
            // double.TryParse uses the current culture, so on a comma decimal locale
            // (de-DE, fr-FR and others) a broker sending "1234.56" as a JSON string
            // would be read as the grouped integer 123456, showing a price 100 times
            // too large. ExcelTable.Cell parses with InvariantCulture and also keeps
            // identifier fields as text.
            return ExcelTable.Cell(fieldValue, field);
        }

        private static readonly string[] OrderFields =
        {
            "orderid", "symbol", "exchange", "action", "quantity", "price", "trigger_price",
            "pricetype", "product", "order_status", "filled_quantity", "pending_quantity",
            "average_price", "rejection_reason", "mode", "broker"
        };

        private static readonly string[] OrderHeaders =
        {
            "Order ID", "Symbol", "Exchange", "Action", "Quantity", "Price", "Trigger Price",
            "Price Type", "Product", "Status", "Filled Qty", "Pending Qty",
            "Avg Price", "Rejection Reason", "Mode", "Broker"
        };

        /// <summary>
        /// Table of buffered order updates, newest first, capped at maxRows.
        /// </summary>
        public static object[,] OrderTable(IReadOnlyList<JObject> updates, int maxRows)
        {
            if (updates == null || updates.Count == 0)
                return ExcelTable.Single("No order updates");

            int count = Math.Min(updates.Count, Math.Max(maxRows, 1));
            var result = new object[count + 1, OrderHeaders.Length];

            for (int col = 0; col < OrderHeaders.Length; col++)
                result[0, col] = OrderHeaders[col];

            for (int row = 0; row < count; row++)
            {
                var update = updates[row];
                for (int col = 0; col < OrderFields.Length; col++)
                    result[row + 1, col] = ExcelTable.Cell(update[OrderFields[col]]);
            }

            return result;
        }
    }
}
