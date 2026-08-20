using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelDna.Integration;
using ExcelDna.Registration.Utils;
using Newtonsoft.Json.Linq;

namespace OpenAlgo
{
    /// <summary>
    /// Account services: funds, order book, trade book, position book, holdings and the
    /// pre-trade margin calculator.
    ///
    /// Every function routes through OpenAlgoClient (which attaches the API key and
    /// normalises failures) and returns a table built by ExcelTable, so numbers reach the
    /// sheet as numbers rather than as text.
    /// </summary>
    public static class AccountsApi
    {
        private const string CategoryName = "OpenAlgo Account";

        /// <summary>
        /// Friendly captions for payload fields whose names do not humanise well,
        /// for example "availablecash" or "totalpnlpercentage".
        /// </summary>
        private static readonly Dictionary<string, string> FieldLabels = new(StringComparer.OrdinalIgnoreCase)
        {
            // funds
            ["availablecash"] = "Available Cash",
            ["collateral"] = "Collateral",
            ["m2mrealized"] = "M2M Realized",
            ["m2munrealized"] = "M2M Unrealized",
            ["utiliseddebits"] = "Utilised Debits",
            // holdings statistics
            ["totalholdingvalue"] = "Total Holding Value",
            ["totalinvvalue"] = "Total Investment Value",
            ["totalprofitandloss"] = "Total Profit And Loss",
            ["totalpnlpercentage"] = "Total PnL Percentage",
            // generic
            ["pnl"] = "PnL",
            ["pnlpercent"] = "PnL Percent",
            ["ltp"] = "LTP",
            ["orderid"] = "Order ID",
            ["pricetype"] = "Price Type"
        };

        /// <summary>
        /// Retrieves account funds: available cash, collateral, realised and unrealised
        /// mark to market and the margin already utilised.
        /// </summary>
        [ExcelFunction(
            Name = "oa_funds",
            Description = "Account funds: available cash, collateral, realised and unrealised M2M and utilised margin. Returns a two column table.",
            Category = CategoryName)]
        public static object oa_funds()
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");

            return AsyncTaskUtil.RunTask(nameof(oa_funds), new object[] { }, async () =>
            {
                var response = await OpenAlgoClient.PostAsync("funds", new JObject());

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                if (response["data"] is not JObject data)
                    return (object)ExcelTable.Error("No funds data returned");

                return (object)LabelledKeyValue(data);
            })!;
        }

        /// <summary>
        /// Retrieves the order book: every order placed for the current trading day.
        /// The statistics block that accompanies the orders is available from
        /// oa_orderbook_stats().
        /// </summary>
        [ExcelFunction(
            Name = "oa_orderbook",
            Description = "All orders placed today, one row per order. Columns cover order id, symbol, exchange, action, quantity, price, trigger price, price type, product, status and timestamp. Use oa_orderbook_stats() for the summary counts.",
            Category = CategoryName)]
        public static object oa_orderbook()
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");

            return AsyncTaskUtil.RunTask(nameof(oa_orderbook), new object[] { }, async () =>
            {
                var response = await OpenAlgoClient.PostAsync("orderbook", new JObject());

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                var orders = ArrayUnder(response, "orders");
                if (orders == null)
                    return (object)ExcelTable.Error("No orders returned");

                string[] fields =
                {
                    "orderid", "symbol", "exchange", "action", "quantity", "price",
                    "trigger_price", "pricetype", "product", "order_status", "timestamp"
                };
                string[] headers =
                {
                    "Order ID", "Symbol", "Exchange", "Action", "Quantity", "Price",
                    "Trigger Price", "Price Type", "Product", "Order Status", "Timestamp"
                };

                return (object)MergedRows(orders, fields, headers, "No orders available");
            })!;
        }

        /// <summary>
        /// Retrieves the statistics block that the order book endpoint returns alongside
        /// the orders: buy, sell, completed, open and rejected order counts.
        /// </summary>
        [ExcelFunction(
            Name = "oa_orderbook_stats",
            Description = "Order book summary counts: buy, sell, completed, open and rejected orders. Companion to oa_orderbook().",
            Category = CategoryName)]
        public static object oa_orderbook_stats()
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");

            return AsyncTaskUtil.RunTask(nameof(oa_orderbook_stats), new object[] { }, async () =>
            {
                var response = await OpenAlgoClient.PostAsync("orderbook", new JObject());

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                if (response["data"] is not JObject data || data["statistics"] is not JObject statistics)
                    return (object)ExcelTable.Error("The orderbook response carried no statistics block");

                return (object)LabelledKeyValue(statistics);
            })!;
        }

        /// <summary>
        /// Retrieves the trade book: every execution for the current trading day.
        /// </summary>
        [ExcelFunction(
            Name = "oa_tradebook",
            Description = "Executed trades for today, one row per fill: order id, symbol, exchange, action, quantity, average price, product, timestamp and trade value.",
            Category = CategoryName)]
        public static object oa_tradebook()
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");

            return AsyncTaskUtil.RunTask(nameof(oa_tradebook), new object[] { }, async () =>
            {
                var response = await OpenAlgoClient.PostAsync("tradebook", new JObject());

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                if (response["data"] is not JArray trades)
                    return (object)ExcelTable.Error("No trade data returned");

                string[] fields =
                {
                    "orderid", "symbol", "exchange", "action", "quantity",
                    "average_price", "product", "timestamp", "trade_value"
                };
                string[] headers =
                {
                    "Order ID", "Symbol", "Exchange", "Action", "Quantity",
                    "Average Price", "Product", "Timestamp", "Trade Value"
                };

                return (object)MergedRows(trades, fields, headers, "No trades available");
            })!;
        }

        /// <summary>
        /// Retrieves the position book, including the live last traded price and the
        /// running profit and loss for every position.
        /// </summary>
        [ExcelFunction(
            Name = "oa_positionbook",
            Description = "Open and closed positions for today with live P&L: symbol, exchange, product, quantity, average price, LTP and PnL. Quantity 0 means the position was closed and the row carries the realised P&L.",
            Category = CategoryName)]
        public static object oa_positionbook()
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");

            return AsyncTaskUtil.RunTask(nameof(oa_positionbook), new object[] { }, async () =>
            {
                var response = await OpenAlgoClient.PostAsync("positionbook", new JObject());

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                if (response["data"] is not JArray positions)
                    return (object)ExcelTable.Error("No position data returned");

                string[] fields = { "symbol", "exchange", "product", "quantity", "average_price", "ltp", "pnl" };
                string[] headers = { "Symbol", "Exchange", "Product", "Quantity", "Average Price", "LTP", "PnL" };

                return (object)MergedRows(positions, fields, headers, "No positions available");
            })!;
        }

        /// <summary>
        /// Retrieves portfolio holdings (delivery positions). The portfolio totals that
        /// accompany the holdings are available from oa_holdings_stats().
        /// </summary>
        [ExcelFunction(
            Name = "oa_holdings",
            Description = "Delivery holdings with P&L: symbol, exchange, product, quantity, PnL and PnL percent. Use oa_holdings_stats() for the portfolio totals.",
            Category = CategoryName)]
        public static object oa_holdings()
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");

            return AsyncTaskUtil.RunTask(nameof(oa_holdings), new object[] { }, async () =>
            {
                var response = await OpenAlgoClient.PostAsync("holdings", new JObject());

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                var holdings = ArrayUnder(response, "holdings");
                if (holdings == null)
                    return (object)ExcelTable.Error("No holdings returned");

                string[] fields = { "symbol", "exchange", "product", "quantity", "pnl", "pnlpercent" };
                string[] headers = { "Symbol", "Exchange", "Product", "Quantity", "PnL", "PnL Percent" };

                return (object)MergedRows(holdings, fields, headers, "No holdings available");
            })!;
        }

        /// <summary>
        /// Retrieves the portfolio statistics block that the holdings endpoint returns
        /// alongside the holdings: market value, invested value and total profit and loss.
        /// </summary>
        [ExcelFunction(
            Name = "oa_holdings_stats",
            Description = "Portfolio totals for the holdings: current market value, invested value, total P&L and total P&L percentage. Companion to oa_holdings().",
            Category = CategoryName)]
        public static object oa_holdings_stats()
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");

            return AsyncTaskUtil.RunTask(nameof(oa_holdings_stats), new object[] { }, async () =>
            {
                var response = await OpenAlgoClient.PostAsync("holdings", new JObject());

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                if (response["data"] is not JObject data || data["statistics"] is not JObject statistics)
                    return (object)ExcelTable.Error("The holdings response carried no statistics block");

                return (object)LabelledKeyValue(statistics);
            })!;
        }

        /// <summary>
        /// Calculates the margin required for a basket of up to 50 positions, before any
        /// order is placed. Reads the basket from a worksheet range.
        /// </summary>
        [ExcelFunction(
            Name = "oa_margin",
            Description = "Pre-trade margin for a basket of up to 50 positions, including hedging benefit. Reads the basket from a range and returns total margin required, SPAN, exposure and margin benefit.",
            Category = CategoryName)]
        public static object oa_margin(
            [ExcelArgument(Name = "Positions", Description = "Range holding the basket, one row per position. With a header row the columns are matched by name (Symbol, Exchange, Action, Quantity, Product, Price Type, Price); without one they are read in that order. Max 50 rows.")] object positions,
            [ExcelArgument(Name = "Exchange", Description = "Exchange used for rows that leave it blank. Default NSE.")] object? exchangeOptional = null,
            [ExcelArgument(Name = "Product", Description = "Product used for rows that leave it blank: MIS, CNC or NRML. Default MIS.")] object? productOptional = null,
            [ExcelArgument(Name = "Price Type", Description = "Price type used for rows that leave it blank: MARKET or LIMIT. Default MARKET.")] object? pricetypeOptional = null)
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");

            string exchange = Arg.Str(exchangeOptional, "NSE").Trim().ToUpperInvariant();
            string product = Arg.Str(productOptional, "MIS").Trim().ToUpperInvariant();
            string pricetype = Arg.Str(pricetypeOptional, "MARKET").Trim().ToUpperInvariant();

            var basket = BuildPositions(positions, exchange, product, pricetype, out string? parseError);
            if (basket == null)
                return ExcelTable.Error(parseError ?? "Could not read the positions range");

            string basketJson = basket.ToString(Newtonsoft.Json.Formatting.None);

            return AsyncTaskUtil.RunTask(
                nameof(oa_margin),
                new object[] { basketJson, exchange, product, pricetype },
                async () =>
                {
                    var payload = new JObject { ["positions"] = basket };
                    var response = await OpenAlgoClient.PostAsync("margin", payload);

                    var error = ExcelTable.ErrorOrNull(response);
                    if (error != null) return (object)error;

                    if (response["data"] is not JObject data)
                        return (object)ExcelTable.Error("No margin data returned. Not every broker supports the margin API.");

                    return (object)LabelledKeyValue(data);
                })!;
        }

        // ------------------------------------------------------------------
        // Helpers
        // ------------------------------------------------------------------

        /// <summary>
        /// Reads the named array out of a response whose "data" is an object, and
        /// tolerates brokers that return the array directly under "data".
        /// </summary>
        private static JArray? ArrayUnder(JObject response, string name)
        {
            var data = response["data"];
            if (data is JArray direct)
                return direct;
            if (data is JObject wrapper && wrapper[name] is JArray nested)
                return nested;
            return null;
        }

        /// <summary>
        /// Builds a row table using the documented column order, then appends any extra
        /// fields the broker returned so nothing in the payload is silently dropped.
        /// </summary>
        private static object[,] MergedRows(JArray items, string[] fields, string[] headers, string emptyMessage)
        {
            var allFields = new List<string>(fields);
            var allHeaders = new List<string>(headers);

            foreach (var item in items.OfType<JObject>())
            {
                foreach (var property in item.Properties())
                {
                    if (allFields.Contains(property.Name, StringComparer.OrdinalIgnoreCase))
                        continue;
                    allFields.Add(property.Name);
                    allHeaders.Add(Label(property.Name));
                }
            }

            return ExcelTable.Rows(items, allFields.ToArray(), allHeaders.ToArray(), emptyMessage);
        }

        /// <summary>
        /// Two column table for a JSON object using friendly captions.
        /// </summary>
        private static object[,] LabelledKeyValue(JObject data)
        {
            if (data.Count == 0)
                return ExcelTable.Error("No data available");

            var pairs = data.Properties()
                .Select(p => new KeyValuePair<string, object>(Label(p.Name), ExcelTable.Cell(p.Value)))
                .ToList();

            return ExcelTable.KeyValue(pairs);
        }

        /// <summary>
        /// Caption for a payload field name.
        /// </summary>
        private static string Label(string field) =>
            FieldLabels.TryGetValue(field, out string? label) ? label : ExcelTable.Humanise(field);

        /// <summary>
        /// Turns the worksheet range passed to oa_margin into the positions array the
        /// margin endpoint expects. Returns null and sets error when the range cannot
        /// be read.
        /// </summary>
        private static JArray? BuildPositions(object positions, string exchange, string product, string pricetype, out string? error)
        {
            error = null;

            if (Arg.IsMissing(positions))
            {
                error = "Positions is required. Pass a range with Symbol, Exchange, Action, Quantity, Product and Price Type columns.";
                return null;
            }

            if (positions is not object[,] grid)
            {
                error = "Positions must be a range of cells, one row per position.";
                return null;
            }

            int rowCount = grid.GetLength(0);
            int colCount = grid.GetLength(1);

            var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            int firstRow = 0;

            if (HasHeaderRow(grid, rowCount, colCount))
            {
                firstRow = 1;
                for (int col = 0; col < colCount; col++)
                {
                    string name = NormaliseField(Arg.Str(grid[0, col]));
                    if (name.Length > 0 && !map.ContainsKey(name))
                        map[name] = col;
                }
            }
            else
            {
                string[] order = { "symbol", "exchange", "action", "quantity", "product", "pricetype", "price" };
                for (int col = 0; col < colCount && col < order.Length; col++)
                    map[order[col]] = col;
            }

            if (!map.ContainsKey("symbol"))
            {
                error = "Could not find a Symbol column in the positions range.";
                return null;
            }

            var basket = new JArray();

            for (int row = firstRow; row < rowCount; row++)
            {
                string symbol = CellText(grid, row, map, "symbol").ToUpperInvariant();
                if (symbol.Length == 0)
                    continue;

                string action = CellText(grid, row, map, "action").ToUpperInvariant();
                if (action.Length == 0)
                {
                    error = $"Row {row + 1}: Action is required and must be BUY or SELL.";
                    return null;
                }

                string quantityText = CellText(grid, row, map, "quantity");
                if (quantityText.Length == 0)
                {
                    error = $"Row {row + 1}: Quantity is required.";
                    return null;
                }

                double quantity = Arg.Num(quantityText, 0);
                if (quantity <= 0)
                {
                    error = $"Row {row + 1}: Quantity must be a positive number.";
                    return null;
                }

                string rowExchange = CellText(grid, row, map, "exchange").ToUpperInvariant();
                string rowProduct = CellText(grid, row, map, "product").ToUpperInvariant();
                string rowPricetype = CellText(grid, row, map, "pricetype").ToUpperInvariant();
                string priceText = CellText(grid, row, map, "price");

                var position = new JObject
                {
                    ["symbol"] = symbol,
                    ["exchange"] = rowExchange.Length > 0 ? rowExchange : exchange,
                    ["action"] = action,
                    ["quantity"] = quantity.ToString("0.############", CultureInfo.InvariantCulture),
                    ["product"] = rowProduct.Length > 0 ? rowProduct : product,
                    ["pricetype"] = rowPricetype.Length > 0 ? rowPricetype : pricetype
                };

                // MarginPositionSchema declares price as a string, exactly like quantity.
                // A JSON number is rejected with "Not a valid string." and the whole
                // request comes back HTTP 400, so send the same textual form.
                if (priceText.Length > 0)
                    position["price"] = Arg.Num(priceText, 0).ToString("0.############", CultureInfo.InvariantCulture);

                basket.Add(position);
            }

            if (basket.Count == 0)
            {
                error = "The positions range held no rows with a symbol.";
                return null;
            }

            if (basket.Count > 50)
            {
                error = $"The margin endpoint accepts at most 50 positions, the range holds {basket.Count}.";
                return null;
            }

            return basket;
        }

        /// <summary>
        /// True when the first row of the range names the columns rather than holding data.
        /// </summary>
        private static bool HasHeaderRow(object[,] grid, int rowCount, int colCount)
        {
            if (rowCount < 2)
                return false;

            for (int col = 0; col < colCount; col++)
            {
                if (NormaliseField(Arg.Str(grid[0, col])) == "symbol")
                    return true;
            }
            return false;
        }

        /// <summary>
        /// Reads one mapped cell from the positions range as trimmed text.
        /// </summary>
        private static string CellText(object[,] grid, int row, Dictionary<string, int> map, string field)
        {
            if (!map.TryGetValue(field, out int col) || col >= grid.GetLength(1))
                return "";
            return Arg.Str(grid[row, col]).Trim();
        }

        /// <summary>
        /// Normalises a column caption typed by the user into a payload field name.
        /// "Price Type", "price_type" and "PRICETYPE" all resolve to "pricetype".
        /// </summary>
        private static string NormaliseField(string caption)
        {
            if (string.IsNullOrWhiteSpace(caption))
                return "";

            string name = new string(caption.Where(char.IsLetterOrDigit).ToArray()).ToLowerInvariant();

            return name switch
            {
                "qty" => "quantity",
                "side" or "transactiontype" or "buysell" => "action",
                "producttype" => "product",
                "ordertype" or "orderpricetype" => "pricetype",
                "tradingsymbol" or "scrip" => "symbol",
                _ => name
            };
        }
    }
}
