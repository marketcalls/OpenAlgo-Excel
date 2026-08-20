using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelDna.Registration.Utils;
using Newtonsoft.Json.Linq;

namespace OpenAlgo
{
    /// <summary>
    /// Worksheet functions covering the OpenAlgo order management endpoints, plus the
    /// two read only order enquiry functions.
    ///
    /// Every function that changes broker state is gated on
    /// <see cref="OpenAlgoConfig.TradingEnabled"/>. A worksheet function re-evaluates
    /// whenever the sheet recalculates, so an order formula left in a cell would
    /// otherwise fire the same order again on every recalculation.
    ///
    /// All requests go through <see cref="OpenAlgoClient"/>, which owns a single shared
    /// HttpClient, applies the configured timeout, adds the API key and normalises any
    /// failure into a document the callers below turn into readable cell text.
    /// </summary>
    public static class OrderApi
    {
        /// <summary>
        /// Trims and upper-cases a symbol, exchange, action, product or price type.
        ///
        /// The server validates these as exact string sets: action is
        /// OneOf("BUY","SELL","buy","sell") and product is OneOf("MIS","NRML","CNC"),
        /// so a cell reading "Buy", "Mis" or " NSE " is rejected with HTTP 400 rather
        /// than being understood. Excel users type these by hand and Excel itself
        /// capitalises words, so normalising here removes a whole class of confusing
        /// schema errors. GttApi already does the same.
        /// </summary>
        private static string Norm(object? value) =>
            Arg.Str(value).Trim().ToUpperInvariant();

        private const string TradingDisabledMessage =
            "Trading is switched off. Call oa_trading_enabled(TRUE) to allow order functions to send.";

        private const string NoApiKeyMessage =
            "OpenAlgo API Key is not set. Use oa_api()";

        /// <summary>
        /// Places an order through the OpenAlgo placeorder endpoint.
        /// </summary>
        [ExcelFunction(
            Name = "oa_placeorder",
            Description = "Places an order through OpenAlgo API.",
            Category = "OpenAlgo Orders")]
        public static object oa_placeorder(
            [ExcelArgument(Name = "Strategy", Description = "Trading strategy name")] string strategy,
            [ExcelArgument(Name = "Symbol", Description = "Trading symbol")] string symbol,
            [ExcelArgument(Name = "Action", Description = "Order action (BUY/SELL)")] string action,
            [ExcelArgument(Name = "Exchange", Description = "Exchange code(NSE/BSE/NFO/MCX)")] string exchange,
            [ExcelArgument(Name = "PriceType", Description = "Price type (MARKET/LIMIT/SL/SL-M). Defaults to MARKET")] string priceType,
            [ExcelArgument(Name = "Product", Description = "Product type (MIS/CNC/NRML). Defaults to MIS")] string product,
            [ExcelArgument(Name = "Quantity", Description = "Order quantity")] object? quantity = null,
            [ExcelArgument(Name = "Price", Description = "Order price, required for LIMIT and SL orders (optional)")] object? price = null,
            [ExcelArgument(Name = "TriggerPrice", Description = "Trigger price, required for SL and SL-M orders (optional)")] object? triggerPrice = null,
            [ExcelArgument(Name = "DisclosedQuantity", Description = "Disclosed quantity for iceberg orders (optional)")] object? disclosedQuantity = null)
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error(NoApiKeyMessage);
            if (!OpenAlgoConfig.TradingEnabled)
                return ExcelTable.Error(TradingDisabledMessage);

            string? missing = FirstMissing(
                ("Strategy", strategy), ("Symbol", symbol), ("Action", action), ("Exchange", exchange));
            if (missing != null)
                return ExcelTable.Error(missing + " is required.");

            object[,]? quantityError = ReadPositiveQuantity(quantity, out double qty);
            if (quantityError != null)
                return quantityError;

            return AsyncTaskUtil.RunTask(nameof(oa_placeorder), new object[]
            {
                strategy, symbol, action, exchange, priceType, product,
                Arg.Str(quantity), Arg.Str(price), Arg.Str(triggerPrice), Arg.Str(disclosedQuantity)
            }, async () =>
            {
                var payload = new JObject
                {
                    ["strategy"] = Arg.Str(strategy).Trim(),
                    ["symbol"] = Norm(symbol),
                    ["action"] = Norm(action),
                    ["exchange"] = Norm(exchange),
                    ["quantity"] = qty
                };
                AddEnum(payload, "pricetype", priceType);
                AddEnum(payload, "product", product);

                // price, trigger_price and disclosed_quantity are optional on this
                // endpoint. Omitting them is not the same as sending 0: a LIMIT order
                // with an explicit price of 0 is rejected, an absent price is not.
                AddNumber(payload, "price", price);
                AddNumber(payload, "trigger_price", triggerPrice);
                AddInteger(payload, "disclosed_quantity", disclosedQuantity);

                var response = await OpenAlgoClient.PostAsync("placeorder", payload);

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                return (object)ExcelTable.Acknowledgement(response, "orderid");
            })!;
        }

        /// <summary>
        /// Places a position aware smart order through the OpenAlgo placesmartorder endpoint.
        /// </summary>
        [ExcelFunction(
            Name = "oa_placesmartorder",
            Description = "Places a smart order through OpenAlgo API.",
            Category = "OpenAlgo Orders")]
        public static object oa_placesmartorder(
            [ExcelArgument(Name = "Strategy", Description = "Trading strategy name")] string strategy,
            [ExcelArgument(Name = "Symbol", Description = "Trading symbol")] string symbol,
            [ExcelArgument(Name = "Action", Description = "Order action (BUY/SELL)")] string action,
            [ExcelArgument(Name = "Exchange", Description = "Exchange code(NSE/BSE/NFO/MCX)")] string exchange,
            [ExcelArgument(Name = "PriceType", Description = "Price type (MARKET/LIMIT/SL/SL-M). Defaults to MARKET")] string priceType,
            [ExcelArgument(Name = "Product", Description = "Product type (MIS/CNC/NRML). Defaults to MIS")] string product,
            [ExcelArgument(Name = "Quantity", Description = "Order quantity")] object? quantity = null,
            [ExcelArgument(Name = "PositionSize", Description = "Target position size: positive long, negative short, 0 flat. Defaults to 0")] object? positionSize = null,
            [ExcelArgument(Name = "Price", Description = "Order price, required for LIMIT and SL orders (optional)")] object? price = null,
            [ExcelArgument(Name = "TriggerPrice", Description = "Trigger price, required for SL and SL-M orders (optional)")] object? triggerPrice = null,
            [ExcelArgument(Name = "DisclosedQuantity", Description = "Disclosed quantity for iceberg orders (optional)")] object? disclosedQuantity = null)
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error(NoApiKeyMessage);
            if (!OpenAlgoConfig.TradingEnabled)
                return ExcelTable.Error(TradingDisabledMessage);

            string? missing = FirstMissing(
                ("Strategy", strategy), ("Symbol", symbol), ("Action", action), ("Exchange", exchange));
            if (missing != null)
                return ExcelTable.Error(missing + " is required.");

            if (Arg.IsMissing(quantity))
                return ExcelTable.Error("Quantity is required.");

            // A smart order accepts a quantity of zero, which asks OpenAlgo to move the
            // book to the target position size without adding a fresh leg.
            double qty = Arg.Num(quantity);
            if (qty < 0)
                return ExcelTable.Error("Quantity cannot be negative.");

            double target = Arg.Num(positionSize);

            return AsyncTaskUtil.RunTask(nameof(oa_placesmartorder), new object[]
            {
                strategy, symbol, action, exchange, priceType, product,
                Arg.Str(quantity), Arg.Str(positionSize), Arg.Str(price),
                Arg.Str(triggerPrice), Arg.Str(disclosedQuantity)
            }, async () =>
            {
                var payload = new JObject
                {
                    ["strategy"] = Arg.Str(strategy).Trim(),
                    ["symbol"] = Norm(symbol),
                    ["action"] = Norm(action),
                    ["exchange"] = Norm(exchange),
                    ["quantity"] = qty,
                    // position_size is mandatory on this endpoint, so it always travels.
                    ["position_size"] = target
                };
                AddEnum(payload, "pricetype", priceType);
                AddEnum(payload, "product", product);
                AddNumber(payload, "price", price);
                AddNumber(payload, "trigger_price", triggerPrice);
                AddInteger(payload, "disclosed_quantity", disclosedQuantity);

                var response = await OpenAlgoClient.PostAsync("placesmartorder", payload);

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                return (object)ExcelTable.Acknowledgement(response, "orderid");
            })!;
        }

        /// <summary>
        /// Places a basket of orders read from a worksheet range.
        ///
        /// Columns are Symbol, Exchange, Action, Quantity, PriceType, Product, Price,
        /// TriggerPrice, DisclosedQuantity. The first four are required, the rest may be
        /// left out of the range or left blank. A header row is detected and skipped.
        /// </summary>
        [ExcelFunction(
            Name = "oa_basketorder",
            Description = "Places a basket of orders through OpenAlgo API.",
            Category = "OpenAlgo Orders")]
        public static object oa_basketorder(
            [ExcelArgument(Name = "Strategy", Description = "Trading strategy name")] string strategy,
            [ExcelArgument(Name = "Orders", Description = "Range of orders: Symbol, Exchange, Action, Quantity, PriceType, Product, Price, TriggerPrice, DisclosedQuantity")] object[,] orders)
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error(NoApiKeyMessage);
            if (!OpenAlgoConfig.TradingEnabled)
                return ExcelTable.Error(TradingDisabledMessage);

            if (string.IsNullOrWhiteSpace(strategy))
                return ExcelTable.Error("Strategy is required.");

            object[,]? basketError = BuildBasket(orders, out JArray orderList);
            if (basketError != null)
                return basketError;

            return AsyncTaskUtil.RunTask(nameof(oa_basketorder),
                new object[] { strategy, RangeKey(orders) }, async () =>
            {
                var payload = new JObject
                {
                    ["strategy"] = Arg.Str(strategy).Trim(),
                    ["orders"] = orderList
                };

                var response = await OpenAlgoClient.PostAsync("basketorder", payload);

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                // Individual legs can fail while the basket as a whole reports success,
                // so every leg gets its own row carrying its status and message.
                if (response["results"] is JArray results)
                {
                    return (object)ExcelTable.Rows(
                        results,
                        new[] { "symbol", "status", "orderid", "message" },
                        new[] { "Symbol", "Status", "Order ID", "Message" },
                        "No orders were placed");
                }

                return (object)ExcelTable.Acknowledgement(response);
            })!;
        }

        /// <summary>
        /// Splits a large order into several smaller child orders.
        /// </summary>
        [ExcelFunction(
            Name = "oa_splitorder",
            Description = "Places a split order through OpenAlgo API.",
            Category = "OpenAlgo Orders")]
        public static object oa_splitorder(
            [ExcelArgument(Name = "Strategy", Description = "Trading strategy name")] string strategy,
            [ExcelArgument(Name = "Symbol", Description = "Trading symbol")] string symbol,
            [ExcelArgument(Name = "Action", Description = "Order action (BUY/SELL)")] string action,
            [ExcelArgument(Name = "Exchange", Description = "Exchange code(NSE/BSE/NFO/MCX)")] string exchange,
            [ExcelArgument(Name = "Quantity", Description = "Total order quantity to split")] object? quantity = null,
            [ExcelArgument(Name = "SplitSize", Description = "Size of each child order, a positive whole number")] object? splitsize = null,
            [ExcelArgument(Name = "PriceType", Description = "Price type (MARKET/LIMIT/SL/SL-M). Defaults to MARKET")] string priceType = "MARKET",
            [ExcelArgument(Name = "Product", Description = "Product type (MIS/CNC/NRML). Defaults to MIS")] string product = "MIS",
            [ExcelArgument(Name = "Price", Description = "Order price, required for LIMIT and SL orders (optional)")] object? price = null,
            [ExcelArgument(Name = "TriggerPrice", Description = "Trigger price, required for SL and SL-M orders (optional)")] object? triggerPrice = null,
            [ExcelArgument(Name = "DisclosedQuantity", Description = "Disclosed quantity for iceberg orders (optional)")] object? disclosedQuantity = null)
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error(NoApiKeyMessage);
            if (!OpenAlgoConfig.TradingEnabled)
                return ExcelTable.Error(TradingDisabledMessage);

            string? missing = FirstMissing(
                ("Strategy", strategy), ("Symbol", symbol), ("Action", action), ("Exchange", exchange));
            if (missing != null)
                return ExcelTable.Error(missing + " is required.");

            object[,]? quantityError = ReadPositiveQuantity(quantity, out double qty);
            if (quantityError != null)
                return quantityError;

            if (Arg.IsMissing(splitsize))
                return ExcelTable.Error("SplitSize is required.");

            int split = Arg.Int(splitsize);
            if (split < 1)
                return ExcelTable.Error("SplitSize must be a positive whole number.");

            return AsyncTaskUtil.RunTask(nameof(oa_splitorder), new object[]
            {
                strategy, symbol, action, exchange, priceType, product,
                Arg.Str(quantity), Arg.Str(splitsize), Arg.Str(price),
                Arg.Str(triggerPrice), Arg.Str(disclosedQuantity)
            }, async () =>
            {
                var payload = new JObject
                {
                    ["strategy"] = Arg.Str(strategy).Trim(),
                    ["symbol"] = Norm(symbol),
                    ["action"] = Norm(action),
                    ["exchange"] = Norm(exchange),
                    ["quantity"] = qty,
                    ["splitsize"] = split
                };
                AddEnum(payload, "pricetype", priceType);
                AddEnum(payload, "product", product);
                AddNumber(payload, "price", price);
                AddNumber(payload, "trigger_price", triggerPrice);
                AddInteger(payload, "disclosed_quantity", disclosedQuantity);

                var response = await OpenAlgoClient.PostAsync("splitorder", payload);

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                if (response["results"] is JArray results)
                {
                    return (object)ExcelTable.Rows(
                        results,
                        new[] { "order_num", "orderid", "quantity", "status", "message" },
                        new[] { "Order Num", "Order ID", "Quantity", "Status", "Message" },
                        "No child orders were placed");
                }

                return (object)ExcelTable.Acknowledgement(response);
            })!;
        }

        /// <summary>
        /// Modifies an open order.
        /// </summary>
        [ExcelFunction(
            Name = "oa_modifyorder",
            Description = "Modifies an existing order through OpenAlgo API.",
            Category = "OpenAlgo Orders")]
        public static object oa_modifyorder(
            [ExcelArgument(Name = "Strategy", Description = "Trading strategy name")] string strategy,
            [ExcelArgument(Name = "OrderID", Description = "ID of the order to modify")] string orderId,
            [ExcelArgument(Name = "Symbol", Description = "Trading symbol")] string symbol,
            [ExcelArgument(Name = "Action", Description = "Order action (BUY/SELL)")] string action,
            [ExcelArgument(Name = "Exchange", Description = "Exchange code(NSE/BSE/NFO/MCX)")] string exchange,
            [ExcelArgument(Name = "Quantity", Description = "New order quantity")] object? quantity = null,
            [ExcelArgument(Name = "PriceType", Description = "Price type (MARKET/LIMIT/SL/SL-M). Defaults to MARKET")] string priceType = "MARKET",
            [ExcelArgument(Name = "Product", Description = "Product type (MIS/CNC/NRML). Defaults to MIS")] string product = "MIS",
            [ExcelArgument(Name = "Price", Description = "New order price, sent as 0 when omitted")] object? price = null,
            [ExcelArgument(Name = "TriggerPrice", Description = "New trigger price, sent as 0 when omitted")] object? triggerPrice = null,
            [ExcelArgument(Name = "DisclosedQuantity", Description = "New disclosed quantity, sent as 0 when omitted")] object? disclosedQuantity = null)
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error(NoApiKeyMessage);
            if (!OpenAlgoConfig.TradingEnabled)
                return ExcelTable.Error(TradingDisabledMessage);

            string? missing = FirstMissing(
                ("Strategy", strategy), ("OrderID", orderId), ("Symbol", symbol),
                ("Action", action), ("Exchange", exchange));
            if (missing != null)
                return ExcelTable.Error(missing + " is required.");

            object[,]? quantityError = ReadPositiveQuantity(quantity, out double qty);
            if (quantityError != null)
                return quantityError;

            return AsyncTaskUtil.RunTask(nameof(oa_modifyorder), new object[]
            {
                strategy, orderId, symbol, action, exchange, priceType, product,
                Arg.Str(quantity), Arg.Str(price), Arg.Str(triggerPrice), Arg.Str(disclosedQuantity)
            }, async () =>
            {
                // Unlike placeorder, every field on modifyorder is mandatory. The server
                // schema rejects the request outright if price, trigger_price or
                // disclosed_quantity are absent, so an omitted value travels as 0.
                var payload = new JObject
                {
                    ["strategy"] = Arg.Str(strategy).Trim(),
                    ["orderid"] = orderId,
                    ["symbol"] = Norm(symbol),
                    ["action"] = Norm(action),
                    ["exchange"] = Norm(exchange),
                    ["quantity"] = qty,
                    ["pricetype"] = Norm(TextOr(priceType, "MARKET")),
                    ["product"] = Norm(TextOr(product, "MIS")),
                    ["price"] = Arg.Num(price),
                    ["trigger_price"] = Arg.Num(triggerPrice),
                    ["disclosed_quantity"] = Arg.Int(disclosedQuantity)
                };

                var response = await OpenAlgoClient.PostAsync("modifyorder", payload);

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                return (object)ExcelTable.Acknowledgement(response, "orderid");
            })!;
        }

        /// <summary>
        /// Cancels a single open order.
        /// </summary>
        [ExcelFunction(
            Name = "oa_cancelorder",
            Description = "Cancels an existing order through OpenAlgo API.",
            Category = "OpenAlgo Orders")]
        public static object oa_cancelorder(
            [ExcelArgument(Name = "Strategy", Description = "Trading strategy name")] string strategy,
            [ExcelArgument(Name = "OrderID", Description = "ID of the order to cancel")] string orderId)
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error(NoApiKeyMessage);
            if (!OpenAlgoConfig.TradingEnabled)
                return ExcelTable.Error(TradingDisabledMessage);

            string? missing = FirstMissing(("Strategy", strategy), ("OrderID", orderId));
            if (missing != null)
                return ExcelTable.Error(missing + " is required.");

            return AsyncTaskUtil.RunTask(nameof(oa_cancelorder), new object[] { strategy, orderId }, async () =>
            {
                var payload = new JObject
                {
                    ["strategy"] = Arg.Str(strategy).Trim(),
                    ["orderid"] = orderId
                };

                var response = await OpenAlgoClient.PostAsync("cancelorder", payload);

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                return (object)ExcelTable.Acknowledgement(response, "orderid");
            })!;
        }

        /// <summary>
        /// Cancels every open order and pending trigger order for a strategy.
        /// </summary>
        [ExcelFunction(
            Name = "oa_cancelallorder",
            Description = "Cancels all open orders for a specified strategy through OpenAlgo API.",
            Category = "OpenAlgo Orders")]
        public static object oa_cancelallorder(
            [ExcelArgument(Name = "Strategy", Description = "Trading strategy name")] string strategy)
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error(NoApiKeyMessage);
            if (!OpenAlgoConfig.TradingEnabled)
                return ExcelTable.Error(TradingDisabledMessage);

            if (string.IsNullOrWhiteSpace(strategy))
                return ExcelTable.Error("Strategy is required.");

            return AsyncTaskUtil.RunTask(nameof(oa_cancelallorder), new object[] { strategy }, async () =>
            {
                var payload = new JObject { ["strategy"] = Arg.Str(strategy).Trim() };

                var response = await OpenAlgoClient.PostAsync("cancelallorder", payload);

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                var rows = new List<object[]>
                {
                    new object[] { "Status", response["status"]?.ToString() ?? "unknown", "" }
                };

                string message = response["message"]?.ToString() ?? "";
                if (message.Length > 0)
                    rows.Add(new object[] { "Message", message, "" });

                var canceled = response["canceled_orders"] as JArray;
                var failed = response["failed_cancellations"] as JArray;

                if ((canceled?.Count ?? 0) > 0 || (failed?.Count ?? 0) > 0)
                {
                    rows.Add(new object[] { "Order ID", "Result", "Reason" });

                    if (canceled != null)
                        foreach (var item in canceled)
                            rows.Add(new object[] { ExcelTable.Cell(item), "Canceled", "" });

                    // Failed cancellations arrive as objects of orderid and reason.
                    if (failed != null)
                        foreach (var item in failed)
                        {
                            var entry = item as JObject;
                            rows.Add(new object[]
                            {
                                entry == null ? ExcelTable.Cell(item) : ExcelTable.Cell(entry["orderid"]),
                                "Failed",
                                entry == null ? "" : ExcelTable.Cell(entry["reason"])
                            });
                        }
                }

                return (object)ExcelTable.FromRows(rows);
            })!;
        }

        /// <summary>
        /// Squares off every open position for a strategy with market orders.
        /// </summary>
        [ExcelFunction(
            Name = "oa_closeposition",
            Description = "Closes all open positions for a specified strategy through OpenAlgo API.",
            Category = "OpenAlgo Orders")]
        public static object oa_closeposition(
            [ExcelArgument(Name = "Strategy", Description = "Trading strategy name")] string strategy)
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error(NoApiKeyMessage);
            if (!OpenAlgoConfig.TradingEnabled)
                return ExcelTable.Error(TradingDisabledMessage);

            if (string.IsNullOrWhiteSpace(strategy))
                return ExcelTable.Error("Strategy is required.");

            return AsyncTaskUtil.RunTask(nameof(oa_closeposition), new object[] { strategy }, async () =>
            {
                var payload = new JObject { ["strategy"] = Arg.Str(strategy).Trim() };

                var response = await OpenAlgoClient.PostAsync("closeposition", payload);

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                return (object)ExcelTable.Acknowledgement(response);
            })!;
        }

        /// <summary>
        /// Reads the current status of a single order. This function does not place or
        /// change anything, so it is not gated on the trading flag.
        /// </summary>
        [ExcelFunction(
            Name = "oa_orderstatus",
            Description = "Retrieves the status of a specific order through OpenAlgo API.",
            Category = "OpenAlgo Orders")]
        public static object oa_orderstatus(
            [ExcelArgument(Name = "Strategy", Description = "Trading strategy name associated with the order")] string strategy,
            [ExcelArgument(Name = "OrderID", Description = "ID of the order to retrieve status for")] string orderId)
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error(NoApiKeyMessage);

            if (string.IsNullOrWhiteSpace(orderId))
                return ExcelTable.Error("OrderID is required.");

            return AsyncTaskUtil.RunTask(nameof(oa_orderstatus), new object[] { strategy, orderId }, async () =>
            {
                var payload = new JObject
                {
                    ["strategy"] = Arg.Str(strategy).Trim(),
                    ["orderid"] = orderId
                };

                var response = await OpenAlgoClient.PostAsync("orderstatus", payload);

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                // The endpoint returns the order under "data", not "order_details".
                if (response["data"] is JObject details && details.Count > 0)
                    return (object)ExcelTable.KeyValue(details);

                return (object)ExcelTable.Acknowledgement(response, "orderid");
            })!;
        }

        /// <summary>
        /// Reads the net open quantity for one symbol, exchange and product. This function
        /// does not place or change anything, so it is not gated on the trading flag.
        /// </summary>
        [ExcelFunction(
            Name = "oa_openposition",
            Description = "Retrieves the net open position quantity for a symbol through OpenAlgo API.",
            Category = "OpenAlgo Orders")]
        public static object oa_openposition(
            [ExcelArgument(Name = "Strategy", Description = "Trading strategy name")] string strategy,
            [ExcelArgument(Name = "Symbol", Description = "Trading symbol")] string symbol,
            [ExcelArgument(Name = "Exchange", Description = "Exchange code(NSE/BSE/NFO/MCX)")] string exchange,
            [ExcelArgument(Name = "Product", Description = "Product type (MIS/CNC/NRML)")] string product)
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error(NoApiKeyMessage);

            string? missing = FirstMissing(("Symbol", symbol), ("Exchange", exchange), ("Product", product));
            if (missing != null)
                return ExcelTable.Error(missing + " is required.");

            return AsyncTaskUtil.RunTask(nameof(oa_openposition),
                new object[] { strategy, symbol, exchange, product }, async () =>
            {
                var payload = new JObject
                {
                    ["strategy"] = Arg.Str(strategy).Trim(),
                    ["symbol"] = Norm(symbol),
                    ["exchange"] = Norm(exchange),
                    ["product"] = Norm(product)
                };

                var response = await OpenAlgoClient.PostAsync("openposition", payload);

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                // The endpoint answers with a bare quantity, not a "position_details"
                // object. Hand it back as a single numeric cell so formulas can use it.
                var quantity = response["quantity"];
                if (quantity != null && quantity.Type != JTokenType.Null)
                    return (object)ExcelTable.Single(ExcelTable.Cell(quantity));

                return (object)ExcelTable.Acknowledgement(response);
            })!;
        }

        /// <summary>
        /// Returns the name of the first required argument that the caller left blank,
        /// or null when they are all present.
        /// </summary>
        private static string? FirstMissing(params (string Name, string? Value)[] required)
        {
            foreach (var (name, value) in required)
                if (string.IsNullOrWhiteSpace(value))
                    return name;
            return null;
        }

        /// <summary>
        /// Reads a mandatory, strictly positive quantity. Returns null when the value is
        /// usable, otherwise the error table to show in the cell.
        /// </summary>
        private static object[,]? ReadPositiveQuantity(object? value, out double quantity)
        {
            quantity = Arg.Num(value);
            if (Arg.IsMissing(value))
                return ExcelTable.Error("Quantity is required.");
            if (quantity <= 0)
                return ExcelTable.Error("Quantity must be greater than zero.");
            return null;
        }

        /// <summary>
        /// Returns the trimmed text, or the fallback when the caller left it blank.
        /// </summary>
        private static string TextOr(string? value, string fallback) =>
            string.IsNullOrWhiteSpace(value) ? fallback : value!.Trim();

        /// <summary>
        /// Adds a text field only when the caller supplied one, leaving the server to
        /// apply its own default otherwise.
        /// </summary>
        private static void AddText(JObject payload, string field, string? value)
        {
            if (!string.IsNullOrWhiteSpace(value))
                payload[field] = value!.Trim();
        }

        /// <summary>
        /// Adds an optional enum field, upper-cased. Used for pricetype and product,
        /// which the server validates against uppercase only sets, so a cell reading
        /// "limit" or "mis" would otherwise be rejected with HTTP 400.
        /// </summary>
        private static void AddEnum(JObject payload, string field, string? value)
        {
            if (!string.IsNullOrWhiteSpace(value))
                payload[field] = value!.Trim().ToUpperInvariant();
        }

        /// <summary>
        /// Adds a numeric field only when the caller supplied one. An omitted price is
        /// not the same request as a price of zero.
        /// </summary>
        private static void AddNumber(JObject payload, string field, object? value)
        {
            if (!Arg.IsMissing(value))
                payload[field] = Arg.Num(value);
        }

        /// <summary>
        /// Adds a whole number field only when the caller supplied one.
        /// </summary>
        private static void AddInteger(JObject payload, string field, object? value)
        {
            if (!Arg.IsMissing(value))
                payload[field] = Arg.Int(value);
        }

        /// <summary>
        /// Turns a worksheet range of order rows into the JSON array the basketorder
        /// endpoint expects. Returns null on success, otherwise the error table naming
        /// the offending row.
        /// </summary>
        private static object[,]? BuildBasket(object[,] orders, out JArray orderList)
        {
            orderList = new JArray();

            if (orders == null)
                return ExcelTable.Error("Orders range is required.");

            int rows = orders.GetLength(0);
            int cols = orders.GetLength(1);

            if (rows == 0 || cols < 4)
                return ExcelTable.Error(
                    "Orders range needs at least four columns: Symbol, Exchange, Action, Quantity.");

            int start = LooksLikeHeaderRow(orders) ? 1 : 0;

            for (int r = start; r < rows; r++)
            {
                string symbol = Arg.Str(orders[r, 0]).Trim();
                string exchange = Arg.Str(orders[r, 1]).Trim();
                string action = Arg.Str(orders[r, 2]).Trim();
                object? quantity = orders[r, 3];

                bool blank = symbol.Length == 0 && exchange.Length == 0
                             && action.Length == 0 && Arg.IsMissing(quantity);
                if (blank)
                    continue;

                int line = r + 1;
                string? missing = FirstMissing(
                    ("Symbol", symbol), ("Exchange", exchange), ("Action", action));
                if (missing != null)
                    return ExcelTable.Error($"Basket row {line}: {missing} is required.");
                if (Arg.IsMissing(quantity))
                    return ExcelTable.Error($"Basket row {line}: Quantity is required.");

                double qty = Arg.Num(quantity);
                if (qty <= 0)
                    return ExcelTable.Error($"Basket row {line}: Quantity must be greater than zero.");

                var order = new JObject
                {
                    ["symbol"] = Norm(symbol),
                    ["exchange"] = Norm(exchange),
                    ["action"] = Norm(action),
                    ["quantity"] = qty
                };

                if (cols > 4) AddEnum(order, "pricetype", Arg.Str(orders[r, 4]));
                if (cols > 5) AddEnum(order, "product", Arg.Str(orders[r, 5]));
                if (cols > 6) AddNumber(order, "price", orders[r, 6]);
                if (cols > 7) AddNumber(order, "trigger_price", orders[r, 7]);
                if (cols > 8) AddInteger(order, "disclosed_quantity", orders[r, 8]);

                orderList.Add(order);
            }

            if (orderList.Count == 0)
                return ExcelTable.Error("Orders range contains no order rows.");

            return null;
        }

        /// <summary>
        /// True when the first row of the range holds column captions rather than an order.
        /// </summary>
        private static bool LooksLikeHeaderRow(object[,] orders)
        {
            if (orders.GetLength(0) < 2)
                return false;

            string first = Arg.Str(orders[0, 0]).Trim();
            string third = Arg.Str(orders[0, 2]).Trim();
            string fourth = Arg.Str(orders[0, 3]).Trim();

            return first.Equals("symbol", StringComparison.OrdinalIgnoreCase)
                || third.Equals("action", StringComparison.OrdinalIgnoreCase)
                || fourth.Equals("quantity", StringComparison.OrdinalIgnoreCase)
                || fourth.Equals("qty", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Builds a stable identity string for a range argument so that two cells calling
        /// the same function with different ranges do not share one cached async result.
        /// </summary>
        private static string RangeKey(object[,]? grid)
        {
            if (grid == null)
                return "";

            int rows = grid.GetLength(0);
            int cols = grid.GetLength(1);

            var parts = new List<string>(rows * cols + 1) { rows + "x" + cols };
            for (int r = 0; r < rows; r++)
                for (int c = 0; c < cols; c++)
                    parts.Add(Arg.Str(grid[r, c]));

            return string.Join("|", parts);
        }
    }
}
