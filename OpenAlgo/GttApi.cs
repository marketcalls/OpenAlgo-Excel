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
    /// GTT (Good Till Triggered) worksheet functions.
    ///
    /// The OpenAlgo GTT endpoints support exactly two trigger shapes and no more:
    ///   SINGLE: one trigger, one child order. Exactly one of triggerprice_sl or
    ///           triggerprice_tg is positive, the other is 0, and "price" carries the
    ///           child order limit.
    ///   OCO:    two triggers, two child orders, one cancels the other. All four of
    ///           triggerprice_sl, stoploss, triggerprice_tg and target are positive,
    ///           "price" is ignored, and both legs share action, quantity and product.
    ///
    /// Because the leg count is fixed at one or two and every leg field has its own
    /// named request key, there is nothing a range driven multi leg signature could
    /// express that the flat argument list below cannot. One function therefore covers
    /// both trigger types: fill the two SINGLE trigger arguments for a SINGLE, or all
    /// four OCO arguments for an OCO.
    ///
    /// The three mutating functions are guarded. A worksheet function re-evaluates on
    /// every recalculation, so oa_placegttorder, oa_modifygttorder and oa_cancelgttorder
    /// refuse to fire while trading is switched off with oa_trading_enabled(FALSE). They
    /// are also non volatile and pass every argument into the async call identity, so two
    /// different GTT formulas can never share one cached result.
    /// </summary>
    public static class GttApi
    {
        private const string TradingDisabledMessage =
            "Trading is switched off. Call oa_trading_enabled(TRUE) to allow order functions to send.";

        /// <summary>
        /// Places a GTT trigger, either SINGLE (one trigger, one order) or OCO
        /// (stoploss trigger plus target trigger, one cancels the other).
        /// </summary>
        [ExcelFunction(
            Name = "oa_placegttorder",
            Description = "Places a GTT (Good Till Triggered) order. Trigger type SINGLE or OCO.",
            Category = "OpenAlgo GTT",
            IsVolatile = false)]
        public static object oa_placegttorder(
            [ExcelArgument(Name = "Strategy", Description = "Strategy identifier recorded with the trigger")] string strategy,
            [ExcelArgument(Name = "Symbol", Description = "Trading symbol in OpenAlgo format, for example RELIANCE or NIFTY25AUG26FUT")] string symbol,
            [ExcelArgument(Name = "Exchange", Description = "Exchange code: NSE, BSE, NFO, BFO, CDS, BCD or MCX")] string exchange,
            [ExcelArgument(Name = "Action", Description = "BUY or SELL. For OCO it applies to both legs")] string action,
            [ExcelArgument(Name = "Product", Description = "CNC (equity delivery) or NRML (overnight F and O). MIS is not supported for GTT")] string product,
            [ExcelArgument(Name = "TriggerType", Description = "SINGLE for one trigger, OCO for a stoploss plus target bracket")] string triggerType,
            [ExcelArgument(Name = "Quantity", Description = "Order quantity. Whole number for equity and F and O")] object quantity,
            [ExcelArgument(Name = "PriceType", Description = "LIMIT or MARKET. Optional, defaults to LIMIT")] object? priceType = null,
            [ExcelArgument(Name = "Price", Description = "SINGLE only: limit price of the child order. Use 0 when PriceType is MARKET. Ignored for OCO")] object? price = null,
            [ExcelArgument(Name = "TriggerPriceSL", Description = "Trigger price below LTP. SINGLE: this or TriggerPriceTG. OCO: the stoploss leg trigger")] object? triggerPriceSl = null,
            [ExcelArgument(Name = "TriggerPriceTG", Description = "Trigger price above LTP. SINGLE: this or TriggerPriceSL. OCO: the target leg trigger")] object? triggerPriceTg = null,
            [ExcelArgument(Name = "StopLoss", Description = "OCO only: limit price of the stoploss leg child order. Ignored for SINGLE")] object? stopLoss = null,
            [ExcelArgument(Name = "Target", Description = "OCO only: limit price of the target leg child order. Ignored for SINGLE")] object? target = null)
        {
            if (!OpenAlgoConfig.TradingEnabled)
                return ExcelTable.Error(TradingDisabledMessage);

            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");

            var spec = GttSpec.Read(strategy, symbol, exchange, action, product, triggerType,
                                    quantity, priceType, price, triggerPriceSl, triggerPriceTg, stopLoss, target);

            string? invalid = spec.Validate();
            if (invalid != null)
                return ExcelTable.Error(invalid);

            return AsyncTaskUtil.RunTask(nameof(oa_placegttorder), spec.Identity(), async () =>
            {
                var response = await OpenAlgoClient.PostAsync("placegttorder", spec.Payload());

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                return (object)ExcelTable.Acknowledgement(response, "trigger_id", "orderid");
            })!;
        }

        /// <summary>
        /// Modifies an active GTT trigger. The request is a full replacement of the
        /// trigger specification, so every field that should survive must be supplied.
        /// </summary>
        [ExcelFunction(
            Name = "oa_modifygttorder",
            Description = "Modifies an active GTT order. Send every field to keep: modify replaces the whole trigger.",
            Category = "OpenAlgo GTT",
            IsVolatile = false)]
        public static object oa_modifygttorder(
            [ExcelArgument(Name = "Strategy", Description = "Strategy identifier recorded with the trigger")] string strategy,
            [ExcelArgument(Name = "TriggerID", Description = "Trigger ID returned by oa_placegttorder")] string triggerId,
            [ExcelArgument(Name = "Symbol", Description = "Trading symbol in OpenAlgo format. Cannot be changed from the original")] string symbol,
            [ExcelArgument(Name = "Exchange", Description = "Exchange code: NSE, BSE, NFO, BFO, CDS, BCD or MCX. Cannot be changed from the original")] string exchange,
            [ExcelArgument(Name = "Action", Description = "BUY or SELL. Cannot be changed from the original")] string action,
            [ExcelArgument(Name = "Product", Description = "CNC (equity delivery) or NRML (overnight F and O). MIS is not supported for GTT")] string product,
            [ExcelArgument(Name = "TriggerType", Description = "SINGLE or OCO. Must match the original trigger. Cancel and re-place to switch")] string triggerType,
            [ExcelArgument(Name = "Quantity", Description = "New order quantity. Must be a valid lot size for F and O")] object quantity,
            [ExcelArgument(Name = "PriceType", Description = "LIMIT or MARKET. Optional, defaults to LIMIT")] object? priceType = null,
            [ExcelArgument(Name = "Price", Description = "SINGLE only: new limit price of the child order. Use 0 when PriceType is MARKET. Ignored for OCO")] object? price = null,
            [ExcelArgument(Name = "TriggerPriceSL", Description = "New trigger price below LTP. SINGLE: this or TriggerPriceTG. OCO: the stoploss leg trigger")] object? triggerPriceSl = null,
            [ExcelArgument(Name = "TriggerPriceTG", Description = "New trigger price above LTP. SINGLE: this or TriggerPriceSL. OCO: the target leg trigger")] object? triggerPriceTg = null,
            [ExcelArgument(Name = "StopLoss", Description = "OCO only: new limit price of the stoploss leg child order. Ignored for SINGLE")] object? stopLoss = null,
            [ExcelArgument(Name = "Target", Description = "OCO only: new limit price of the target leg child order. Ignored for SINGLE")] object? target = null)
        {
            if (!OpenAlgoConfig.TradingEnabled)
                return ExcelTable.Error(TradingDisabledMessage);

            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");

            string id = Arg.Str(triggerId).Trim();
            if (id.Length == 0)
                return ExcelTable.Error("TriggerID is required. Use the trigger id returned by oa_placegttorder.");

            var spec = GttSpec.Read(strategy, symbol, exchange, action, product, triggerType,
                                    quantity, priceType, price, triggerPriceSl, triggerPriceTg, stopLoss, target);

            string? invalid = spec.Validate();
            if (invalid != null)
                return ExcelTable.Error(invalid);

            var identity = new List<object>(spec.Identity()) { id }.ToArray();

            return AsyncTaskUtil.RunTask(nameof(oa_modifygttorder), identity, async () =>
            {
                var payload = spec.Payload();
                payload["trigger_id"] = id;

                var response = await OpenAlgoClient.PostAsync("modifygttorder", payload);

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                return (object)ExcelTable.Acknowledgement(response, "trigger_id", "orderid");
            })!;
        }

        /// <summary>
        /// Cancels an active GTT trigger. Cancelling an OCO removes both legs together.
        /// </summary>
        [ExcelFunction(
            Name = "oa_cancelgttorder",
            Description = "Cancels an active GTT order by trigger id. An OCO loses both legs.",
            Category = "OpenAlgo GTT",
            IsVolatile = false)]
        public static object oa_cancelgttorder(
            [ExcelArgument(Name = "Strategy", Description = "Strategy identifier recorded in the event log")] string strategy,
            [ExcelArgument(Name = "TriggerID", Description = "Active trigger ID returned by oa_placegttorder")] string triggerId)
        {
            if (!OpenAlgoConfig.TradingEnabled)
                return ExcelTable.Error(TradingDisabledMessage);

            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");

            string strategyName = Arg.Str(strategy).Trim();
            string id = Arg.Str(triggerId).Trim();

            if (strategyName.Length == 0)
                return ExcelTable.Error("Strategy is required.");
            if (id.Length == 0)
                return ExcelTable.Error("TriggerID is required. Use the trigger id returned by oa_placegttorder.");

            return AsyncTaskUtil.RunTask(nameof(oa_cancelgttorder), new object[] { strategyName, id }, async () =>
            {
                var payload = new JObject
                {
                    ["strategy"] = strategyName,
                    ["trigger_id"] = id
                };

                var response = await OpenAlgoClient.PostAsync("cancelgttorder", payload);

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                return (object)ExcelTable.Acknowledgement(response, "trigger_id", "orderid");
            })!;
        }

        /// <summary>
        /// Retrieves the active GTT order book as a table with a header row.
        /// </summary>
        [ExcelFunction(
            Name = "oa_gttorderbook",
            Description = "Retrieves the active GTT order book, one row per trigger with its legs flattened.",
            Category = "OpenAlgo GTT")]
        public static object oa_gttorderbook()
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");

            return AsyncTaskUtil.RunTask(nameof(oa_gttorderbook), new object[] { }, async () =>
            {
                var response = await OpenAlgoClient.PostAsync("gttorderbook", new JObject());

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                if (response["data"] is not JArray rules)
                    return (object)ExcelTable.Error("No GTT orders returned");

                return (object)BuildOrderBook(rules);
            })!;
        }

        /// <summary>
        /// Flattens the GTT order book into one Excel row per trigger.
        ///
        /// Each entry nests two parallel arrays: trigger_prices and legs, documented as
        /// sorted ascending and index aligned (for an OCO, index 0 is the stoploss leg
        /// and index 1 is the target leg). Nesting cannot be shown in a single cell, so
        /// each index becomes its own block of columns: trigger price followed by that
        /// leg's action, quantity, price, price type and product. Every documented field
        /// of both the entry and the leg object is therefore present in the row.
        ///
        /// The number of leg blocks is taken from the widest entry in the payload, so a
        /// book of SINGLE triggers stays narrow and unlabelled while a book containing an
        /// OCO gains "Leg 1" and "Leg 2" prefixes. Deriving the count from the data rather
        /// than hardcoding two also keeps any future broker with more legs readable.
        /// </summary>
        private static object[,] BuildOrderBook(JArray rules)
        {
            if (rules.Count == 0)
                return ExcelTable.Single("No active GTT orders");

            // Widest entry decides how many leg blocks the table carries.
            int legBlocks = 1;
            foreach (var rule in rules.OfType<JObject>())
            {
                int legCount = (rule["legs"] as JArray)?.Count ?? 0;
                int triggerCount = (rule["trigger_prices"] as JArray)?.Count ?? 0;
                legBlocks = Math.Max(legBlocks, Math.Max(legCount, triggerCount));
            }

            var headers = new List<object>
            {
                "Trigger ID", "Trigger Type", "Status", "Symbol", "Exchange", "Last Price"
            };

            for (int leg = 0; leg < legBlocks; leg++)
            {
                // A one leg book needs no leg numbering, which keeps SINGLE books readable.
                string prefix = legBlocks == 1
                    ? ""
                    : "Leg " + (leg + 1).ToString(CultureInfo.InvariantCulture) + " ";
                headers.Add(prefix + "Trigger Price");
                headers.Add(prefix + "Action");
                headers.Add(prefix + "Quantity");
                headers.Add(prefix + "Price");
                headers.Add(prefix + "Price Type");
                headers.Add(prefix + "Product");
            }

            headers.Add("Created At");
            headers.Add("Updated At");
            headers.Add("Expires At");

            var rows = new List<object[]> { headers.ToArray() };

            foreach (var token in rules)
            {
                var rule = token as JObject;
                if (rule == null)
                    continue;

                var triggerPrices = rule["trigger_prices"] as JArray;
                var legs = rule["legs"] as JArray;

                var row = new List<object>
                {
                    ExcelTable.Cell(rule["trigger_id"]),
                    ExcelTable.Cell(rule["trigger_type"]),
                    ExcelTable.Cell(rule["status"]),
                    ExcelTable.Cell(rule["symbol"]),
                    ExcelTable.Cell(rule["exchange"]),
                    ExcelTable.Cell(rule["last_price"])
                };

                for (int leg = 0; leg < legBlocks; leg++)
                {
                    var legPrice = triggerPrices != null && leg < triggerPrices.Count ? triggerPrices[leg] : null;
                    var legDetail = legs != null && leg < legs.Count ? legs[leg] as JObject : null;

                    row.Add(ExcelTable.Cell(legPrice));
                    row.Add(ExcelTable.Cell(legDetail?["action"]));
                    row.Add(ExcelTable.Cell(legDetail?["quantity"]));
                    row.Add(ExcelTable.Cell(legDetail?["price"]));
                    row.Add(ExcelTable.Cell(legDetail?["pricetype"]));
                    row.Add(ExcelTable.Cell(legDetail?["product"]));
                }

                row.Add(ExcelTable.Cell(rule["created_at"]));
                row.Add(ExcelTable.Cell(rule["updated_at"]));
                row.Add(ExcelTable.Cell(rule["expires_at"]));

                rows.Add(row.ToArray());
            }

            return ExcelTable.FromRows(rows);
        }

        /// <summary>
        /// A validated GTT trigger specification shared by place and modify, which take
        /// the same body apart from the trigger id.
        /// </summary>
        private sealed class GttSpec
        {
            public string Strategy = "";
            public string Symbol = "";
            public string Exchange = "";
            public string Action = "";
            public string Product = "";
            public string TriggerType = "";
            public string PriceType = "";
            public double Quantity;
            public double Price;
            public double TriggerPriceSl;
            public double TriggerPriceTg;
            public double StopLoss;
            public double Target;

            public bool IsOco => TriggerType == "OCO";

            /// <summary>
            /// Normalises the loosely typed Excel arguments into a specification.
            /// </summary>
            public static GttSpec Read(
                string strategy, string symbol, string exchange, string action, string product,
                string triggerType, object? quantity, object? priceType, object? price,
                object? triggerPriceSl, object? triggerPriceTg, object? stopLoss, object? target)
            {
                return new GttSpec
                {
                    Strategy = Arg.Str(strategy).Trim(),
                    Symbol = Arg.Str(symbol).Trim().ToUpperInvariant(),
                    Exchange = Arg.Str(exchange).Trim().ToUpperInvariant(),
                    Action = Arg.Str(action).Trim().ToUpperInvariant(),
                    Product = Arg.Str(product).Trim().ToUpperInvariant(),
                    TriggerType = Arg.Str(triggerType).Trim().ToUpperInvariant(),
                    PriceType = Arg.Str(priceType, "LIMIT").Trim().ToUpperInvariant(),
                    Quantity = Arg.Num(quantity),
                    Price = Arg.Num(price),
                    TriggerPriceSl = Arg.Num(triggerPriceSl),
                    TriggerPriceTg = Arg.Num(triggerPriceTg),
                    StopLoss = Arg.Num(stopLoss),
                    Target = Arg.Num(target)
                };
            }

            /// <summary>
            /// Returns the first problem with the specification, or null when it is usable.
            /// Catching these locally saves a round trip and gives the cell a clear message.
            /// </summary>
            public string? Validate()
            {
                if (Strategy.Length == 0)
                    return "Strategy is required.";
                if (Symbol.Length == 0)
                    return "Symbol is required.";
                if (Exchange.Length == 0)
                    return "Exchange is required.";
                if (Action != "BUY" && Action != "SELL")
                    return "Action must be BUY or SELL.";
                if (Product == "MIS")
                    return "GTT supports only CNC (delivery) or NRML (overnight F and O). MIS is intraday only.";
                if (Product.Length == 0)
                    return "Product is required. Use CNC or NRML.";
                if (TriggerType != "SINGLE" && TriggerType != "OCO")
                    return "TriggerType must be SINGLE or OCO.";
                if (PriceType != "LIMIT" && PriceType != "MARKET")
                    return "PriceType must be LIMIT or MARKET.";
                if (Quantity <= 0)
                    return "Quantity must be a positive number.";

                if (IsOco)
                {
                    if (TriggerPriceSl <= 0 || StopLoss <= 0 || TriggerPriceTg <= 0 || Target <= 0)
                        return "OCO requires TriggerPriceSL, StopLoss, TriggerPriceTG and Target, all greater than zero.";
                    if (TriggerPriceSl >= TriggerPriceTg)
                        return "Stoploss trigger must be less than target trigger.";
                }
                else
                {
                    if (TriggerPriceSl <= 0 && TriggerPriceTg <= 0)
                        return "SINGLE GTT requires a positive TriggerPriceSL (trigger below LTP) or TriggerPriceTG (trigger above LTP).";
                    if (TriggerPriceSl > 0 && TriggerPriceTg > 0)
                        return "SINGLE GTT takes exactly one trigger. Leave TriggerPriceSL or TriggerPriceTG at zero.";
                    if (PriceType == "LIMIT" && Price <= 0)
                        return "SINGLE GTT with PriceType LIMIT requires a positive Price for the child order.";
                }

                return null;
            }

            /// <summary>
            /// Builds the request body. The API key is added by OpenAlgoClient.
            /// </summary>
            public JObject Payload()
            {
                var payload = new JObject
                {
                    ["strategy"] = Strategy,
                    ["trigger_type"] = TriggerType,
                    ["exchange"] = Exchange,
                    ["symbol"] = Symbol,
                    ["action"] = Action,
                    ["product"] = Product,
                    ["quantity"] = Quantity,
                    ["pricetype"] = PriceType
                };

                if (IsOco)
                {
                    // OCO ignores the shared price and uses a limit price per leg.
                    payload["price"] = 0;
                    payload["triggerprice_sl"] = TriggerPriceSl;
                    payload["stoploss"] = StopLoss;
                    payload["triggerprice_tg"] = TriggerPriceTg;
                    payload["target"] = Target;
                }
                else
                {
                    // SINGLE has no second leg, so the leg limit prices go out as null.
                    payload["price"] = PriceType == "MARKET" ? 0 : Price;
                    payload["triggerprice_sl"] = TriggerPriceSl;
                    payload["triggerprice_tg"] = TriggerPriceTg;
                    payload["stoploss"] = JValue.CreateNull();
                    payload["target"] = JValue.CreateNull();
                }

                return payload;
            }

            /// <summary>
            /// Every argument, flattened to simple values, so that two GTT formulas with
            /// different arguments never share one cached async result.
            /// </summary>
            public object[] Identity() => new object[]
            {
                Strategy, Symbol, Exchange, Action, Product, TriggerType, PriceType,
                Quantity, Price, TriggerPriceSl, TriggerPriceTg, StopLoss, Target
            };
        }
    }
}
