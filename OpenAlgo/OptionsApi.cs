using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelDna.Integration;
using ExcelDna.Registration.Utils;
using Newtonsoft.Json.Linq;

namespace OpenAlgo
{
    /// <summary>
    /// Worksheet functions for the OpenAlgo options services.
    ///
    /// Analytics (symbol resolution, option chain, Greeks and the synthetic future)
    /// are read only and safe to leave on a sheet. The two order placing functions
    /// are guarded by OpenAlgoConfig.TradingEnabled because a worksheet function
    /// fires again on every recalculation.
    ///
    /// Every table keeps the fields the endpoint documents and appends any further
    /// field the server sends, so a payload that grows never loses data on the way
    /// into Excel.
    /// </summary>
    public static class OptionsApi
    {
        private const string OptionsCategory = "OpenAlgo Options";
        private const string OrdersCategory = "OpenAlgo Options Orders";

        private const string NoApiKey = "OpenAlgo API Key is not set. Use oa_api()";
        private const string TradingDisabledMessage =
            "Trading is disabled. Call oa_trading_enabled(TRUE) to arm order functions.";

        /// <summary>
        /// Properties that never belong in a key/value table because they are transport
        /// level or are expanded separately.
        /// </summary>
        private static readonly string[] Hidden = { "status", "apikey", "message", "error", "greeks", "data" };

        private static readonly Regex OffsetPattern =
            new Regex(@"^(ATM|(ITM|OTM)([1-9]|[1-4][0-9]|50))$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex ClockTimePattern =
            new Regex(@"^([01]?[0-9]|2[0-3]):[0-5][0-9]$", RegexOptions.Compiled);

        /// <summary>
        /// Captions for the option chain and Greek columns. Anything not listed here
        /// falls back to ExcelTable.Humanise.
        /// </summary>
        private static readonly Dictionary<string, string> Captions = new(StringComparer.OrdinalIgnoreCase)
        {
            ["ltp"] = "LTP",
            ["oi"] = "OI",
            ["oi_change"] = "OI Chg",
            ["change_oi"] = "OI Chg",
            ["bid_qty"] = "Bid Qty",
            ["ask_qty"] = "Ask Qty",
            ["prev_close"] = "Prev Close",
            ["lotsize"] = "Lot Size",
            ["tick_size"] = "Tick Size",
            ["implied_volatility"] = "IV",
            ["iv"] = "IV"
        };

        // ------------------------------------------------------------------
        // Analytics
        // ------------------------------------------------------------------

        /// <summary>
        /// Resolves an option contract from an underlying, expiry and strike offset.
        /// The resolved trading symbol is the first data row so another formula can
        /// pick it up with INDEX(range, 2, 2).
        /// </summary>
        [ExcelFunction(
            Name = "oa_optionsymbol",
            Description = "Resolves an option trading symbol from underlying, expiry and strike offset.",
            Category = OptionsCategory)]
        public static object oa_optionsymbol(
            [ExcelArgument(Name = "Underlying", Description = "Underlying symbol, for example NIFTY, BANKNIFTY or SENSEX")] string underlying,
            [ExcelArgument(Name = "Exchange", Description = "Underlying exchange: NSE_INDEX or BSE_INDEX")] string exchange,
            [ExcelArgument(Name = "Expiry", Description = "Expiry date in DDMMMYY format, for example 25AUG26")] string expiry,
            [ExcelArgument(Name = "StrikeOffset", Description = "Strike offset: ATM, ITM1 to ITM50, or OTM1 to OTM50")] string strikeOffset,
            [ExcelArgument(Name = "OptionType", Description = "CE for a call or PE for a put")] string optionType)
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error(NoApiKey);
            if (string.IsNullOrWhiteSpace(underlying) || string.IsNullOrWhiteSpace(exchange) || string.IsNullOrWhiteSpace(expiry))
                return ExcelTable.Error("Underlying, Exchange and Expiry are required");

            string offset = NormaliseOffset(strikeOffset, out string? offsetError);
            if (offsetError != null)
                return ExcelTable.Error(offsetError);
            string type = NormaliseOptionType(optionType, out string? typeError);
            if (typeError != null)
                return ExcelTable.Error(typeError);

            string expiryDate = expiry.Trim().ToUpperInvariant();

            return AsyncTaskUtil.RunTask(
                nameof(oa_optionsymbol),
                new object[] { underlying, exchange, expiryDate, offset, type },
                async () =>
                {
                    var payload = new JObject
                    {
                        ["underlying"] = underlying,
                        ["exchange"] = exchange,
                        ["expiry_date"] = expiryDate,
                        ["offset"] = offset,
                        ["option_type"] = type
                    };

                    var response = await OpenAlgoClient.PostAsync("optionsymbol", payload);

                    var error = ExcelTable.ErrorOrNull(response);
                    if (error != null) return (object)error;

                    var data = Body(response);
                    if (data["symbol"] == null)
                        return (object)ExcelTable.Error("No option symbol returned");

                    // Symbol first so INDEX(range, 2, 2) always reads the trading symbol.
                    var pairs = Pairs(data, new[]
                    {
                        "symbol", "exchange", "strike", "option_type", "expiry_date",
                        "lotsize", "tick_size", "freeze_qty", "underlying", "underlying_ltp"
                    });

                    return (object)ExcelTable.KeyValue(pairs, "Field", "Value");
                })!;
        }

        /// <summary>
        /// Returns the full option chain for an expiry as a ladder.
        ///
        /// Row one carries the underlying context: underlying, spot, previous close,
        /// ATM strike, expiry, expiry and server timestamps, forward price and the
        /// quotes and Greeks flags. Row two carries the column headers. Every data row
        /// is one strike, with the call leg on the left, the strike in the middle and
        /// the put leg on the right, the two sides mirroring each other around the
        /// strike. Greek columns appear when the payload carries Greeks, and any
        /// further leg field the server sends is appended at the outer edge of each side.
        /// </summary>
        [ExcelFunction(
            Name = "oa_optionchain",
            Description = "Option chain ladder for an expiry: calls, strike, puts, with full quotes and Greeks.",
            Category = OptionsCategory)]
        public static object oa_optionchain(
            [ExcelArgument(Name = "Underlying", Description = "Underlying symbol, for example NIFTY, BANKNIFTY or SENSEX")] string underlying,
            [ExcelArgument(Name = "Exchange", Description = "Underlying exchange: NSE_INDEX or BSE_INDEX")] string exchange,
            [ExcelArgument(Name = "Expiry", Description = "Expiry date in DDMMMYY format, for example 25AUG26")] string expiry,
            [ExcelArgument(Name = "StrikeCount", Description = "Strikes above and below ATM, 1 to 100 (optional, default all strikes)")] object? strikeCount = null,
            [ExcelArgument(Name = "WithGreeks", Description = "TRUE to attach IV and Greeks to every leg (optional, default TRUE, costs no extra broker call)")] object? withGreeks = null,
            [ExcelArgument(Name = "InterestRate", Description = "Risk free rate as an annualised percent, 0 to 100, Greeks only (optional, default 0)")] object? interestRate = null)
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error(NoApiKey);
            if (string.IsNullOrWhiteSpace(underlying) || string.IsNullOrWhiteSpace(exchange) || string.IsNullOrWhiteSpace(expiry))
                return ExcelTable.Error("Underlying, Exchange and Expiry are required");

            int strikes = Arg.IsMissing(strikeCount) ? 0 : Arg.Int(strikeCount);
            if (!Arg.IsMissing(strikeCount) && (strikes < 1 || strikes > 100))
                return ExcelTable.Error("StrikeCount must be between 1 and 100");

            bool greeksRequested = Arg.Bool(withGreeks, true);

            double rate = Arg.Num(interestRate);
            if (!Arg.IsMissing(interestRate) && (rate < 0 || rate > 100))
                return ExcelTable.Error("InterestRate must be between 0 and 100");
            bool rateSupplied = !Arg.IsMissing(interestRate);

            string expiryDate = expiry.Trim().ToUpperInvariant();

            return AsyncTaskUtil.RunTask(
                nameof(oa_optionchain),
                new object[] { underlying, exchange, expiryDate, strikes, greeksRequested, rateSupplied ? rate : double.NaN },
                async () =>
                {
                    var payload = new JObject
                    {
                        ["underlying"] = underlying,
                        ["exchange"] = exchange,
                        ["expiry_date"] = expiryDate,
                        ["with_greeks"] = greeksRequested
                    };
                    if (strikes > 0)
                        payload["strike_count"] = strikes;
                    if (rateSupplied)
                        payload["interest_rate"] = rate;

                    var response = await OpenAlgoClient.PostAsync("optionchain", payload);

                    var error = ExcelTable.ErrorOrNull(response);
                    if (error != null) return (object)error;

                    var data = Body(response);
                    if (data["chain"] is not JArray chain || chain.Count == 0)
                        return (object)ExcelTable.Error("No option chain returned");

                    bool greeks = data["greeks_included"]?.ToObject<bool?>() ?? greeksRequested;

                    // Put side, read outwards from the strike. The call side is the
                    // mirror image of this sequence.
                    var putFields = new List<string> { "ltp", "bid", "bid_qty", "ask", "ask_qty" };
                    if (greeks)
                        putFields.AddRange(new[] { "implied_volatility", "delta", "gamma", "theta", "vega" });
                    putFields.AddRange(new[]
                    {
                        "volume", "oi", "open", "high", "low", "prev_close", "lotsize", "tick_size", "label", "symbol"
                    });

                    // Keep anything the payload carries that is not in the documented set,
                    // for example a change in open interest, rather than dropping it.
                    var legs = new List<JObject>();
                    foreach (var entry in chain.OfType<JObject>())
                    {
                        if (entry["ce"] is JObject call) legs.Add(call);
                        if (entry["pe"] is JObject put) legs.Add(put);
                    }
                    putFields.AddRange(ExtraFields(legs, putFields));

                    var callFields = new List<string>(putFields);
                    callFields.Reverse();

                    int width = callFields.Count + 1 + putFields.Count;
                    var rows = new List<object[]>();

                    // Row one: underlying context for the whole ladder.
                    var info = new object[width];
                    for (int i = 0; i < width; i++)
                        info[i] = "";
                    void Put(int index, object value)
                    {
                        if (index < width) info[index] = value;
                    }
                    Put(0, "Underlying");
                    Put(1, ExcelTable.Cell(data["underlying"]));
                    Put(2, "Spot");
                    Put(3, ExcelTable.Cell(data["underlying_ltp"]));
                    Put(4, "Prev Close");
                    Put(5, ExcelTable.Cell(data["underlying_prev_close"]));
                    Put(6, "ATM Strike");
                    Put(7, ExcelTable.Cell(data["atm_strike"]));
                    Put(8, "Expiry");
                    Put(9, ExcelTable.Cell(data["expiry_date"]));
                    Put(10, "Expiry TS");
                    Put(11, ExcelTable.Cell(data["expiry_ts"]));
                    Put(12, "Server TS");
                    Put(13, ExcelTable.Cell(data["server_ts"]));
                    Put(14, "Forward");
                    Put(15, ExcelTable.Cell(data["forward_price"]));
                    Put(16, "Quotes Included");
                    Put(17, ExcelTable.Cell(data["quotes_included"]));
                    Put(18, "Greeks Included");
                    Put(19, ExcelTable.Cell(data["greeks_included"]));
                    Put(20, "Strikes");
                    Put(21, (double)chain.Count);
                    rows.Add(info);

                    // Row two: column headers.
                    var header = new object[width];
                    int column = 0;
                    foreach (var field in callFields)
                        header[column++] = "CE " + Caption(field);
                    header[column++] = "Strike";
                    foreach (var field in putFields)
                        header[column++] = "PE " + Caption(field);
                    rows.Add(header);

                    foreach (var entry in chain.OfType<JObject>())
                    {
                        var call = entry["ce"] as JObject;
                        var put = entry["pe"] as JObject;

                        var row = new object[width];
                        int cell = 0;
                        foreach (var field in callFields)
                            row[cell++] = call == null ? "" : ExcelTable.Cell(call[field]);
                        row[cell++] = ExcelTable.Cell(entry["strike"]);
                        foreach (var field in putFields)
                            row[cell++] = put == null ? "" : ExcelTable.Cell(put[field]);
                        rows.Add(row);
                    }

                    return (object)ExcelTable.FromRows(rows);
                })!;
        }

        /// <summary>
        /// Black-76 Greeks and implied volatility for a single option contract.
        /// </summary>
        [ExcelFunction(
            Name = "oa_optiongreeks",
            Description = "Black-76 Greeks and implied volatility for an option.",
            Category = OptionsCategory)]
        public static object oa_optiongreeks(
            [ExcelArgument(Name = "Symbol", Description = "Option trading symbol, for example NIFTY25AUG2625000CE")] string symbol,
            [ExcelArgument(Name = "Exchange", Description = "NFO, BFO, CDS, MCX or CRYPTO")] string exchange,
            [ExcelArgument(Name = "InterestRate", Description = "Risk free rate as an annualised percent, 0 to 100 (optional)")] object? interestRate = null,
            [ExcelArgument(Name = "ExpiryTime", Description = "Custom expiry time as HH:MM, for example 15:30 (optional)")] object? expiryTime = null,
            [ExcelArgument(Name = "UnderlyingSymbol", Description = "Underlying symbol to use for the spot price (optional, derived from the option)")] object? underlyingSymbol = null,
            [ExcelArgument(Name = "UnderlyingExchange", Description = "Underlying exchange (optional, default NSE_INDEX)")] object? underlyingExchange = null,
            [ExcelArgument(Name = "ForwardPrice", Description = "Custom forward or synthetic futures price, bypasses forward resolution (optional)")] object? forwardPrice = null)
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error(NoApiKey);
            if (string.IsNullOrWhiteSpace(symbol) || string.IsNullOrWhiteSpace(exchange))
                return ExcelTable.Error("Symbol and Exchange are required");

            double rate = Arg.Num(interestRate);
            if (!Arg.IsMissing(interestRate) && (rate < 0 || rate > 100))
                return ExcelTable.Error("InterestRate must be between 0 and 100");
            bool rateSupplied = !Arg.IsMissing(interestRate);

            string time = Arg.Str(expiryTime).Trim();
            if (time.Length > 0 && !ClockTimePattern.IsMatch(time))
                return ExcelTable.Error("ExpiryTime must be in HH:MM format, for example 15:30");

            string spotSymbol = Arg.Str(underlyingSymbol).Trim();
            string spotExchange = Arg.Str(underlyingExchange).Trim();
            double forward = Arg.Num(forwardPrice);
            bool forwardSupplied = !Arg.IsMissing(forwardPrice);

            return AsyncTaskUtil.RunTask(
                nameof(oa_optiongreeks),
                new object[]
                {
                    symbol, exchange, rateSupplied ? rate : double.NaN, time,
                    spotSymbol, spotExchange, forwardSupplied ? forward : double.NaN
                },
                async () =>
                {
                    var payload = new JObject
                    {
                        ["symbol"] = symbol,
                        ["exchange"] = exchange
                    };
                    if (rateSupplied)
                        payload["interest_rate"] = rate;
                    if (time.Length > 0)
                        payload["expiry_time"] = time;
                    if (spotSymbol.Length > 0)
                        payload["underlying_symbol"] = spotSymbol;
                    if (spotExchange.Length > 0)
                        payload["underlying_exchange"] = spotExchange;
                    if (forwardSupplied)
                        payload["forward_price"] = forward;

                    var response = await OpenAlgoClient.PostAsync("optiongreeks", payload);

                    var error = ExcelTable.ErrorOrNull(response);
                    if (error != null) return (object)error;

                    var data = Body(response);
                    if (data.Count == 0)
                        return (object)ExcelTable.Error("No Greeks returned");

                    var pairs = Pairs(data, new[]
                    {
                        "symbol", "exchange", "underlying", "strike", "option_type", "expiry_date",
                        "days_to_expiry", "spot_price", "forward_price", "option_price",
                        "interest_rate", "implied_volatility"
                    }, flattenGreeks: true);

                    return (object)ExcelTable.KeyValue(pairs, "Field", "Value");
                })!;
        }

        /// <summary>
        /// Black-76 Greeks and implied volatility for 1 to 50 option contracts in one call.
        ///
        /// Row one is the batch summary, row two the column headers, then one row per
        /// contract. Individual contracts can fail while the batch reports success, so
        /// every row carries its own status and error cell.
        /// </summary>
        [ExcelFunction(
            Name = "oa_multioptiongreeks",
            Description = "Black-76 Greeks and implied volatility for up to 50 options in one call.",
            Category = OptionsCategory)]
        public static object oa_multioptiongreeks(
            [ExcelArgument(Name = "Symbols", Description = "Range of option symbols in one column, symbol and exchange in two, or symbol, exchange, underlying symbol and underlying exchange in four")] object symbols,
            [ExcelArgument(Name = "InterestRate", Description = "Common risk free rate as an annualised percent, 0 to 100 (optional)")] object? interestRate = null,
            [ExcelArgument(Name = "ExpiryTime", Description = "Common expiry time as HH:MM, for example 15:30 (optional)")] object? expiryTime = null)
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error(NoApiKey);

            var requested = Arg.SymbolPairs(symbols, "NFO");
            if (requested.Count == 0)
                return ExcelTable.Error("Symbols is empty. Point it at a column of option symbols");
            if (requested.Count > 50)
                return ExcelTable.Error($"MultiOptionGreeks accepts at most 50 symbols, {requested.Count} were supplied");

            double rate = Arg.Num(interestRate);
            if (!Arg.IsMissing(interestRate) && (rate < 0 || rate > 100))
                return ExcelTable.Error("InterestRate must be between 0 and 100");
            bool rateSupplied = !Arg.IsMissing(interestRate);

            string time = Arg.Str(expiryTime).Trim();
            if (time.Length > 0 && !ClockTimePattern.IsMatch(time))
                return ExcelTable.Error("ExpiryTime must be in HH:MM format, for example 15:30");

            // Columns three and four of the range, when present, override the underlying
            // resolution for that contract.
            var overrides = UnderlyingOverrides(symbols, requested.Count);

            string identity = string.Join(";", requested.Select((pair, index) =>
                pair.Symbol + ":" + pair.Exchange + ":" +
                (index < overrides.Count ? overrides[index].Symbol + ":" + overrides[index].Exchange : "")));

            return AsyncTaskUtil.RunTask(
                nameof(oa_multioptiongreeks),
                new object[] { identity, rateSupplied ? rate : double.NaN, time },
                async () =>
                {
                    var list = new JArray();
                    for (int i = 0; i < requested.Count; i++)
                    {
                        var item = new JObject
                        {
                            ["symbol"] = requested[i].Symbol,
                            ["exchange"] = requested[i].Exchange
                        };
                        if (i < overrides.Count)
                        {
                            if (overrides[i].Symbol.Length > 0)
                                item["underlying_symbol"] = overrides[i].Symbol;
                            if (overrides[i].Exchange.Length > 0)
                                item["underlying_exchange"] = overrides[i].Exchange;
                        }
                        list.Add(item);
                    }

                    var payload = new JObject { ["symbols"] = list };
                    if (rateSupplied)
                        payload["interest_rate"] = rate;
                    if (time.Length > 0)
                        payload["expiry_time"] = time;

                    var response = await OpenAlgoClient.PostAsync("multioptiongreeks", payload);

                    var error = ExcelTable.ErrorOrNull(response);
                    if (error != null) return (object)error;

                    var items = (response["data"] as JArray ?? response["results"] as JArray)?.OfType<JObject>().ToList();
                    if (items == null || items.Count == 0)
                        return (object)ExcelTable.Error("No Greeks returned");

                    var itemFields = new List<string> { "symbol", "exchange", "status", "implied_volatility" };
                    itemFields.AddRange(ExtraFields(items, itemFields.Concat(new[] { "greeks", "message", "error" })));

                    var greekFields = new List<string> { "delta", "gamma", "theta", "vega", "rho" };
                    greekFields.AddRange(ExtraFields(items.Select(i => i["greeks"] as JObject).Where(g => g != null)!, greekFields));

                    var summary = response["summary"] as JObject ?? new JObject();
                    int width = itemFields.Count + greekFields.Count + 1;

                    var info = new object[width];
                    for (int i = 0; i < width; i++)
                        info[i] = "";
                    void Put(int index, object value)
                    {
                        if (index < width) info[index] = value;
                    }
                    Put(0, "Total");
                    Put(1, summary["total"] != null ? ExcelTable.Cell(summary["total"]) : (double)items.Count);
                    Put(2, "Success");
                    Put(3, ExcelTable.Cell(summary["success"]));
                    Put(4, "Failed");
                    Put(5, ExcelTable.Cell(summary["failed"]));

                    var rows = new List<object[]> { info };

                    var header = new object[width];
                    int column = 0;
                    foreach (var field in itemFields)
                        header[column++] = Caption(field);
                    foreach (var field in greekFields)
                        header[column++] = Caption(field);
                    header[column] = "Error";
                    rows.Add(header);

                    for (int i = 0; i < items.Count; i++)
                    {
                        var item = items[i];
                        var greeks = item["greeks"] as JObject ?? new JObject();

                        var row = new object[width];
                        int cell = 0;
                        foreach (var field in itemFields)
                        {
                            object value = ExcelTable.Cell(item[field]);
                            // Fall back to what was asked for when the item omits it.
                            if (value is string text && text.Length == 0 && i < requested.Count)
                            {
                                if (field == "symbol") value = requested[i].Symbol;
                                else if (field == "exchange") value = requested[i].Exchange;
                            }
                            row[cell++] = value;
                        }
                        foreach (var field in greekFields)
                            row[cell++] = ExcelTable.Cell(greeks[field]);
                        row[cell] = item["message"]?.ToString() ?? item["error"]?.ToString() ?? "";
                        rows.Add(row);
                    }

                    return (object)ExcelTable.FromRows(rows);
                })!;
        }

        /// <summary>
        /// Synthetic futures price from the ATM call and put through put-call parity.
        /// The basis over spot is appended as a convenience row.
        /// </summary>
        [ExcelFunction(
            Name = "oa_syntheticfuture",
            Description = "Synthetic futures price for an expiry from ATM options using put-call parity.",
            Category = OptionsCategory)]
        public static object oa_syntheticfuture(
            [ExcelArgument(Name = "Underlying", Description = "Underlying symbol, for example NIFTY, BANKNIFTY or SENSEX")] string underlying,
            [ExcelArgument(Name = "Exchange", Description = "Underlying exchange: NSE_INDEX or BSE_INDEX")] string exchange,
            [ExcelArgument(Name = "Expiry", Description = "Expiry date in DDMMMYY format, for example 25AUG26")] string expiry)
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error(NoApiKey);
            if (string.IsNullOrWhiteSpace(underlying) || string.IsNullOrWhiteSpace(exchange) || string.IsNullOrWhiteSpace(expiry))
                return ExcelTable.Error("Underlying, Exchange and Expiry are required");

            string expiryDate = expiry.Trim().ToUpperInvariant();

            return AsyncTaskUtil.RunTask(
                nameof(oa_syntheticfuture),
                new object[] { underlying, exchange, expiryDate },
                async () =>
                {
                    var payload = new JObject
                    {
                        ["underlying"] = underlying,
                        ["exchange"] = exchange,
                        ["expiry_date"] = expiryDate
                    };

                    var response = await OpenAlgoClient.PostAsync("syntheticfuture", payload);

                    var error = ExcelTable.ErrorOrNull(response);
                    if (error != null) return (object)error;

                    var data = Body(response);
                    if (data["synthetic_future_price"] == null)
                        return (object)ExcelTable.Error("No synthetic future price returned");

                    var pairs = Pairs(data, new[]
                    {
                        "synthetic_future_price", "underlying", "expiry", "atm_strike", "underlying_ltp"
                    });

                    double? spot = data["underlying_ltp"]?.ToObject<double?>();
                    double? synthetic = data["synthetic_future_price"]?.ToObject<double?>();
                    if (spot.HasValue && synthetic.HasValue)
                        pairs.Add(new KeyValuePair<string, object>("Basis", synthetic.Value - spot.Value));

                    return (object)ExcelTable.KeyValue(pairs, "Field", "Value");
                })!;
        }

        // ------------------------------------------------------------------
        // Order placement
        // ------------------------------------------------------------------

        /// <summary>
        /// Places a single option order by strike offset instead of an exact strike.
        /// Refuses to fire until trading has been armed with oa_trading_enabled(TRUE),
        /// because Excel re-evaluates the formula on every recalculation.
        /// </summary>
        [ExcelFunction(
            Name = "oa_optionsorder",
            Description = "Places an option order by strike offset (ATM, ITM1-ITM50, OTM1-OTM50). Requires oa_trading_enabled(TRUE).",
            Category = OrdersCategory)]
        public static object oa_optionsorder(
            [ExcelArgument(Name = "Strategy", Description = "Strategy identifier recorded against the order")] string strategy,
            [ExcelArgument(Name = "Underlying", Description = "Underlying symbol, for example NIFTY or BANKNIFTY")] string underlying,
            [ExcelArgument(Name = "Exchange", Description = "NSE_INDEX, BSE_INDEX, NFO or BFO")] string exchange,
            [ExcelArgument(Name = "Expiry", Description = "Expiry date in DDMMMYY format. Leave blank when the underlying already carries the expiry")] object expiry,
            [ExcelArgument(Name = "StrikeOffset", Description = "Strike offset: ATM, ITM1 to ITM50, or OTM1 to OTM50")] string strikeOffset,
            [ExcelArgument(Name = "OptionType", Description = "CE for a call or PE for a put")] string optionType,
            [ExcelArgument(Name = "Action", Description = "BUY or SELL")] string action,
            [ExcelArgument(Name = "Quantity", Description = "Order quantity in units, not lots")] object quantity,
            [ExcelArgument(Name = "PriceType", Description = "MARKET, LIMIT, SL or SL-M (optional, default MARKET)")] object? priceType = null,
            [ExcelArgument(Name = "Product", Description = "MIS or NRML (optional, default MIS)")] object? product = null,
            [ExcelArgument(Name = "Price", Description = "Limit price for LIMIT and SL orders (optional, default 0)")] object? price = null,
            [ExcelArgument(Name = "TriggerPrice", Description = "Trigger price for SL and SL-M orders (optional, default 0)")] object? triggerPrice = null,
            [ExcelArgument(Name = "SplitSize", Description = "Split the order into chunks of this size, 0 for no split (optional)")] object? splitSize = null)
        {
            if (!OpenAlgoConfig.TradingEnabled)
                return ExcelTable.Error(TradingDisabledMessage);
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error(NoApiKey);
            if (string.IsNullOrWhiteSpace(strategy) || string.IsNullOrWhiteSpace(underlying) || string.IsNullOrWhiteSpace(exchange))
                return ExcelTable.Error("Strategy, Underlying and Exchange are required");

            string offset = NormaliseOffset(strikeOffset, out string? offsetError);
            if (offsetError != null)
                return ExcelTable.Error(offsetError);
            string type = NormaliseOptionType(optionType, out string? typeError);
            if (typeError != null)
                return ExcelTable.Error(typeError);
            string side = NormaliseAction(action, out string? actionError);
            if (actionError != null)
                return ExcelTable.Error(actionError);

            int qty = Arg.Int(quantity);
            if (qty <= 0)
                return ExcelTable.Error("Quantity must be a positive whole number");

            string expiryDate = Arg.Str(expiry).Trim().ToUpperInvariant();
            string priceTypeText = Arg.Str(priceType, "MARKET").Trim().ToUpperInvariant();
            string productText = Arg.Str(product, "MIS").Trim().ToUpperInvariant();
            double limitPrice = Arg.Num(price);
            double trigger = Arg.Num(triggerPrice);
            int split = Arg.Int(splitSize);

            return AsyncTaskUtil.RunTask(
                nameof(oa_optionsorder),
                new object[]
                {
                    strategy, underlying, exchange, expiryDate, offset, type, side,
                    qty, priceTypeText, productText, limitPrice, trigger, split
                },
                async () =>
                {
                    var payload = new JObject
                    {
                        ["strategy"] = strategy,
                        ["underlying"] = underlying,
                        ["exchange"] = exchange,
                        ["offset"] = offset,
                        ["option_type"] = type,
                        ["action"] = side,
                        ["quantity"] = qty.ToString(CultureInfo.InvariantCulture),
                        ["pricetype"] = priceTypeText,
                        ["product"] = productText,
                        ["splitsize"] = split.ToString(CultureInfo.InvariantCulture),
                        ["price"] = limitPrice.ToString(CultureInfo.InvariantCulture),
                        ["trigger_price"] = trigger.ToString(CultureInfo.InvariantCulture)
                    };
                    if (expiryDate.Length > 0)
                        payload["expiry_date"] = expiryDate;

                    var response = await OpenAlgoClient.PostAsync("optionsorder", payload);

                    var error = ExcelTable.ErrorOrNull(response);
                    if (error != null) return (object)error;

                    return (object)Receipt(response, new[]
                    {
                        "orderid", "symbol", "exchange", "offset", "option_type",
                        "underlying", "underlying_ltp", "mode"
                    });
                })!;
        }

        /// <summary>
        /// Places a multi leg option strategy from a table of legs laid out on the sheet.
        ///
        /// The first row of the range names the columns; recognised names are Offset,
        /// Option Type, Action, Quantity, Expiry, PriceType, Product, SplitSize, Price and
        /// TriggerPrice. The result is one row per leg carrying that leg's order id or its
        /// error. Refuses to fire until trading has been armed with oa_trading_enabled(TRUE).
        /// </summary>
        [ExcelFunction(
            Name = "oa_optionsmultiorder",
            Description = "Places a multi leg option strategy from a table of legs. Requires oa_trading_enabled(TRUE).",
            Category = OrdersCategory)]
        public static object oa_optionsmultiorder(
            [ExcelArgument(Name = "Strategy", Description = "Strategy identifier recorded against every leg")] string strategy,
            [ExcelArgument(Name = "Underlying", Description = "Underlying symbol, for example NIFTY or BANKNIFTY")] string underlying,
            [ExcelArgument(Name = "Exchange", Description = "NSE_INDEX or BSE_INDEX")] string exchange,
            [ExcelArgument(Name = "Expiry", Description = "Common expiry in DDMMMYY format. Leave blank when every leg carries its own expiry")] object expiry,
            [ExcelArgument(Name = "Legs", Description = "Range of legs with a header row: Offset, Option Type, Action, Quantity, and optionally Expiry, PriceType, Product, SplitSize, Price, TriggerPrice")] object legs)
        {
            if (!OpenAlgoConfig.TradingEnabled)
                return ExcelTable.Error(TradingDisabledMessage);
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error(NoApiKey);
            if (string.IsNullOrWhiteSpace(strategy) || string.IsNullOrWhiteSpace(underlying) || string.IsNullOrWhiteSpace(exchange))
                return ExcelTable.Error("Strategy, Underlying and Exchange are required");

            var parsed = ParseLegs(legs, out string? legError);
            if (legError != null)
                return ExcelTable.Error(legError);

            string expiryDate = Arg.Str(expiry).Trim().ToUpperInvariant();
            string legsJson = parsed.ToString(Newtonsoft.Json.Formatting.None);

            return AsyncTaskUtil.RunTask(
                nameof(oa_optionsmultiorder),
                new object[] { strategy, underlying, exchange, expiryDate, legsJson },
                async () =>
                {
                    var payload = new JObject
                    {
                        ["strategy"] = strategy,
                        ["underlying"] = underlying,
                        ["exchange"] = exchange,
                        ["legs"] = parsed
                    };
                    if (expiryDate.Length > 0)
                        payload["expiry_date"] = expiryDate;

                    var response = await OpenAlgoClient.PostAsync("optionsmultiorder", payload);

                    var error = ExcelTable.ErrorOrNull(response);
                    if (error != null) return (object)error;

                    var results = (response["results"] as JArray ?? response["data"] as JArray)?.OfType<JObject>().ToList();
                    if (results == null || results.Count == 0)
                        return (object)ExcelTable.Error("No leg results returned");

                    var fields = new List<string> { "leg", "action", "offset", "option_type", "symbol", "orderid", "status", "mode" };
                    fields.AddRange(ExtraFields(results, fields.Concat(new[] { "message", "error" })));

                    int width = Math.Max(fields.Count + 1, 6);

                    var info = new object[width];
                    for (int i = 0; i < width; i++)
                        info[i] = "";
                    info[0] = "Underlying";
                    info[1] = ExcelTable.Cell(response["underlying"]);
                    info[2] = "Underlying LTP";
                    info[3] = ExcelTable.Cell(response["underlying_ltp"]);
                    info[4] = "Status";
                    info[5] = ExcelTable.Cell(response["status"]);

                    var rows = new List<object[]> { info };

                    var header = new object[width];
                    for (int i = 0; i < width; i++)
                        header[i] = "";
                    int column = 0;
                    foreach (var field in fields)
                        header[column++] = field == "orderid" ? "Order ID" : Caption(field);
                    header[column] = "Message";
                    rows.Add(header);

                    for (int i = 0; i < results.Count; i++)
                    {
                        var leg = results[i];
                        var row = new object[width];
                        for (int c = 0; c < width; c++)
                            row[c] = "";

                        int cell = 0;
                        foreach (var field in fields)
                        {
                            object value = ExcelTable.Cell(leg[field]);
                            if (field == "leg" && value is string text && text.Length == 0)
                                value = (double)(i + 1);
                            row[cell++] = value;
                        }
                        row[cell] = leg["message"]?.ToString() ?? leg["error"]?.ToString() ?? "";
                        rows.Add(row);
                    }

                    return (object)ExcelTable.FromRows(rows);
                })!;
        }

        // ------------------------------------------------------------------
        // Helpers
        // ------------------------------------------------------------------

        /// <summary>
        /// The options endpoints answer with a flat document. A few deployments wrap the
        /// payload in "data", so unwrap that when it is present.
        /// </summary>
        private static JObject Body(JObject response) => response["data"] as JObject ?? response;

        /// <summary>
        /// Column caption for a JSON field name.
        /// </summary>
        private static string Caption(string field) =>
            Captions.TryGetValue(field, out string? caption) ? caption : ExcelTable.Humanise(field);

        /// <summary>
        /// Field names present in the payload but not in the known list, in the order
        /// they were first seen. Used to keep columns the documentation does not list.
        /// </summary>
        private static List<string> ExtraFields(IEnumerable<JObject> items, IEnumerable<string> known)
        {
            var seen = new HashSet<string>(known, StringComparer.OrdinalIgnoreCase);
            var extras = new List<string>();
            foreach (var item in items)
                foreach (var property in item.Properties())
                    if (seen.Add(property.Name))
                        extras.Add(property.Name);
            return extras;
        }

        /// <summary>
        /// Builds key/value pairs from a response, listing the preferred fields first and
        /// appending anything else the server returned.
        /// </summary>
        private static List<KeyValuePair<string, object>> Pairs(JObject data, string[] preferred, bool flattenGreeks = false)
        {
            var pairs = new List<KeyValuePair<string, object>>();
            var used = new HashSet<string>(preferred, StringComparer.OrdinalIgnoreCase);

            // A field the server returned as null still gets a row, blank, so a
            // documented field never disappears from the table.
            foreach (var name in preferred)
            {
                var token = data[name];
                if (token != null)
                    pairs.Add(new KeyValuePair<string, object>(Caption(name), ExcelTable.Cell(token)));
            }

            if (flattenGreeks && data["greeks"] is JObject greeks)
                foreach (var property in greeks.Properties())
                    pairs.Add(new KeyValuePair<string, object>(Caption(property.Name), ExcelTable.Cell(property.Value)));

            foreach (var property in data.Properties())
            {
                if (used.Contains(property.Name))
                    continue;
                if (Hidden.Contains(property.Name, StringComparer.OrdinalIgnoreCase))
                    continue;
                pairs.Add(new KeyValuePair<string, object>(Caption(property.Name), ExcelTable.Cell(property.Value)));
            }

            return pairs;
        }

        /// <summary>
        /// Key/value receipt for an order response: status first, then the documented
        /// fields, then anything else the server returned, then the message.
        /// </summary>
        private static object[,] Receipt(JObject response, string[] preferred)
        {
            var pairs = new List<KeyValuePair<string, object>>
            {
                new("Status", response["status"]?.ToString() ?? "unknown")
            };
            pairs.AddRange(Pairs(response, preferred));

            string? message = response["message"]?.ToString() ?? response["error"]?.ToString();
            if (!string.IsNullOrWhiteSpace(message))
                pairs.Add(new KeyValuePair<string, object>("Message", message));

            return ExcelTable.KeyValue(pairs, "Field", "Value");
        }

        /// <summary>
        /// Reads the optional underlying override columns of a symbol range. Column three
        /// is the underlying symbol and column four the underlying exchange. The rows are
        /// scanned exactly as Arg.SymbolPairs scans them so the two lists stay aligned.
        /// </summary>
        private static List<(string Symbol, string Exchange)> UnderlyingOverrides(object? value, int expected)
        {
            var list = new List<(string, string)>();
            if (value is not object[,] grid)
                return list;

            int rows = grid.GetLength(0);
            int cols = grid.GetLength(1);
            if (cols < 3)
                return list;
            // A single row of more than two cells is a horizontal symbol list.
            if (rows == 1)
                return list;

            int startRow = Arg.Str(grid[0, 0]).Trim().Equals("symbol", StringComparison.OrdinalIgnoreCase) ? 1 : 0;
            for (int r = startRow; r < rows; r++)
            {
                if (Arg.Str(grid[r, 0]).Trim().Length == 0)
                    continue;
                string spotSymbol = Arg.Str(grid[r, 2]).Trim();
                string spotExchange = cols > 3 ? Arg.Str(grid[r, 3]).Trim() : "";
                list.Add((spotSymbol, spotExchange));
            }

            // Never risk pairing an override with the wrong contract.
            if (list.Count != expected)
                list.Clear();
            return list;
        }

        /// <summary>
        /// Validates and upper cases a strike offset such as ATM, ITM3 or OTM12.
        /// </summary>
        private static string NormaliseOffset(string value, out string? error)
        {
            string text = (value ?? "").Trim().ToUpperInvariant();
            if (text.Length == 0)
            {
                error = "StrikeOffset is required. Use ATM, ITM1 to ITM50, or OTM1 to OTM50";
                return "";
            }
            if (!OffsetPattern.IsMatch(text))
            {
                error = $"StrikeOffset '{text}' is not valid. Use ATM, ITM1 to ITM50, or OTM1 to OTM50";
                return "";
            }
            error = null;
            return text;
        }

        /// <summary>
        /// Validates and upper cases an option type.
        /// </summary>
        private static string NormaliseOptionType(string value, out string? error)
        {
            string text = (value ?? "").Trim().ToUpperInvariant();
            if (text != "CE" && text != "PE")
            {
                error = $"OptionType '{text}' is not valid. Use CE or PE";
                return "";
            }
            error = null;
            return text;
        }

        /// <summary>
        /// Validates and upper cases an order action.
        /// </summary>
        private static string NormaliseAction(string value, out string? error)
        {
            string text = (value ?? "").Trim().ToUpperInvariant();
            if (text != "BUY" && text != "SELL")
            {
                error = $"Action '{text}' is not valid. Use BUY or SELL";
                return "";
            }
            error = null;
            return text;
        }

        /// <summary>
        /// Column captions the leg table accepts, mapped to the JSON field the
        /// optionsmultiorder endpoint expects.
        /// </summary>
        private static readonly Dictionary<string, string> LegColumns = new(StringComparer.OrdinalIgnoreCase)
        {
            ["offset"] = "offset",
            ["strikeoffset"] = "offset",
            ["strike offset"] = "offset",
            ["strike_offset"] = "offset",
            ["optiontype"] = "option_type",
            ["option type"] = "option_type",
            ["option_type"] = "option_type",
            ["type"] = "option_type",
            ["ce/pe"] = "option_type",
            ["action"] = "action",
            ["side"] = "action",
            ["buy/sell"] = "action",
            ["quantity"] = "quantity",
            ["qty"] = "quantity",
            ["expiry"] = "expiry_date",
            ["expirydate"] = "expiry_date",
            ["expiry date"] = "expiry_date",
            ["expiry_date"] = "expiry_date",
            ["pricetype"] = "pricetype",
            ["price type"] = "pricetype",
            ["price_type"] = "pricetype",
            ["ordertype"] = "pricetype",
            ["order type"] = "pricetype",
            ["product"] = "product",
            ["producttype"] = "product",
            ["product type"] = "product",
            ["product_type"] = "product",
            ["splitsize"] = "splitsize",
            ["split size"] = "splitsize",
            ["split_size"] = "splitsize",
            ["split"] = "splitsize",
            ["price"] = "price",
            ["limit price"] = "price",
            ["limit_price"] = "price",
            ["triggerprice"] = "trigger_price",
            ["trigger price"] = "trigger_price",
            ["trigger_price"] = "trigger_price"
        };

        /// <summary>
        /// Column captions that are for the reader rather than the endpoint, and are
        /// skipped instead of rejected.
        /// </summary>
        private static readonly HashSet<string> IgnoredLegColumns = new(StringComparer.OrdinalIgnoreCase)
        {
            "leg", "leg no", "legno", "no", "#", "note", "notes", "comment", "comments",
            "remark", "remarks", "label", "description"
        };

        /// <summary>
        /// Turns a worksheet range of legs into the JSON array the endpoint expects.
        /// The first row of the range must name the columns.
        /// </summary>
        private static JArray ParseLegs(object? value, out string? error)
        {
            var result = new JArray();

            if (Arg.IsMissing(value) || value is not object[,] grid)
            {
                error = "Legs must be a range with a header row naming the columns, for example Offset, Option Type, Action, Quantity";
                return result;
            }

            int rows = grid.GetLength(0);
            int cols = grid.GetLength(1);
            if (rows < 2)
            {
                error = "Legs needs a header row and at least one leg row";
                return result;
            }

            var fields = new string?[cols];
            for (int c = 0; c < cols; c++)
            {
                string caption = Arg.Str(grid[0, c]).Trim();
                if (caption.Length == 0 || IgnoredLegColumns.Contains(caption))
                {
                    fields[c] = null;
                    continue;
                }
                if (!LegColumns.TryGetValue(caption, out string? field))
                {
                    error = $"Leg column '{caption}' is not recognised. Use Offset, Option Type, Action, Quantity, Expiry, PriceType, Product, SplitSize, Price or TriggerPrice";
                    return result;
                }
                fields[c] = field;
            }

            for (int r = 1; r < rows; r++)
            {
                // Skip blank spacer rows so a user can leave gaps in the layout.
                bool empty = true;
                for (int c = 0; c < cols; c++)
                    if (!Arg.IsMissing(grid[r, c]))
                    {
                        empty = false;
                        break;
                    }
                if (empty)
                    continue;

                var leg = new JObject();
                int legNumber = result.Count + 1;

                for (int c = 0; c < cols; c++)
                {
                    string? field = fields[c];
                    if (field == null || Arg.IsMissing(grid[r, c]))
                        continue;

                    object cell = grid[r, c];
                    switch (field)
                    {
                        case "quantity":
                            int quantity = Arg.Int(cell);
                            if (quantity <= 0)
                            {
                                error = $"Leg {legNumber}: quantity must be a positive whole number";
                                return result;
                            }
                            leg["quantity"] = quantity;
                            break;
                        case "splitsize":
                            leg["splitsize"] = Arg.Int(cell);
                            break;
                        case "price":
                        case "trigger_price":
                            leg[field] = Arg.Num(cell);
                            break;
                        case "offset":
                            string offset = NormaliseOffset(Arg.Str(cell), out string? offsetError);
                            if (offsetError != null)
                            {
                                error = $"Leg {legNumber}: {offsetError}";
                                return result;
                            }
                            leg["offset"] = offset;
                            break;
                        case "option_type":
                            string type = NormaliseOptionType(Arg.Str(cell), out string? typeError);
                            if (typeError != null)
                            {
                                error = $"Leg {legNumber}: {typeError}";
                                return result;
                            }
                            leg["option_type"] = type;
                            break;
                        case "action":
                            string side = NormaliseAction(Arg.Str(cell), out string? actionError);
                            if (actionError != null)
                            {
                                error = $"Leg {legNumber}: {actionError}";
                                return result;
                            }
                            leg["action"] = side;
                            break;
                        default:
                            leg[field] = Arg.Str(cell).Trim().ToUpperInvariant();
                            break;
                    }
                }

                foreach (var required in new[] { "offset", "option_type", "action", "quantity" })
                {
                    if (leg[required] == null)
                    {
                        error = $"Leg {legNumber}: {Caption(required)} is missing";
                        return result;
                    }
                }

                result.Add(leg);
            }

            if (result.Count == 0)
            {
                error = "Legs contains no leg rows";
                return result;
            }
            if (result.Count > 20)
            {
                error = $"OptionsMultiOrder accepts at most 20 legs, {result.Count} were supplied";
                return result;
            }

            error = null;
            return result;
        }
    }
}
