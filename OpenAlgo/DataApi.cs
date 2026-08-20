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
    /// Market data worksheet functions: quotes, multi quotes, depth, history
    /// and the broker interval list.
    /// </summary>
    public static class DataApi
    {
        /// <summary>
        /// Trims and upper-cases an exchange code. Every request schema validates
        /// exchange with a case sensitive OneOf against VALID_EXCHANGES, so a cell
        /// reading "nse" or " NSE " is rejected with HTTP 400 rather than understood.
        /// </summary>
        private static string NormExchange(object? value) =>
            Arg.Str(value).Trim().ToUpperInvariant();

        /// <summary>
        /// Indian Standard Time. OpenAlgo reports market data against this offset and
        /// India does not observe daylight saving, so a fixed offset is exact.
        /// </summary>
        private static readonly TimeSpan IstOffset = new TimeSpan(5, 30, 0);

        // ---------------------------------------------------------------------
        // Quotes
        // ---------------------------------------------------------------------

        /// <summary>
        /// Retrieves the full market quote for one symbol as a key/value table,
        /// with Change and Change % computed from ltp and prev_close.
        /// </summary>
        [ExcelFunction(Name = "oa_quotes", Description = "Full market quote for a symbol as a key/value table.", Category = "OpenAlgo Data")]
        public static object oa_quotes(
            [ExcelArgument(Name = "Symbol", Description = "Trading symbol, for example RELIANCE")] string symbol,
            [ExcelArgument(Name = "Exchange", Description = "Exchange code: NSE, BSE, NFO, BFO, CDS, BCD, MCX")] string exchange)
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");
            if (string.IsNullOrWhiteSpace(symbol) || string.IsNullOrWhiteSpace(exchange))
                return ExcelTable.Error("Symbol and Exchange are required");

            return AsyncTaskUtil.RunTask(nameof(oa_quotes), new object[] { symbol, exchange }, async () =>
            {
                var payload = new JObject { ["symbol"] = Arg.Str(symbol).Trim(), ["exchange"] = NormExchange(exchange) };
                var response = await OpenAlgoClient.PostAsync("quotes", payload);

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                if (response["data"] is not JObject data || data.Count == 0)
                    return (object)ExcelTable.Error("No quote data found");

                var pairs = new List<KeyValuePair<string, object>>();
                foreach (var property in data.Properties())
                    pairs.Add(new KeyValuePair<string, object>(ExcelTable.Humanise(property.Name), ExcelTable.Cell(property.Value)));

                double? ltp = Number(data, "ltp");
                double? previous = Number(data, "prev_close");
                if (ltp.HasValue && previous.HasValue && Math.Abs(previous.Value) > double.Epsilon)
                {
                    double change = ltp.Value - previous.Value;
                    pairs.Add(new KeyValuePair<string, object>("Change", change));
                    pairs.Add(new KeyValuePair<string, object>("Change %", change / previous.Value * 100d));
                }

                return (object)ExcelTable.KeyValue(pairs, $"{symbol} ({exchange})", "Value");
            })!;
        }

        /// <summary>
        /// Retrieves the last traded price for one symbol as a single number.
        /// The compact form to fill a watchlist column with.
        /// </summary>
        [ExcelFunction(Name = "oa_ltp", Description = "Last traded price for a symbol as a single number.", Category = "OpenAlgo Data")]
        public static object oa_ltp(
            [ExcelArgument(Name = "Symbol", Description = "Trading symbol, for example RELIANCE")] string symbol,
            [ExcelArgument(Name = "Exchange", Description = "Exchange code: NSE, BSE, NFO, BFO, CDS, BCD, MCX")] string exchange)
        {
            if (!OpenAlgoClient.HasApiKey)
                return "Error: OpenAlgo API Key is not set. Use oa_api()";
            if (string.IsNullOrWhiteSpace(symbol) || string.IsNullOrWhiteSpace(exchange))
                return "Error: Symbol and Exchange are required";

            return AsyncTaskUtil.RunTask(nameof(oa_ltp), new object[] { symbol, exchange }, async () =>
            {
                var payload = new JObject { ["symbol"] = Arg.Str(symbol).Trim(), ["exchange"] = NormExchange(exchange) };
                var response = await OpenAlgoClient.PostAsync("quotes", payload);

                if (OpenAlgoClient.IsError(response))
                    return (object)("Error: " + OpenAlgoClient.MessageOf(response));

                if (response["data"] is not JObject data)
                    return (object)"Error: No quote data found";

                double? ltp = Number(data, "ltp");
                if (!ltp.HasValue)
                    return (object)$"Error: No ltp for {symbol} on {exchange}";

                return (object)ltp.Value;
            })!;
        }

        /// <summary>
        /// Retrieves one named field of a quote as a single value.
        /// Accepts the fields the API returns plus the derived change and changepct.
        /// </summary>
        [ExcelFunction(Name = "oa_field", Description = "One named field of a market quote as a single value.", Category = "OpenAlgo Data")]
        public static object oa_field(
            [ExcelArgument(Name = "Symbol", Description = "Trading symbol, for example RELIANCE")] string symbol,
            [ExcelArgument(Name = "Exchange", Description = "Exchange code: NSE, BSE, NFO, BFO, CDS, BCD, MCX")] string exchange,
            [ExcelArgument(Name = "Field", Description = "ltp, open, high, low, prev_close, bid, ask, volume, oi, change or changepct")] string field)
        {
            if (!OpenAlgoClient.HasApiKey)
                return "Error: OpenAlgo API Key is not set. Use oa_api()";
            if (string.IsNullOrWhiteSpace(symbol) || string.IsNullOrWhiteSpace(exchange))
                return "Error: Symbol and Exchange are required";
            if (string.IsNullOrWhiteSpace(field))
                return "Error: Field is required";

            string wanted = NormaliseField(field);

            return AsyncTaskUtil.RunTask(nameof(oa_field), new object[] { symbol, exchange, wanted }, async () =>
            {
                var payload = new JObject { ["symbol"] = Arg.Str(symbol).Trim(), ["exchange"] = NormExchange(exchange) };
                var response = await OpenAlgoClient.PostAsync("quotes", payload);

                if (OpenAlgoClient.IsError(response))
                    return (object)("Error: " + OpenAlgoClient.MessageOf(response));

                if (response["data"] is not JObject data)
                    return (object)"Error: No quote data found";

                if (wanted == "change" || wanted == "changepct")
                {
                    double? ltp = Number(data, "ltp");
                    double? previous = Number(data, "prev_close");
                    if (!ltp.HasValue || !previous.HasValue)
                        return (object)"Error: ltp and prev_close are needed to compute change";
                    if (wanted == "change")
                        return (object)(ltp.Value - previous.Value);
                    if (Math.Abs(previous.Value) <= double.Epsilon)
                        return (object)"Error: prev_close is zero";
                    return (object)((ltp.Value - previous.Value) / previous.Value * 100d);
                }

                var token = data[wanted];
                if (token == null || token.Type == JTokenType.Null)
                    return (object)$"Error: Unknown field '{field}'. Use ltp, open, high, low, prev_close, bid, ask, volume, oi, change or changepct";

                return ExcelTable.Cell(token);
            })!;
        }

        /// <summary>
        /// Retrieves quotes for a range of symbols in one call, one row per symbol.
        /// The range holds one column of symbols, or two columns of symbol and exchange.
        /// </summary>
        [ExcelFunction(Name = "oa_multiquotes", Description = "Quotes for a range of symbols, one row per symbol.", Category = "OpenAlgo Data")]
        public static object oa_multiquotes(
            [ExcelArgument(Name = "Symbols", Description = "Range of symbols, or two columns of symbol and exchange", AllowReference = false)] object symbolRange,
            [ExcelArgument(Name = "DefaultExchange", Description = "Exchange applied to symbols with no exchange column. Defaults to NSE")] object defaultExchange)
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");

            string fallbackExchange = Arg.Str(defaultExchange, "NSE").Trim();
            if (fallbackExchange.Length == 0)
                fallbackExchange = "NSE";

            var pairs = Arg.SymbolPairs(symbolRange, fallbackExchange);
            if (pairs.Count == 0)
                return ExcelTable.Error("No symbols supplied");

            // The identity of the async call has to cover every resolved symbol,
            // otherwise two ranges would share one cached result.
            string identity = string.Join(",", pairs.Select(p => p.Exchange + ":" + p.Symbol));

            return AsyncTaskUtil.RunTask(nameof(oa_multiquotes), new object[] { identity }, async () =>
            {
                var symbols = new JArray();
                foreach (var (symbol, exchange) in pairs)
                    symbols.Add(new JObject { ["symbol"] = Arg.Str(symbol).Trim(), ["exchange"] = NormExchange(exchange) });

                var response = await OpenAlgoClient.PostAsync("multiquotes", new JObject { ["symbols"] = symbols });

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                // The multiquotes wrapper key is "results", not "data".
                if (response["results"] is not JArray results)
                    return (object)ExcelTable.Error("No results returned");

                string[] headers =
                {
                    "Symbol", "Exchange", "LTP", "Open", "High", "Low", "Prev Close",
                    "Change", "Change %", "Bid", "Ask", "OI", "Volume", "Error"
                };

                var table = new object[results.Count + 1, headers.Length];
                for (int col = 0; col < headers.Length; col++)
                    table[0, col] = headers[col];

                for (int i = 0; i < results.Count; i++)
                {
                    var result = results[i] as JObject;
                    var data = result?["data"] as JObject;
                    int row = i + 1;

                    // Fall back to the requested pair when the server echoes neither.
                    string symbol = result?["symbol"]?.ToString() ?? (i < pairs.Count ? pairs[i].Symbol : "");
                    string exchange = result?["exchange"]?.ToString() ?? (i < pairs.Count ? pairs[i].Exchange : "");

                    double? ltp = Number(data, "ltp");
                    double? previous = Number(data, "prev_close");

                    table[row, 0] = symbol;
                    table[row, 1] = exchange;
                    table[row, 2] = Value(ltp);
                    table[row, 3] = Value(Number(data, "open"));
                    table[row, 4] = Value(Number(data, "high"));
                    table[row, 5] = Value(Number(data, "low"));
                    table[row, 6] = Value(previous);

                    if (ltp.HasValue && previous.HasValue)
                    {
                        double change = ltp.Value - previous.Value;
                        table[row, 7] = change;
                        table[row, 8] = Math.Abs(previous.Value) > double.Epsilon
                            ? (object)(change / previous.Value * 100d)
                            : "";
                    }
                    else
                    {
                        table[row, 7] = "";
                        table[row, 8] = "";
                    }

                    table[row, 9] = Value(Number(data, "bid"));
                    table[row, 10] = Value(Number(data, "ask"));
                    table[row, 11] = Value(Number(data, "oi"));
                    table[row, 12] = Value(Number(data, "volume"));

                    string? rowError = result?["error"]?.ToString();
                    if (string.IsNullOrWhiteSpace(rowError) && data == null)
                        rowError = "No quote returned";
                    table[row, 13] = rowError ?? "";
                }

                return (object)table;
            })!;
        }

        // ---------------------------------------------------------------------
        // Depth
        // ---------------------------------------------------------------------

        /// <summary>
        /// Retrieves the order book depth for a symbol: a summary block followed by
        /// the top of book ladder of asks and bids.
        /// </summary>
        [ExcelFunction(Name = "oa_depth", Description = "Order book depth for a symbol with the day summary.", Category = "OpenAlgo Data")]
        public static object oa_depth(
            [ExcelArgument(Name = "Symbol", Description = "Trading symbol, for example SBIN")] string symbol,
            [ExcelArgument(Name = "Exchange", Description = "Exchange code: NSE, BSE, NFO, BFO, CDS, BCD, MCX")] string exchange)
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");
            if (string.IsNullOrWhiteSpace(symbol) || string.IsNullOrWhiteSpace(exchange))
                return ExcelTable.Error("Symbol and Exchange are required");

            return AsyncTaskUtil.RunTask(nameof(oa_depth), new object[] { symbol, exchange }, async () =>
            {
                var payload = new JObject { ["symbol"] = Arg.Str(symbol).Trim(), ["exchange"] = NormExchange(exchange) };
                var response = await OpenAlgoClient.PostAsync("depth", payload);

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                if (response["data"] is not JObject data)
                    return (object)ExcelTable.Error("No depth data found");

                var asks = data["asks"] as JArray ?? new JArray();
                var bids = data["bids"] as JArray ?? new JArray();
                int ladderRows = Math.Max(asks.Count, bids.Count);

                // Five summary rows, one ladder header row, then the ladder itself.
                const int summaryRows = 5;
                var table = new object[summaryRows + 1 + ladderRows, 4];

                void Summary(int row, string leftLabel, double? leftValue, string rightLabel, double? rightValue)
                {
                    table[row, 0] = leftLabel;
                    table[row, 1] = Value(leftValue);
                    table[row, 2] = rightLabel;
                    table[row, 3] = Value(rightValue);
                }

                Summary(0, "LTP", Number(data, "ltp"), "Volume", Number(data, "volume"));
                Summary(1, "Open", Number(data, "open"), "High", Number(data, "high"));
                Summary(2, "Low", Number(data, "low"), "Prev Close", Number(data, "prev_close"));
                Summary(3, "LTQ", Number(data, "ltq"), "OI", Number(data, "oi"));
                Summary(4, "Total Buy Qty", Number(data, "totalbuyqty"), "Total Sell Qty", Number(data, "totalsellqty"));

                table[summaryRows, 0] = "Ask Price";
                table[summaryRows, 1] = "Ask Quantity";
                table[summaryRows, 2] = "Bid Price";
                table[summaryRows, 3] = "Bid Quantity";

                for (int i = 0; i < ladderRows; i++)
                {
                    var ask = asks.ElementAtOrDefault(i) as JObject;
                    var bid = bids.ElementAtOrDefault(i) as JObject;
                    int row = summaryRows + 1 + i;

                    table[row, 0] = Value(Number(ask, "price"));
                    table[row, 1] = Value(Number(ask, "quantity"));
                    table[row, 2] = Value(Number(bid, "price"));
                    table[row, 3] = Value(Number(bid, "quantity"));
                }

                return (object)table;
            })!;
        }

        // ---------------------------------------------------------------------
        // History
        // ---------------------------------------------------------------------

        /// <summary>
        /// Retrieves historical OHLCV candles. The Date column is a real Excel date
        /// serial including the time of day, so the result can be charted directly.
        /// </summary>
        [ExcelFunction(Name = "oa_history", Description = "Historical OHLCV candles for a symbol.", Category = "OpenAlgo Data")]
        public static object oa_history(
            [ExcelArgument(Name = "Symbol", Description = "Trading symbol, for example SBIN")] string symbol,
            [ExcelArgument(Name = "Exchange", Description = "Exchange code: NSE, BSE, NFO, BFO, CDS, BCD, MCX")] string exchange,
            [ExcelArgument(Name = "Interval", Description = "Candle interval: 1m, 3m, 5m, 10m, 15m, 30m, 1h, D, W or M. See oa_intervals()")] string interval,
            [ExcelArgument(Name = "StartDate", Description = "Start date as a date cell or YYYY-MM-DD text")] object startDate,
            [ExcelArgument(Name = "EndDate", Description = "End date as a date cell or YYYY-MM-DD text")] object endDate,
            [ExcelArgument(Name = "Source", Description = "Optional data source: broker (default) or db for locally stored Historify data")] object source)
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");
            if (string.IsNullOrWhiteSpace(symbol) || string.IsNullOrWhiteSpace(exchange))
                return ExcelTable.Error("Symbol and Exchange are required");
            if (string.IsNullOrWhiteSpace(interval))
                return ExcelTable.Error("Interval is required. See oa_intervals()");

            string start = Arg.Date(startDate);
            string end = Arg.Date(endDate);
            if (start.Length == 0 || end.Length == 0)
                return ExcelTable.Error("StartDate and EndDate are required");

            string dataSource = NormaliseSource(Arg.Str(source));
            if (dataSource.Length == 0)
                return ExcelTable.Error("Source must be broker or db");

            return AsyncTaskUtil.RunTask(nameof(oa_history), new object[] { symbol, exchange, interval, start, end, dataSource }, async () =>
            {
                var payload = new JObject
                {
                    ["symbol"] = Arg.Str(symbol).Trim(),
                    ["exchange"] = NormExchange(exchange),
                    ["interval"] = interval,
                    ["start_date"] = start,
                    ["end_date"] = end,
                    ["source"] = dataSource
                };

                var response = await OpenAlgoClient.PostAsync("history", payload);

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                if (response["data"] is not JArray candles)
                    return (object)ExcelTable.Error("No historical data found");

                return (object)CandleTable(candles, symbol);
            })!;
        }

        // ---------------------------------------------------------------------
        // Intervals
        // ---------------------------------------------------------------------

        /// <summary>
        /// Retrieves the candle intervals the connected broker supports, grouped by
        /// whatever categories the server reports.
        /// </summary>
        [ExcelFunction(Name = "oa_intervals", Description = "Candle intervals supported by the connected broker.", Category = "OpenAlgo Data")]
        public static object oa_intervals()
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");

            return AsyncTaskUtil.RunTask(nameof(oa_intervals), new object[] { OpenAlgoConfig.HostUrl }, async () =>
            {
                var response = await OpenAlgoClient.PostAsync("intervals", new JObject());

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                if (response["data"] is not JObject data)
                    return (object)ExcelTable.Error("No interval data found");

                var rows = new List<object[]> { new object[] { "Category", "Interval" } };

                // Categories vary by broker, so read whatever the response carries
                // rather than assuming a fixed set of six.
                foreach (var property in data.Properties())
                {
                    string category = ExcelTable.Humanise(property.Name);

                    if (property.Value is JArray values)
                    {
                        foreach (var value in values)
                            rows.Add(new object[] { category, ExcelTable.Cell(value) });
                    }
                    else if (property.Value.Type != JTokenType.Null)
                    {
                        rows.Add(new object[] { category, ExcelTable.Cell(property.Value) });
                    }
                }

                if (rows.Count == 1)
                    return (object)ExcelTable.Error("The broker reported no intervals");

                return (object)ExcelTable.FromRows(rows);
            })!;
        }

        // ---------------------------------------------------------------------
        // Helpers
        // ---------------------------------------------------------------------

        /// <summary>
        /// Builds the OHLCV candle table.
        /// The OI column appears only when the payload carries open interest.
        /// </summary>
        private static object[,] CandleTable(JArray candles, string symbol)
        {
            if (candles.Count == 0)
                return ExcelTable.Error("No candles returned for the requested range");

            bool hasOpenInterest = candles.OfType<JObject>()
                .Any(candle => candle["oi"] != null && candle["oi"]!.Type != JTokenType.Null);

            var headers = new List<string> { "Symbol", "Date", "Time (IST)", "Open", "High", "Low", "Close", "Volume" };
            if (hasOpenInterest)
                headers.Add("OI");

            var table = new object[candles.Count + 1, headers.Count];
            for (int col = 0; col < headers.Count; col++)
                table[0, col] = headers[col];

            for (int i = 0; i < candles.Count; i++)
            {
                var candle = candles[i] as JObject;
                DateTime? stamp = ParseTimestamp(candle?["timestamp"]);
                int row = i + 1;

                table[row, 0] = symbol;
                // A real Excel serial, time of day included, so the result charts directly.
                table[row, 1] = stamp.HasValue ? (object)stamp.Value.ToOADate() : "";
                table[row, 2] = stamp.HasValue ? stamp.Value.ToString("HH:mm:ss", CultureInfo.InvariantCulture) : "";
                table[row, 3] = Value(Number(candle, "open"));
                table[row, 4] = Value(Number(candle, "high"));
                table[row, 5] = Value(Number(candle, "low"));
                table[row, 6] = Value(Number(candle, "close"));
                // Volume is read as a double: an int overflows on high volume instruments.
                table[row, 7] = Value(Number(candle, "volume"));
                if (hasOpenInterest)
                    table[row, 8] = Value(Number(candle, "oi"));
            }

            return table;
        }

        /// <summary>
        /// Reads a candle timestamp in IST.
        ///
        /// The field arrives either as a Unix epoch count or as a formatted string such
        /// as "2025-04-01 09:15:00+05:30". A numeric value is epoch seconds (epoch
        /// milliseconds are detected by magnitude). A string that carries its own UTC
        /// offset is interpreted with that offset; one that does not falls back to
        /// treating the value as UTC and shifting it by +05:30, which is what the
        /// endpoint historically returned.
        /// </summary>
        private static DateTime? ParseTimestamp(JToken? token)
        {
            if (token == null || token.Type == JTokenType.Null || token.Type == JTokenType.Undefined)
                return null;

            if (token.Type == JTokenType.Integer || token.Type == JTokenType.Float)
                return FromUnix(token.ToObject<double>());

            if (token.Type == JTokenType.Date)
                return token.ToObject<DateTime>();

            string text = token.ToString().Trim();
            if (text.Length == 0)
                return null;

            // A bare number arriving as a string is still an epoch count.
            if (text.IndexOf(':') < 0 && text.IndexOf('-') < 0 && text.IndexOf('/') < 0
                && double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out double epoch))
                return FromUnix(epoch);

            var styles = HasExplicitOffset(text) ? DateTimeStyles.None : DateTimeStyles.AssumeUniversal;
            if (DateTimeOffset.TryParse(text, CultureInfo.InvariantCulture, styles, out DateTimeOffset parsed))
                return parsed.ToOffset(IstOffset).DateTime;

            return null;
        }

        /// <summary>
        /// Converts an epoch count in seconds, or in milliseconds when the magnitude
        /// says so, into IST wall clock time.
        /// </summary>
        private static DateTime? FromUnix(double epoch)
        {
            if (double.IsNaN(epoch) || double.IsInfinity(epoch) || Math.Abs(epoch) < 1d)
                return null;

            // Seconds past the year 5000 are far more likely to be milliseconds.
            if (Math.Abs(epoch) > 100000000000d)
                epoch /= 1000d;

            try
            {
                long milliseconds = (long)Math.Round(epoch * 1000d);
                return DateTimeOffset.FromUnixTimeMilliseconds(milliseconds).ToOffset(IstOffset).DateTime;
            }
            catch (ArgumentOutOfRangeException)
            {
                return null;
            }
        }

        /// <summary>
        /// True when a timestamp string ends in a UTC offset such as Z, +05:30 or -04:00.
        /// </summary>
        private static bool HasExplicitOffset(string text)
        {
            if (text.EndsWith("Z", StringComparison.OrdinalIgnoreCase))
                return true;

            // The offset always follows the time part, so only look past the first colon.
            int timePart = text.IndexOf(':');
            if (timePart < 0)
                return false;

            return text.IndexOf('+', timePart) >= 0 || text.IndexOf('-', timePart) >= 0;
        }

        /// <summary>
        /// Reads the first of the named fields that holds a number.
        /// </summary>
        private static double? Number(JObject? data, params string[] names)
        {
            if (data == null)
                return null;

            foreach (var name in names)
            {
                var token = data[name];
                if (token == null || token.Type == JTokenType.Null || token.Type == JTokenType.Undefined)
                    continue;

                if (token.Type == JTokenType.Integer || token.Type == JTokenType.Float)
                    return token.ToObject<double>();

                if (double.TryParse(token.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out double parsed))
                    return parsed;
            }
            return null;
        }

        /// <summary>
        /// Presents an optional number as a cell value, blank when it is absent.
        /// A missing price is not the same as a price of zero.
        /// </summary>
        private static object Value(double? number) => number.HasValue ? number.Value : (object)"";

        /// <summary>
        /// Reduces a user supplied quote field name to its canonical form.
        /// </summary>
        private static string NormaliseField(string field)
        {
            string text = field.Trim().ToLowerInvariant()
                .Replace(" ", "").Replace("_", "").Replace("-", "").Replace("%", "pct");

            switch (text)
            {
                case "last":
                case "lastprice":
                case "lasttradedprice":
                    return "ltp";
                case "prevclose":
                case "previousclose":
                case "close":
                    return "prev_close";
                case "openinterest":
                    return "oi";
                case "vol":
                    return "volume";
                case "chg":
                    return "change";
                case "changepct":
                case "changepercent":
                case "chgpct":
                case "pctchange":
                    return "changepct";
                default:
                    return text;
            }
        }

        /// <summary>
        /// Maps the Source argument onto the values /history accepts.
        /// Returns an empty string when the argument is not recognised.
        /// </summary>
        private static string NormaliseSource(string source)
        {
            string text = source.Trim().ToLowerInvariant();
            switch (text)
            {
                case "":
                case "api":
                case "broker":
                    return "api";
                case "db":
                case "database":
                case "historify":
                    return "db";
                default:
                    return "";
            }
        }
    }
}
