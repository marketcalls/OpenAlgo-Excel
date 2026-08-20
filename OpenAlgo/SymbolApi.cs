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
    /// Symbol service worksheet functions: instrument metadata, symbol search,
    /// expiry dates and the instrument master download.
    /// </summary>
    public static class SymbolApi
    {
        /// <summary>
        /// Trims and upper-cases an exchange code. Every request schema validates
        /// exchange with a case sensitive OneOf against VALID_EXCHANGES, so a cell
        /// reading "nse" or " NSE " is rejected with HTTP 400 rather than understood.
        /// </summary>
        private static string NormExchange(object? value) =>
            Arg.Str(value).Trim().ToUpperInvariant();

        /// <summary>
        /// Date formats used by the expiry and instrument endpoints.
        /// </summary>
        private static readonly string[] ExpiryFormats =
        {
            "dd-MMM-yy", "dd-MMM-yyyy", "d-MMM-yy", "d-MMM-yyyy",
            "ddMMMyy", "ddMMMyyyy", "yyyy-MM-dd", "dd-MM-yyyy"
        };

        /// <summary>
        /// Retrieves the full instrument metadata for a symbol, including the broker
        /// specific symbol and token needed for order placement.
        /// </summary>
        [ExcelFunction(Name = "oa_symbol", Description = "Instrument metadata for a symbol.", Category = "OpenAlgo Symbols")]
        public static object oa_symbol(
            [ExcelArgument(Name = "Symbol", Description = "Trading symbol in OpenAlgo format, for example RELIANCE")] string symbol,
            [ExcelArgument(Name = "Exchange", Description = "Exchange code: NSE, BSE, NFO, BFO, CDS, BCD, MCX")] string exchange)
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");
            if (string.IsNullOrWhiteSpace(symbol) || string.IsNullOrWhiteSpace(exchange))
                return ExcelTable.Error("Symbol and Exchange are required");

            return AsyncTaskUtil.RunTask(nameof(oa_symbol), new object[] { symbol, exchange }, async () =>
            {
                var payload = new JObject { ["symbol"] = Arg.Str(symbol).Trim(), ["exchange"] = NormExchange(exchange) };
                var response = await OpenAlgoClient.PostAsync("symbol", payload);

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                if (response["data"] is not JObject data || data.Count == 0)
                    return (object)ExcelTable.Error($"No metadata for {symbol} on {exchange}");

                return (object)ExcelTable.KeyValue(data);
            })!;
        }

        /// <summary>
        /// Retrieves the lot size of an instrument as a single number, which is what
        /// an F&amp;O order quantity has to be a multiple of.
        /// </summary>
        [ExcelFunction(Name = "oa_lotsize", Description = "Lot size of an instrument as a single number.", Category = "OpenAlgo Symbols")]
        public static object oa_lotsize(
            [ExcelArgument(Name = "Symbol", Description = "Trading symbol in OpenAlgo format, for example NIFTY25AUG26FUT")] string symbol,
            [ExcelArgument(Name = "Exchange", Description = "Exchange code: NSE, BSE, NFO, BFO, CDS, BCD, MCX")] string exchange)
        {
            return SymbolField(nameof(oa_lotsize), symbol, exchange, "lotsize");
        }

        /// <summary>
        /// Retrieves the broker instrument token of a symbol as a single value.
        /// </summary>
        [ExcelFunction(Name = "oa_token", Description = "Broker instrument token of a symbol as a single value.", Category = "OpenAlgo Symbols")]
        public static object oa_token(
            [ExcelArgument(Name = "Symbol", Description = "Trading symbol in OpenAlgo format, for example RELIANCE")] string symbol,
            [ExcelArgument(Name = "Exchange", Description = "Exchange code: NSE, BSE, NFO, BFO, CDS, BCD, MCX")] string exchange)
        {
            return SymbolField(nameof(oa_token), symbol, exchange, "token", asText: true);
        }

        /// <summary>
        /// Searches the instrument master by name, strike, expiry or any combination,
        /// for example "NIFTY 26000 DEC CE".
        /// </summary>
        [ExcelFunction(Name = "oa_search", Description = "Searches instruments by name, strike, month or option type.", Category = "OpenAlgo Symbols")]
        public static object oa_search(
            [ExcelArgument(Name = "Query", Description = "Search text, for example RELIANCE or NIFTY 26000 DEC CE")] string query,
            [ExcelArgument(Name = "Exchange", Description = "Optional exchange filter: NSE, BSE, NFO, BFO, CDS, BCD, MCX")] object exchange)
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");
            if (string.IsNullOrWhiteSpace(query))
                return ExcelTable.Error("Query is required");

            string filter = Arg.Str(exchange).Trim().ToUpperInvariant();

            return AsyncTaskUtil.RunTask(nameof(oa_search), new object[] { query, filter }, async () =>
            {
                var payload = new JObject { ["query"] = query };
                if (filter.Length > 0)
                    payload["exchange"] = filter;

                var response = await OpenAlgoClient.PostAsync("search", payload);

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                if (response["data"] is not JArray matches || matches.Count == 0)
                    return (object)ExcelTable.Error($"No instruments matched '{query}'");

                string[] fields =
                {
                    "symbol", "name", "exchange", "instrumenttype", "expiry", "strike",
                    "lotsize", "tick_size", "freeze_qty", "brsymbol", "brexchange", "token"
                };
                string[] headers =
                {
                    "Symbol", "Name", "Exchange", "Instrument Type", "Expiry", "Strike",
                    "Lot Size", "Tick Size", "Freeze Qty", "Broker Symbol", "Broker Exchange", "Token"
                };

                return (object)ExcelTable.Rows(matches, fields, headers, "No instruments matched");
            })!;
        }

        /// <summary>
        /// Retrieves the available expiry dates for a futures or options underlying.
        /// Each row carries the expiry exactly as the API states it, an Excel date
        /// serial for sorting and charting, and whether it is a monthly contract.
        /// </summary>
        [ExcelFunction(Name = "oa_expiry", Description = "Available expiry dates for a futures or options underlying.", Category = "OpenAlgo Symbols")]
        public static object oa_expiry(
            [ExcelArgument(Name = "Symbol", Description = "Underlying symbol, for example NIFTY or BANKNIFTY")] string symbol,
            [ExcelArgument(Name = "Exchange", Description = "Derivatives exchange: NFO, BFO, MCX, CDS, BCD, NCO, NCDEX or CRYPTO")] string exchange,
            [ExcelArgument(Name = "InstrumentType", Description = "options or futures")] string instrumentType,
            [ExcelArgument(Name = "ExpiryType", Description = "Optional local filter: all (default), weekly for every expiry, or monthly for the last expiry of each month")] object expiryType)
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");
            if (string.IsNullOrWhiteSpace(symbol) || string.IsNullOrWhiteSpace(exchange))
                return ExcelTable.Error("Symbol and Exchange are required");

            string kind = NormaliseInstrumentType(instrumentType);
            if (kind.Length == 0)
                return ExcelTable.Error("InstrumentType must be options or futures");

            string wanted = Arg.Str(expiryType, "all").Trim().ToLowerInvariant();
            if (wanted.Length == 0)
                wanted = "all";
            if (wanted != "all" && wanted != "weekly" && wanted != "monthly")
                return ExcelTable.Error("ExpiryType must be all, weekly or monthly");

            return AsyncTaskUtil.RunTask(nameof(oa_expiry), new object[] { symbol, exchange, kind, wanted }, async () =>
            {
                var payload = new JObject
                {
                    ["symbol"] = Arg.Str(symbol).Trim(),
                    ["exchange"] = NormExchange(exchange),
                    ["instrumenttype"] = kind
                };

                var response = await OpenAlgoClient.PostAsync("expiry", payload);

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                if (response["data"] is not JArray expiries || expiries.Count == 0)
                    return (object)ExcelTable.Error($"No {kind} expiries for {symbol} on {exchange}");

                return (object)ExpiryTable(expiries, wanted);
            })!;
        }

        // ---------------------------------------------------------------------
        // Helpers
        // ---------------------------------------------------------------------

        /// <summary>
        /// Reads one field of the instrument metadata as a single cell value.
        /// Backs the scalar lookups such as oa_lotsize and oa_token.
        /// </summary>
        private static object SymbolField(string caller, string symbol, string exchange, string field, bool asText = false)
        {
            if (!OpenAlgoClient.HasApiKey)
                return "Error: OpenAlgo API Key is not set. Use oa_api()";
            if (string.IsNullOrWhiteSpace(symbol) || string.IsNullOrWhiteSpace(exchange))
                return "Error: Symbol and Exchange are required";

            return AsyncTaskUtil.RunTask(caller, new object[] { symbol, exchange, field }, async () =>
            {
                var payload = new JObject { ["symbol"] = Arg.Str(symbol).Trim(), ["exchange"] = NormExchange(exchange) };
                var response = await OpenAlgoClient.PostAsync("symbol", payload);

                if (OpenAlgoClient.IsError(response))
                    return (object)("Error: " + OpenAlgoClient.MessageOf(response));

                if (response["data"] is not JObject data)
                    return (object)$"Error: No metadata for {symbol} on {exchange}";

                var token = data[field];
                if (token == null || token.Type == JTokenType.Null)
                    return (object)$"Error: {symbol} on {exchange} has no {field}";

                // An identifier stays text: turning "02885" into a number loses it.
                if (asText)
                    return (object)token.ToString();

                return ExcelTable.Cell(token);
            })!;
        }

        /// <summary>
        /// Builds the expiry table, marking the last expiry of each calendar month as
        /// the monthly contract and applying the requested local filter.
        ///
        /// The /expiry endpoint itself takes no expiry type, so the filter is applied
        /// here: "monthly" keeps one row per calendar month, "weekly" and "all" keep
        /// every expiry, since a weekly series includes the month end contract.
        /// </summary>
        private static object[,] ExpiryTable(JArray expiries, string wanted)
        {
            var entries = new List<(string Raw, DateTime? When)>();
            foreach (var token in expiries)
            {
                string raw = ExcelTable.Cell(token)?.ToString()?.Trim() ?? "";
                if (raw.Length == 0)
                    continue;
                entries.Add((raw, ParseExpiry(raw)));
            }

            if (entries.Count == 0)
                return ExcelTable.Error("No expiry dates returned");

            // Nearest first. Anything unparsable sorts to the end but is still listed.
            var ordered = entries.OrderBy(e => e.When ?? DateTime.MaxValue).ToList();

            // The monthly contract is the last expiry inside its calendar month.
            var monthEnd = new Dictionary<int, DateTime>();
            foreach (var entry in ordered)
            {
                if (!entry.When.HasValue)
                    continue;
                int key = entry.When.Value.Year * 100 + entry.When.Value.Month;
                if (!monthEnd.TryGetValue(key, out DateTime latest) || entry.When.Value > latest)
                    monthEnd[key] = entry.When.Value;
            }

            var rows = new List<object[]> { new object[] { "Expiry", "Date", "Type" } };
            foreach (var entry in ordered)
            {
                bool isMonthly = entry.When.HasValue
                    && monthEnd.TryGetValue(entry.When.Value.Year * 100 + entry.When.Value.Month, out DateTime latest)
                    && latest == entry.When.Value;

                if (wanted == "monthly" && !isMonthly)
                    continue;

                rows.Add(new object[]
                {
                    entry.Raw,
                    entry.When.HasValue ? (object)entry.When.Value.ToOADate() : "",
                    entry.When.HasValue ? (isMonthly ? "Monthly" : "Weekly") : ""
                });
            }

            if (rows.Count == 1)
                return ExcelTable.Error($"No {wanted} expiries in the list returned");

            return ExcelTable.FromRows(rows);
        }

        /// <summary>
        /// Parses an expiry stated as DD-MMM-YY, and the other forms brokers use.
        /// </summary>
        private static DateTime? ParseExpiry(string text)
        {
            if (DateTime.TryParseExact(text.ToUpperInvariant(), ExpiryFormats,
                    CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime exact))
                return exact;

            if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime loose))
                return loose;

            return null;
        }

        /// <summary>
        /// Maps the InstrumentType argument onto the values /expiry accepts.
        /// Returns an empty string when the argument is not recognised.
        /// </summary>
        private static string NormaliseInstrumentType(string instrumentType)
        {
            string text = (instrumentType ?? "").Trim().ToLowerInvariant();
            switch (text)
            {
                case "options":
                case "option":
                case "opt":
                case "ce":
                case "pe":
                    return "options";
                case "futures":
                case "future":
                case "fut":
                    return "futures";
                default:
                    return "";
            }
        }

        /// <summary>
        /// Appends a note row to a finished table, in the first column.
        /// </summary>
        private static object[,] WithNote(object[,] table, string note)
        {
            int rows = table.GetLength(0);
            int columns = table.GetLength(1);

            var result = new object[rows + 1, columns];
            for (int r = 0; r < rows; r++)
                for (int c = 0; c < columns; c++)
                    result[r, c] = table[r, c];

            result[rows, 0] = note;
            for (int c = 1; c < columns; c++)
                result[rows, c] = "";

            return result;
        }
    }
}
