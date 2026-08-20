using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelDna.Integration;
using Newtonsoft.Json.Linq;

namespace OpenAlgo
{
    /// <summary>
    /// Helpers that turn OpenAlgo JSON documents into the rectangular object arrays
    /// Excel spills into a range, and that read loosely typed Excel arguments.
    ///
    /// Every function in the add-in returns a table built here so that layouts,
    /// error text and number formatting stay consistent across the whole surface.
    /// </summary>
    public static class ExcelTable
    {
        /// <summary>
        /// A single cell carrying an error message.
        /// </summary>
        public static object[,] Error(string message) =>
            new object[,] { { "Error: " + message } };

        /// <summary>
        /// A single cell carrying an arbitrary value.
        /// </summary>
        public static object[,] Single(object value) =>
            new object[,] { { value } };

        /// <summary>
        /// Returns an error table when the response failed, otherwise null.
        /// </summary>
        public static object[,]? ErrorOrNull(JObject response)
        {
            if (response == null)
                return Error("Empty response");
            if (OpenAlgoClient.IsError(response))
                return Error(OpenAlgoClient.MessageOf(response));
            return null;
        }

        /// <summary>
        /// Two column key/value table for a JSON object, one property per row.
        /// </summary>
        public static object[,] KeyValue(JObject data, string keyHeader = "Field", string valueHeader = "Value")
        {
            if (data == null || data.Count == 0)
                return Error("No data available");

            var result = new object[data.Count + 1, 2];
            result[0, 0] = keyHeader;
            result[0, 1] = valueHeader;

            int row = 1;
            foreach (var property in data.Properties())
            {
                result[row, 0] = Humanise(property.Name);
                result[row, 1] = Cell(property.Value, property.Name);
                row++;
            }
            return result;
        }

        /// <summary>
        /// Two column key/value table built from explicit pairs.
        /// </summary>
        public static object[,] KeyValue(IEnumerable<KeyValuePair<string, object>> pairs, string keyHeader = "Field", string valueHeader = "Value")
        {
            var list = pairs.ToList();
            var result = new object[list.Count + 1, 2];
            result[0, 0] = keyHeader;
            result[0, 1] = valueHeader;
            for (int i = 0; i < list.Count; i++)
            {
                result[i + 1, 0] = list[i].Key;
                result[i + 1, 1] = list[i].Value;
            }
            return result;
        }

        /// <summary>
        /// Table of rows from a JSON array using an explicit column order.
        /// Column entries are JSON field names; headers are the Excel captions.
        /// </summary>
        public static object[,] Rows(JArray items, string[] fields, string[] headers, string emptyMessage = "No records")
        {
            if (fields.Length != headers.Length)
                throw new ArgumentException("fields and headers must have the same length");

            if (items == null)
                return Error(emptyMessage);

            var result = new object[Math.Max(items.Count + 1, 1), headers.Length];
            for (int col = 0; col < headers.Length; col++)
                result[0, col] = headers[col];

            for (int i = 0; i < items.Count; i++)
            {
                var item = items[i] as JObject;
                for (int col = 0; col < fields.Length; col++)
                    result[i + 1, col] = item == null ? "" : Cell(item[fields[col]], fields[col]);
            }
            return result;
        }

        /// <summary>
        /// Table of rows from a JSON array, discovering columns from the data itself.
        /// Useful where the payload shape depends on the broker.
        /// </summary>
        public static object[,] AutoRows(JArray items, string emptyMessage = "No records")
        {
            if (items == null || items.Count == 0)
                return Error(emptyMessage);

            var fields = new List<string>();
            foreach (var item in items.OfType<JObject>())
                foreach (var property in item.Properties())
                    if (!fields.Contains(property.Name))
                        fields.Add(property.Name);

            if (fields.Count == 0)
                return Error(emptyMessage);

            var result = new object[items.Count + 1, fields.Count];
            for (int col = 0; col < fields.Count; col++)
                result[0, col] = Humanise(fields[col]);

            for (int i = 0; i < items.Count; i++)
            {
                var item = items[i] as JObject;
                for (int col = 0; col < fields.Count; col++)
                    result[i + 1, col] = item == null ? "" : Cell(item[fields[col]], fields[col]);
            }
            return result;
        }

        /// <summary>
        /// Builds a table from a list of equal length rows.
        /// </summary>
        public static object[,] FromRows(List<object[]> rows)
        {
            if (rows == null || rows.Count == 0)
                return Error("No records");

            int width = rows.Max(r => r.Length);
            var result = new object[rows.Count, width];
            for (int r = 0; r < rows.Count; r++)
                for (int c = 0; c < width; c++)
                    result[r, c] = c < rows[r].Length ? (rows[r][c] ?? "") : "";
            return result;
        }

        /// <summary>
        /// Standard two row acknowledgement used by action style functions.
        /// </summary>
        public static object[,] Acknowledgement(JObject response, params string[] extraFields)
        {
            var pairs = new List<KeyValuePair<string, object>>
            {
                new("Status", response["status"]?.ToString() ?? "unknown")
            };

            foreach (var field in extraFields)
            {
                var token = response[field];
                if (token != null && token.Type != JTokenType.Null)
                    pairs.Add(new KeyValuePair<string, object>(Humanise(field), Cell(token, field)));
            }

            string? message = response["message"]?.ToString();
            if (!string.IsNullOrWhiteSpace(message))
                pairs.Add(new KeyValuePair<string, object>("Message", message));

            string? mode = response["mode"]?.ToString();
            if (!string.IsNullOrWhiteSpace(mode))
                pairs.Add(new KeyValuePair<string, object>("Mode", mode));

            return KeyValue(pairs);
        }

        /// <summary>
        /// Converts a JSON token into a value Excel can display.
        /// Numbers stay numeric so formulas can act on them; objects and arrays are
        /// flattened to compact JSON rather than showing a type name.
        /// </summary>
        public static object Cell(JToken? token) => Cell(token, null);

        /// <summary>
        /// Converts a JSON token into a value Excel can display, using the field name to
        /// decide whether a numeric looking value is a quantity or an identifier.
        /// </summary>
        public static object Cell(JToken? token, string? field)
        {
            if (token == null)
                return "";

            bool identifier = IsIdentifierField(field);

            switch (token.Type)
            {
                case JTokenType.Null:
                case JTokenType.Undefined:
                    return "";
                case JTokenType.Integer:
                case JTokenType.Float:
                    // An id that arrives as a JSON number is still an id. Excel stores every
                    // number as a double and shows anything past 11 digits in scientific
                    // notation, so a 15 digit order id would read as 2.5082E+14.
                    if (identifier)
                        return token.ToString();
                    return token.ToObject<double>();
                case JTokenType.Boolean:
                    return token.ToObject<bool>();
                case JTokenType.Date:
                    return token.ToObject<DateTime>().ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                case JTokenType.Object:
                case JTokenType.Array:
                    return token.ToString(Newtonsoft.Json.Formatting.None);
                default:
                    string text = token.ToString();
                    if (identifier)
                        return text;
                    // Brokers often send prices and quantities as strings. Present those as
                    // numbers so formulas can act on them, but only when the conversion is
                    // lossless: "007" and "1.230" must not come back as 7 and 1.23.
                    if (IsLosslessNumber(text, out double number))
                        return number;
                    return text;
            }
        }

        /// <summary>
        /// True when a field holds an identifier rather than a measurement. Identifiers must
        /// stay text: BSE scrip codes, order ids, instrument tokens and phone numbers are all
        /// digit strings that lose their meaning, their leading zeros or their precision once
        /// Excel treats them as numbers.
        /// </summary>
        private static bool IsIdentifierField(string? field)
        {
            if (string.IsNullOrWhiteSpace(field))
                return false;

            string name = field.Trim().ToLowerInvariant();

            if (IdentifierFields.Contains(name))
                return true;

            // Suffix rules. "bid" is deliberately not matched by the "id" rule.
            return name.EndsWith("_id", StringComparison.Ordinal)
                || name.EndsWith("orderid", StringComparison.Ordinal)
                || name.EndsWith("_token", StringComparison.Ordinal)
                || name.EndsWith("symbol", StringComparison.Ordinal);
        }

        private static readonly HashSet<string> IdentifierFields = new(StringComparer.OrdinalIgnoreCase)
        {
            "id", "orderid", "order_id", "orderno", "order_no",
            "triggerid", "trigger_id", "token", "instrument_token", "exchange_token",
            "symbol", "brsymbol", "tradingsymbol", "underlying",
            "tradeid", "trade_id", "exchange_order_id", "brorderid", "norenordno",
            "telegram_id", "telegramid", "chat_id", "chatid", "user_id", "userid",
            "phone", "phones", "username", "filename", "isin", "expiry"
        };

        /// <summary>
        /// Parses a numeric string only when the number renders back to the same text, so a
        /// zero padded code or a trailing zero decimal is preserved rather than normalised.
        /// </summary>
        private static bool IsLosslessNumber(string text, out double number)
        {
            number = 0;
            string trimmed = text.Trim();
            if (trimmed.Length == 0)
                return false;

            // A leading zero followed by another digit marks a zero padded code such as
            // "007" or "0123", where the zeros carry meaning. Everything else that parses
            // is a genuine quantity.
            //
            // Trailing zeros after the decimal point are NOT a reason to reject: brokers
            // routinely format prices to a fixed two decimals, and treating "1234.50" as
            // text while "1234.5" becomes a number would mix types in one price column
            // and make SUM and charts silently skip rows.
            string digits = trimmed.TrimStart('+', '-');
            if (digits.Length > 1 && digits[0] == '0' && char.IsDigit(digits[1]))
                return false;

            return double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out number);
        }

        /// <summary>
        /// Turns a JSON field name such as "prev_close" into "Prev Close".
        /// </summary>
        public static string Humanise(string field)
        {
            if (string.IsNullOrWhiteSpace(field))
                return "";

            var words = field.Replace('_', ' ').Split(' ', StringSplitOptions.RemoveEmptyEntries);
            return string.Join(" ", words.Select(w =>
                w.Length <= 3 && w.All(char.IsUpper)
                    ? w
                    : char.ToUpperInvariant(w[0]) + w.Substring(1)));
        }
    }

    /// <summary>
    /// Reads optional Excel arguments, which arrive as ExcelMissing, ExcelEmpty,
    /// ExcelError, double, string or bool depending on what the user typed.
    /// </summary>
    public static class Arg
    {
        /// <summary>
        /// True when the caller left the argument out or pointed it at an empty cell.
        /// </summary>
        public static bool IsMissing(object? value) =>
            value == null
            || value is ExcelMissing
            || value is ExcelEmpty
            || value is ExcelError
            || (value is string s && string.IsNullOrWhiteSpace(s));

        /// <summary>
        /// Reads a string argument, falling back to a default when omitted.
        /// </summary>
        public static string Str(object? value, string fallback = "")
        {
            if (IsMissing(value))
                return fallback;
            if (value is double d)
                return d.ToString("0.############", CultureInfo.InvariantCulture);
            return value!.ToString() ?? fallback;
        }

        /// <summary>
        /// Reads a numeric argument, falling back to a default when omitted or unparsable.
        /// </summary>
        public static double Num(object? value, double fallback = 0)
        {
            if (IsMissing(value))
                return fallback;
            if (value is double d)
                return d;
            if (value is bool b)
                return b ? 1 : 0;
            return double.TryParse(value!.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out double parsed)
                ? parsed
                : fallback;
        }

        /// <summary>
        /// Reads an integer argument.
        /// </summary>
        public static int Int(object? value, int fallback = 0) => (int)Math.Round(Num(value, fallback));

        /// <summary>
        /// Reads a boolean argument. Accepts TRUE/FALSE, 1/0 and yes/no.
        /// </summary>
        public static bool Bool(object? value, bool fallback = false)
        {
            if (IsMissing(value))
                return fallback;
            if (value is bool b)
                return b;
            if (value is double d)
                return Math.Abs(d) > double.Epsilon;

            string text = value!.ToString()!.Trim();
            if (bool.TryParse(text, out bool parsed))
                return parsed;
            if (text.Equals("yes", StringComparison.OrdinalIgnoreCase) || text == "1")
                return true;
            if (text.Equals("no", StringComparison.OrdinalIgnoreCase) || text == "0")
                return false;
            return fallback;
        }

        /// <summary>
        /// Reads a date argument. Accepts a real Excel date serial or a YYYY-MM-DD string.
        /// </summary>
        public static string Date(object? value, string fallback = "")
        {
            if (IsMissing(value))
                return fallback;
            if (value is double serial)
                return DateTime.FromOADate(serial).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

            string text = value!.ToString()!.Trim();
            if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsed))
                return parsed.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            return text;
        }

        /// <summary>
        /// Flattens an Excel range argument into a list of non empty cells, row by row.
        /// </summary>
        public static List<object> Flatten(object? value)
        {
            var cells = new List<object>();
            if (IsMissing(value))
                return cells;

            if (value is object[,] grid)
            {
                int rows = grid.GetLength(0);
                int cols = grid.GetLength(1);
                for (int r = 0; r < rows; r++)
                    for (int c = 0; c < cols; c++)
                        if (!IsMissing(grid[r, c]))
                            cells.Add(grid[r, c]);
                return cells;
            }

            if (value is object[] array)
            {
                foreach (var cell in array)
                    if (!IsMissing(cell))
                        cells.Add(cell);
                return cells;
            }

            cells.Add(value!);
            return cells;
        }

        /// <summary>
        /// Reads a symbol list from an Excel range.
        ///
        /// Accepts a single column of symbols (the default exchange is applied),
        /// or two columns of symbol and exchange. A header row is detected and skipped.
        /// </summary>
        public static List<(string Symbol, string Exchange)> SymbolPairs(object? value, string defaultExchange = "NSE")
        {
            var pairs = new List<(string, string)>();
            if (IsMissing(value))
                return pairs;

            if (value is object[,] grid)
            {
                int rows = grid.GetLength(0);
                int cols = grid.GetLength(1);

                // A first row reading "symbol" is a header the user included by accident.
                int startRow = 0;
                if (rows > 1 && Str(grid[0, 0]).Trim().Equals("symbol", StringComparison.OrdinalIgnoreCase))
                    startRow = 1;

                // A 1 x N range is a horizontal list of symbols.
                if (rows == 1 && cols > 2)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        string symbol = Str(grid[0, c]).Trim();
                        if (symbol.Length > 0)
                            pairs.Add((symbol, defaultExchange));
                    }
                    return pairs;
                }

                for (int r = startRow; r < rows; r++)
                {
                    string symbol = Str(grid[r, 0]).Trim();
                    if (symbol.Length == 0)
                        continue;
                    string exchange = cols > 1 ? Str(grid[r, 1]).Trim() : "";
                    pairs.Add((symbol, exchange.Length > 0 ? exchange : defaultExchange));
                }
                return pairs;
            }

            string single = Str(value).Trim();
            if (single.Length > 0)
            {
                // Allow the compact "NSE:RELIANCE" form in a single cell.
                int colon = single.IndexOf(':');
                if (colon > 0)
                    pairs.Add((single.Substring(colon + 1).Trim(), single.Substring(0, colon).Trim()));
                else
                    pairs.Add((single, defaultExchange));
            }
            return pairs;
        }
    }
}
