using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelDna.Integration;
using ExcelDna.Registration.Utils;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace OpenAlgo
{
    /// <summary>
    /// Analyzer (sandbox) control, the market calendar, connectivity and preference
    /// helpers, and the generic escape hatches that reach any endpoint the add-in does
    /// not wrap by name.
    /// </summary>
    public static class UtilityApi
    {
        private const string AnalyzerCategory = "OpenAlgo Analyzer";
        private const string CalendarCategory = "OpenAlgo Calendar";
        private const string UtilityCategory = "OpenAlgo Utility";

        // ------------------------------------------------------------------
        // Analyzer
        // ------------------------------------------------------------------

        /// <summary>
        /// Reports whether the server is in analyzer (sandbox) mode or live mode, and how
        /// many orders the analyzer has logged.
        /// </summary>
        [ExcelFunction(
            Name = "oa_analyzer",
            Description = "Analyzer status: whether orders are simulated in the sandbox or sent to the live broker, and how many orders the analyzer has logged. Check this before arming a strategy.",
            Category = AnalyzerCategory)]
        public static object oa_analyzer()
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");

            return AsyncTaskUtil.RunTask(nameof(oa_analyzer), new object[] { }, async () =>
            {
                var response = await OpenAlgoClient.PostAsync("analyzer", new JObject());

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                if (response["data"] is not JObject data)
                    return (object)ExcelTable.Error("No analyzer status returned");

                return (object)ModeTable(data);
            })!;
        }

        /// <summary>
        /// Switches the server between analyzer (sandbox) mode and live mode. In live mode
        /// orders reach the real broker, so the returned table states the resulting mode
        /// first.
        /// </summary>
        [ExcelFunction(
            Name = "oa_analyzer_toggle",
            Description = "Switches between analyzer (sandbox) and live mode. WARNING: switching to live means every order function sends real orders to the broker. The returned table states the resulting mode in its first row.",
            Category = AnalyzerCategory)]
        public static object oa_analyzer_toggle(
            [ExcelArgument(Name = "Mode", Description = "TRUE or \"analyze\" routes orders to the sandbox. FALSE or \"live\" sends them to the real broker.")] object mode)
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");

            if (!TryReadMode(mode, out bool analyze))
                return ExcelTable.Error("Mode must be TRUE, FALSE, \"analyze\" or \"live\".");

            return AsyncTaskUtil.RunTask(nameof(oa_analyzer_toggle), new object[] { analyze }, async () =>
            {
                var response = await OpenAlgoClient.PostAsync("analyzer/toggle", new JObject { ["mode"] = analyze });

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                if (response["data"] is not JObject data)
                    return (object)ExcelTable.Error("No analyzer status returned");

                return (object)ModeTable(data);
            })!;
        }

        /// <summary>
        /// Sandbox profit and loss per symbol, with the totals the endpoint carries at the
        /// top level. The endpoint is rejected with HTTP 400 while the server is in live
        /// mode, which this function reports in plain language.
        /// </summary>
        [ExcelFunction(
            Name = "oa_pnl_symbols",
            Description = "Sandbox P&L per symbol with realised, unrealised and today totals. Analyzer mode only: in live mode the API answers HTTP 400 and this function says so. Switch with oa_analyzer_toggle(TRUE).",
            Category = AnalyzerCategory)]
        public static object oa_pnl_symbols()
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");

            return AsyncTaskUtil.RunTask(nameof(oa_pnl_symbols), new object[] { }, async () =>
            {
                var response = await OpenAlgoClient.PostAsync("pnl/symbols", new JObject());

                if (OpenAlgoClient.IsError(response))
                {
                    string message = OpenAlgoClient.MessageOf(response);
                    if (message.Contains("HTTP 400", StringComparison.OrdinalIgnoreCase))
                        return (object)ExcelTable.Error(
                            "Sandbox P&L is available only while analyzer mode is on. The server rejected the request because it is in live mode (" +
                            message + "). Switch with oa_analyzer_toggle(TRUE).");
                    return (object)ExcelTable.Error(message);
                }

                if (response["data"] is not JArray items)
                    return (object)ExcelTable.Error("No sandbox P&L data returned");

                string[] fields = { "symbol", "exchange", "product", "quantity", "pnl", "unrealized_pnl", "today_realized_pnl", "total_pnl_today" };

                var rows = new List<object[]>
                {
                    new object[] { "Symbol", "Exchange", "Product", "Quantity", "PnL", "Unrealized PnL", "Today Realized PnL", "Total PnL Today" }
                };

                foreach (var item in items.OfType<JObject>())
                {
                    var row = new object[fields.Length];
                    for (int col = 0; col < fields.Length; col++)
                        row[col] = ExcelTable.Cell(item[fields[col]]);
                    rows.Add(row);
                }

                rows.Add(new object[]
                {
                    "TOTAL", "", "", "",
                    ExcelTable.Cell(response["total_pnl"]),
                    ExcelTable.Cell(response["total_unrealized_pnl"]),
                    ExcelTable.Cell(response["total_today_realized_pnl"]),
                    ExcelTable.Cell(response["total_pnl_today"])
                });

                rows.Add(new object[] { "Mode", ExcelTable.Cell(response["mode"]), "", "", "", "", "", "" });

                return (object)ExcelTable.FromRows(rows);
            })!;
        }

        // ------------------------------------------------------------------
        // Market calendar
        // ------------------------------------------------------------------

        /// <summary>
        /// Market holidays for a year, including the special sessions some exchanges run
        /// on days the rest of the market is closed.
        /// </summary>
        [ExcelFunction(
            Name = "oa_holidays",
            Description = "Market holidays for a year: date, description, holiday type, the exchanges that are closed and any special sessions. Times shown in IST.",
            Category = CalendarCategory)]
        public static object oa_holidays(
            [ExcelArgument(Name = "Year", Description = "Year between 2020 and 2050. Omit for the current year.")] object? yearOptional = null,
            [ExcelArgument(Name = "Exchange", Description = "Optional filter, for example NSE or MCX. Keeps only the days on which that exchange is closed or runs a special session.")] object? exchangeOptional = null)
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");

            int year = Arg.IsMissing(yearOptional) ? 0 : Arg.Int(yearOptional, 0);
            if (year != 0 && (year < 2020 || year > 2050))
                return ExcelTable.Error("Year must be between 2020 and 2050.");

            string exchange = Arg.Str(exchangeOptional).Trim().ToUpperInvariant();

            return AsyncTaskUtil.RunTask(nameof(oa_holidays), new object[] { (double)year, exchange }, async () =>
            {
                var payload = new JObject();
                if (year != 0)
                    payload["year"] = year;

                var response = await OpenAlgoClient.PostAsync("market/holidays", payload);

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                if (response["data"] is not JArray items)
                    return (object)ExcelTable.Error("No holiday data returned");

                var rows = new List<object[]>
                {
                    new object[] { "Date", "Description", "Holiday Type", "Closed Exchanges", "Special Sessions (IST)" }
                };

                foreach (var item in items.OfType<JObject>())
                {
                    var closedList = UtilitySupport.StringList(item["closed_exchanges"] as JArray);
                    var sessions = item["open_exchanges"] as JArray;

                    if (exchange.Length > 0)
                    {
                        bool closed = closedList.Any(e => e.Equals(exchange, StringComparison.OrdinalIgnoreCase));
                        bool special = sessions != null && sessions.OfType<JObject>()
                            .Any(s => string.Equals(s["exchange"]?.ToString(), exchange, StringComparison.OrdinalIgnoreCase));
                        if (!closed && !special)
                            continue;
                    }

                    rows.Add(new object[]
                    {
                        ExcelTable.Cell(item["date"]),
                        ExcelTable.Cell(item["description"]),
                        ExcelTable.Cell(item["holiday_type"]),
                        string.Join(", ", closedList),
                        UtilitySupport.DescribeSessions(sessions)
                    });
                }

                return (object)ExcelTable.FromRows(rows);
            })!;
        }

        /// <summary>
        /// Trading sessions for a date across every exchange. An empty schedule means the
        /// market is closed on that date.
        /// </summary>
        [ExcelFunction(
            Name = "oa_timings",
            Description = "Trading sessions for a date, one row per exchange, with IST start and end times and the raw epoch milliseconds. An empty schedule means the market is closed that day.",
            Category = CalendarCategory)]
        public static object oa_timings(
            [ExcelArgument(Name = "Date", Description = "Date as a cell date or YYYY-MM-DD text, between 2020-01-01 and 2050-12-31. Omit for today.")] object? dateOptional = null)
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");

            string date = Arg.Date(dateOptional, DateTime.Today.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));

            return AsyncTaskUtil.RunTask(nameof(oa_timings), new object[] { date }, async () =>
            {
                var response = await OpenAlgoClient.PostAsync("market/timings", new JObject { ["date"] = date });

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                if (response["data"] is not JArray items)
                    return (object)ExcelTable.Error("No timings data returned");

                if (items.Count == 0)
                {
                    return (object)ExcelTable.KeyValue(new List<KeyValuePair<string, object>>
                    {
                        new("Date", date),
                        new("Status", "Closed. No trading session is scheduled, so the date is a weekend or a market holiday.")
                    });
                }

                var rows = new List<object[]>
                {
                    new object[] { "Exchange", "Session Start (IST)", "Session End (IST)", "Start Epoch Ms", "End Epoch Ms" }
                };

                foreach (var item in items.OfType<JObject>())
                {
                    rows.Add(new object[]
                    {
                        ExcelTable.Cell(item["exchange"]),
                        UtilitySupport.FormatEpoch(item["start_time"], "yyyy-MM-dd HH:mm:ss"),
                        UtilitySupport.FormatEpoch(item["end_time"], "yyyy-MM-dd HH:mm:ss"),
                        ExcelTable.Cell(item["start_time"]),
                        ExcelTable.Cell(item["end_time"])
                    });
                }

                return (object)ExcelTable.FromRows(rows);
            })!;
        }

        /// <summary>
        /// Convenience test for a closed market. Derived from /market/timings for the
        /// date, because OpenAlgo has no /checkholiday endpoint: an empty schedule means
        /// the whole market is shut, and a missing exchange means that exchange is shut.
        /// </summary>
        [ExcelFunction(
            Name = "oa_isholiday",
            Description = "TRUE when the market is closed on the date. Derived from /market/timings, since OpenAlgo has no /checkholiday endpoint: an empty session schedule means a weekend or holiday, and with an Exchange given, TRUE means that exchange has no session that day.",
            Category = CalendarCategory)]
        public static object oa_isholiday(
            [ExcelArgument(Name = "Date", Description = "Date as a cell date or YYYY-MM-DD text.")] object date,
            [ExcelArgument(Name = "Exchange", Description = "Optional exchange, for example NSE or MCX. Omit to ask whether every exchange is closed.")] object? exchangeOptional = null)
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");

            if (Arg.IsMissing(date))
                return ExcelTable.Error("Date is required.");

            string day = Arg.Date(date);
            string exchange = Arg.Str(exchangeOptional).Trim().ToUpperInvariant();

            return AsyncTaskUtil.RunTask(nameof(oa_isholiday), new object[] { day, exchange }, async () =>
            {
                var response = await OpenAlgoClient.PostAsync("market/timings", new JObject { ["date"] = day });

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                if (response["data"] is not JArray items)
                    return (object)ExcelTable.Error("No timings data returned");

                // No session at all: weekend or a full market holiday.
                if (items.Count == 0)
                    return (object)true;

                if (exchange.Length == 0)
                    return (object)false;

                bool open = items.OfType<JObject>()
                    .Any(s => string.Equals(s["exchange"]?.ToString(), exchange, StringComparison.OrdinalIgnoreCase));

                return (object)!open;
            })!;
        }

        // ------------------------------------------------------------------
        // Utility and preferences
        // ------------------------------------------------------------------

        /// <summary>
        /// Verifies that the configured API key resolves to an active broker session.
        /// </summary>
        [ExcelFunction(
            Name = "oa_ping",
            Description = "Verifies the API key resolves to an active broker session and names the broker. This is an authenticated check, not a process health probe: a revoked key or a logged out broker fails it.",
            Category = UtilityCategory)]
        public static object oa_ping()
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");

            return AsyncTaskUtil.RunTask(nameof(oa_ping), new object[] { }, async () =>
            {
                var response = await OpenAlgoClient.PostAsync("ping", new JObject());

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                if (response["data"] is not JObject data)
                    return (object)ExcelTable.Error("No data returned");

                var pairs = new List<KeyValuePair<string, object>>
                {
                    new("Status", response["status"]?.ToString() ?? "unknown"),
                    new("Message", ExcelTable.Cell(data["message"])),
                    new("Broker", ExcelTable.Cell(data["broker"]))
                };
                UtilitySupport.AppendRemaining(pairs, data, "message", "broker");

                return (object)ExcelTable.KeyValue(pairs);
            })!;
        }

        /// <summary>
        /// Reads the stored chart workspace preferences for the API key.
        /// </summary>
        [ExcelFunction(
            Name = "oa_chart",
            Description = "Chart workspace preferences stored for this API key. The first column holds the exact preference key, so it can be fed straight back into oa_chart_set.",
            Category = UtilityCategory)]
        public static object oa_chart()
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");

            return AsyncTaskUtil.RunTask(nameof(oa_chart), new object[] { }, async () =>
            {
                var response = await OpenAlgoClient.GetAsync("chart");

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                if (response["data"] is not JObject data)
                    return (object)ExcelTable.Error("No chart preferences returned");

                if (data.Count == 0)
                    return (object)ExcelTable.Error("No chart preferences are stored for this API key");

                var pairs = data.Properties()
                    .Select(p => new KeyValuePair<string, object>(p.Name, ExcelTable.Cell(p.Value)))
                    .ToList();

                return (object)ExcelTable.KeyValue(pairs, "Preference", "Value");
            })!;
        }

        /// <summary>
        /// Updates one chart workspace preference. Values that parse as JSON are sent as
        /// JSON so nested layouts survive the round trip; anything else is sent as text.
        /// </summary>
        [ExcelFunction(
            Name = "oa_chart_set",
            Description = "Updates one chart preference. A value that parses as JSON is sent as JSON (objects, arrays, numbers, booleans), anything else is sent as text. Keys are limited to 50 characters.",
            Category = UtilityCategory)]
        public static object oa_chart_set(
            [ExcelArgument(Name = "Key", Description = "Preference key, for example tv_theme or tv_chart_layout. Max 50 characters.")] object key,
            [ExcelArgument(Name = "Value", Description = "New value. JSON text such as {\"interval\":\"15m\"} is stored as JSON; plain text such as dark is stored as a string.")] object value)
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");

            string name = Arg.Str(key).Trim();
            if (name.Length == 0)
                return ExcelTable.Error("Key is required.");
            if (name.Length > 50)
                return ExcelTable.Error("Key must be 50 characters or fewer.");

            JToken token = UtilitySupport.ValueToken(value);
            string tokenJson = token.ToString(Formatting.None);

            return AsyncTaskUtil.RunTask(nameof(oa_chart_set), new object[] { name, tokenJson }, async () =>
            {
                var payload = new JObject { [name] = token };
                var response = await OpenAlgoClient.PostAsync("chart", payload);

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                var pairs = UtilitySupport.StatusPairs(response);
                pairs.Add(new KeyValuePair<string, object>("Preference", name));
                pairs.Add(new KeyValuePair<string, object>("Value", tokenJson));

                return (object)ExcelTable.KeyValue(pairs);
            })!;
        }

        /// <summary>
        /// Reports the add-in build and the configuration it is using. Makes no network
        /// call, so it answers even when the server is unreachable.
        /// </summary>
        [ExcelFunction(
            Name = "oa_version",
            Description = "Add-in version and the configuration in force: host URL, API version, REST base, WebSocket URL, timeouts and whether an API key is set. Makes no network call, so quote this first when reporting a problem.",
            Category = UtilityCategory)]
        public static object oa_version()
        {
            var pairs = new List<KeyValuePair<string, object>>
            {
                new("Add-in Version", OpenAlgoConfig.AddInVersion),
                new("Host URL", OpenAlgoConfig.HostUrl),
                new("API Version", OpenAlgoConfig.Version),
                new("REST Base URL", OpenAlgoClient.BaseUrl),
                new("WebSocket URL", OpenAlgoConfig.WebSocketUrl),
                new("API Key Set", OpenAlgoClient.HasApiKey),
                new("Timeout Seconds", (double)OpenAlgoConfig.TimeoutSeconds),
                new("Stream Throttle Ms", (double)OpenAlgoConfig.StreamThrottleMs),
                new("Trading Enabled", OpenAlgoConfig.TradingEnabled)
            };

            return ExcelTable.KeyValue(pairs);
        }

        /// <summary>
        /// Reads or sets the trading arm switch that order placing functions consult
        /// before they fire. Deliberately not volatile, so it does not re-evaluate on
        /// every recalculation.
        /// </summary>
        [ExcelFunction(
            Name = "oa_trading_enabled",
            Description = "Reads the trading arm switch, or sets it when passed TRUE or FALSE. Order functions refuse to send while it is FALSE, which stops a recalculation from re-placing orders. The setting lives for this Excel session only.",
            Category = UtilityCategory)]
        public static object oa_trading_enabled(
            [ExcelArgument(Name = "Enable", Description = "TRUE arms order placing functions, FALSE disarms them. Omit to read the current state without changing it.")] object? enableOptional = null)
        {
            if (!Arg.IsMissing(enableOptional))
            {
                if (!UtilitySupport.TryReadBool(enableOptional, out bool enable))
                    return ExcelTable.Error("Enable must be TRUE or FALSE.");
                OpenAlgoConfig.TradingEnabled = enable;
            }

            var pairs = new List<KeyValuePair<string, object>>
            {
                new("Trading Enabled", OpenAlgoConfig.TradingEnabled),
                new("State", OpenAlgoConfig.TradingEnabled
                    ? "ARMED. Order functions will send orders."
                    : "DISARMED. Order functions will refuse to send."),
                new("Note", "This flag is held in memory for the current Excel session and resets to armed when Excel restarts.")
            };

            return ExcelTable.KeyValue(pairs);
        }

        /// <summary>
        /// Issues a raw call against the configured OpenAlgo host and returns the JSON
        /// response text in one cell. Covers endpoints that have no wrapper of their own.
        /// </summary>
        [ExcelFunction(
            Name = "oa_request",
            Description = "Raw call to any OpenAlgo endpoint, returning the JSON response in a single cell. Use with oa_json to pull values out. The API key is added for you: do not put apikey in the body.",
            Category = UtilityCategory)]
        public static object oa_request(
            [ExcelArgument(Name = "Method", Description = "GET or POST.")] object method,
            [ExcelArgument(Name = "Path", Description = "Path relative to the API base, for example telegram/stats or telegram/stats?days=30. Leading slashes are ignored.")] object path,
            [ExcelArgument(Name = "JSON Body", Description = "POST only: the request body as JSON text, for example {\"username\":\"rajan\"}. Omit for GET, whose parameters belong in the path.")] object? jsonBodyOptional = null)
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");

            string verb = Arg.Str(method).Trim().ToUpperInvariant();
            if (verb != "GET" && verb != "POST")
                return ExcelTable.Error("Method must be GET or POST.");

            string target = Arg.Str(path).Trim();
            if (target.Length == 0)
                return ExcelTable.Error("Path is required, for example telegram/stats.");

            string bodyText = Arg.Str(jsonBodyOptional).Trim();

            if (verb == "GET" && bodyText.Length > 0)
                return ExcelTable.Error("A JSON body applies to POST only. Put GET parameters in the path, for example telegram/stats?days=30.");

            JObject body = new JObject();
            if (bodyText.Length > 0)
            {
                try
                {
                    body = JObject.Parse(bodyText);
                }
                catch (JsonException ex)
                {
                    return ExcelTable.Error("JSON Body is not a JSON object: " + ex.Message);
                }
            }

            return AsyncTaskUtil.RunTask(nameof(oa_request), new object[] { verb, target, bodyText }, async () =>
            {
                JObject response;
                if (verb == "GET")
                {
                    var (route, query) = UtilitySupport.SplitQuery(target);
                    response = await OpenAlgoClient.GetAsync(route, query);
                }
                else
                {
                    response = await OpenAlgoClient.PostAsync(target, body);
                }

                return (object)ExcelTable.Single(UtilitySupport.TruncateForCell(response.ToString(Formatting.None)));
            })!;
        }

        /// <summary>
        /// Extracts a value from JSON text using a JSONPath expression, so oa_request
        /// output becomes usable in formulas. Makes no network call.
        /// </summary>
        [ExcelFunction(
            Name = "oa_json",
            Description = "Pulls a value out of JSON text with a JSONPath expression such as data.broker or data[0].symbol. Objects and arrays come back as compact JSON, scalars as numbers, booleans or text.",
            Category = UtilityCategory)]
        public static object oa_json(
            [ExcelArgument(Name = "JSON Text", Description = "JSON text, typically a cell holding oa_request output.")] object jsonText,
            [ExcelArgument(Name = "Path", Description = "JSONPath expression, for example data.broker, data[0].symbol or $.data.statistics.total_open_orders.")] object path)
        {
            string text = Arg.Str(jsonText).Trim();
            if (text.Length == 0)
                return ExcelTable.Error("JSON Text is required.");

            string expression = Arg.Str(path).Trim();
            if (expression.Length == 0)
                return ExcelTable.Error("Path is required, for example data.broker.");

            JToken root;
            try
            {
                root = JToken.Parse(text);
            }
            catch (JsonException ex)
            {
                return ExcelTable.Error("JSON Text is not valid JSON: " + ex.Message);
            }

            JToken? token;
            try
            {
                token = root.SelectToken(expression);
            }
            catch (JsonException ex)
            {
                return ExcelTable.Error("Path is not a valid JSONPath expression: " + ex.Message);
            }

            if (token == null || token.Type == JTokenType.Undefined)
                return ExcelTable.Error("No value at path " + expression);

            if (token.Type == JTokenType.Object || token.Type == JTokenType.Array)
                return ExcelTable.Single(UtilitySupport.TruncateForCell(token.ToString(Formatting.None)));

            return ExcelTable.Single(ExcelTable.Cell(token));
        }

        // ------------------------------------------------------------------
        // Helpers
        // ------------------------------------------------------------------

        /// <summary>
        /// Key/value table for an analyzer payload, leading with the resulting mode and
        /// what it means for order routing.
        /// </summary>
        private static object[,] ModeTable(JObject data)
        {
            string mode = data["mode"]?.ToString() ?? "";
            bool analyze = data["analyze_mode"]?.ToObject<bool?>()
                           ?? mode.Equals("analyze", StringComparison.OrdinalIgnoreCase);
            if (mode.Length == 0)
                mode = analyze ? "analyze" : "live";

            var pairs = new List<KeyValuePair<string, object>>
            {
                new("Mode", mode.ToUpperInvariant()),
                new("Orders Go To", analyze
                    ? "The sandbox simulator. No order reaches the broker and no real money is at risk."
                    : "The live broker. Orders are real and real money is at risk."),
                new("Analyze Mode", analyze)
            };

            UtilitySupport.AppendRemaining(pairs, data, "mode", "analyze_mode");

            return ExcelTable.KeyValue(pairs);
        }

        /// <summary>
        /// Reads the analyzer mode argument, which accepts TRUE/FALSE as well as the
        /// words analyze and live.
        /// </summary>
        private static bool TryReadMode(object? value, out bool analyze)
        {
            analyze = false;
            if (Arg.IsMissing(value))
                return false;

            string text = Arg.Str(value).Trim();
            if (text.Equals("analyze", StringComparison.OrdinalIgnoreCase)
                || text.Equals("analyzer", StringComparison.OrdinalIgnoreCase)
                || text.Equals("sandbox", StringComparison.OrdinalIgnoreCase))
            {
                analyze = true;
                return true;
            }
            if (text.Equals("live", StringComparison.OrdinalIgnoreCase)
                || text.Equals("real", StringComparison.OrdinalIgnoreCase))
            {
                analyze = false;
                return true;
            }

            return UtilitySupport.TryReadBool(value, out analyze);
        }
    }

    /// <summary>
    /// Messaging services: the single public WhatsApp endpoint and the full Telegram
    /// REST surface.
    /// </summary>
    public static class MessagingApi
    {
        private const string MessagingCategory = "OpenAlgo Messaging";

        /// <summary>
        /// Preference keys the Telegram preferences endpoint accepts, with the type each
        /// one is sent as.
        /// </summary>
        private static readonly Dictionary<string, string> PreferenceKeys = new(StringComparer.OrdinalIgnoreCase)
        {
            ["order_notifications"] = "bool",
            ["trade_notifications"] = "bool",
            ["pnl_notifications"] = "bool",
            ["daily_summary"] = "bool",
            ["summary_time"] = "text",
            ["language"] = "text",
            ["timezone"] = "text"
        };

        /// <summary>
        /// Configuration keys the Telegram config endpoint accepts, with the type each
        /// one is sent as.
        /// </summary>
        private static readonly Dictionary<string, string> ConfigKeys = new(StringComparer.OrdinalIgnoreCase)
        {
            ["token"] = "text",
            ["bot_token"] = "text",
            ["webhook_url"] = "text",
            ["polling_mode"] = "bool",
            ["broadcast_enabled"] = "bool",
            ["rate_limit_per_minute"] = "int"
        };

        // ------------------------------------------------------------------
        // WhatsApp
        // ------------------------------------------------------------------

        /// <summary>
        /// Sends a WhatsApp message, optionally with an image or a document, to the
        /// paired device itself, a linked OpenAlgo user, one phone number or up to five.
        /// </summary>
        [ExcelFunction(
            Name = "oa_whatsapp_notify",
            Description = "Sends a WhatsApp text, image or document to yourself, a linked username, one phone number or up to 5. This is the whole public WhatsApp REST surface: pairing, config, users, broadcast and stats are admin-only behind the web session. Attachments are read from the OpenAlgo server filesystem (never uploaded from Excel) and must sit inside WHATSAPP_ATTACHMENT_ROOTS. Limit 30 calls per minute.",
            Category = MessagingCategory)]
        public static object oa_whatsapp_notify(
            [ExcelArgument(Name = "Message", Description = "Text body, max 4096 characters. May be blank when an Image Path or Document Path is given.")] object? message = null,
            [ExcelArgument(Name = "Recipient", Description = "Blank or \"self\" for the paired device. A linked OpenAlgo username, one E.164 phone number such as 919876543210, or a range or comma separated list of up to 5 numbers.")] object? recipientOptional = null,
            [ExcelArgument(Name = "Recipient Type", Description = "Forces the recipient form: self, username, phone or phones. Omit to infer it from Recipient. Exactly one form is sent; the API does not allow combining them.")] object? recipientTypeOptional = null,
            [ExcelArgument(Name = "Image Path", Description = "Server-local path to an image, inside the WHATSAPP_ATTACHMENT_ROOTS allowlist.")] object? imagePathOptional = null,
            [ExcelArgument(Name = "Document Path", Description = "Server-local path to a document such as a PDF or CSV, inside the WHATSAPP_ATTACHMENT_ROOTS allowlist.")] object? documentPathOptional = null,
            [ExcelArgument(Name = "Caption", Description = "Caption for the image. For a document it is sent as a follow up text.")] object? captionOptional = null,
            [ExcelArgument(Name = "Filename", Description = "Overrides the document name shown on the recipient's device.")] object? filenameOptional = null,
            [ExcelArgument(Name = "Wait For Delivery", Description = "TRUE blocks until the send is attempted and returns a per recipient report. FALSE (default) queues and returns immediately.")] object? waitForDeliveryOptional = null)
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");

            string text = Arg.Str(message);
            string imagePath = Arg.Str(imagePathOptional).Trim();
            string documentPath = Arg.Str(documentPathOptional).Trim();
            string caption = Arg.Str(captionOptional);
            string filename = Arg.Str(filenameOptional).Trim();
            string recipientType = Arg.Str(recipientTypeOptional).Trim();

            if (text.Length == 0 && imagePath.Length == 0 && documentPath.Length == 0)
                return ExcelTable.Error("Provide a Message, an Image Path or a Document Path.");

            if (text.Length > 4096)
                return ExcelTable.Error($"Message must be 4096 characters or fewer, this one is {text.Length}.");

            bool wait = false;
            if (!Arg.IsMissing(waitForDeliveryOptional) && !UtilitySupport.TryReadBool(waitForDeliveryOptional, out wait))
                return ExcelTable.Error("Wait For Delivery must be TRUE or FALSE.");

            var payload = new JObject();
            if (!BuildRecipient(recipientOptional, recipientType, payload, out string? recipientError))
                return ExcelTable.Error(recipientError ?? "Could not read the recipient.");

            if (text.Length > 0) payload["message"] = text;
            if (imagePath.Length > 0) payload["image_path"] = imagePath;
            if (documentPath.Length > 0) payload["document_path"] = documentPath;
            if (caption.Length > 0) payload["caption"] = caption;
            if (filename.Length > 0) payload["filename"] = filename;
            if (wait) payload["wait_for_delivery"] = true;

            string payloadJson = payload.ToString(Formatting.None);

            return AsyncTaskUtil.RunTask(nameof(oa_whatsapp_notify), new object[] { payloadJson }, async () =>
            {
                var response = await OpenAlgoClient.PostAsync("whatsapp/notify", payload);

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                if (response["data"] is JObject report)
                {
                    var rows = new List<object[]>
                    {
                        new object[] { "Result", "Recipient", "Detail" },
                        new object[] { "Status", response["status"]?.ToString() ?? "unknown", response["message"]?.ToString() ?? "" }
                    };

                    foreach (var sent in (report["sent"] as JArray) ?? new JArray())
                        rows.Add(new object[] { "Sent", ExcelTable.Cell(sent), "" });

                    foreach (var failed in (report["failed"] as JArray) ?? new JArray())
                    {
                        if (failed is JObject entry)
                            rows.Add(new object[] { "Failed", ExcelTable.Cell(entry["to"]), ExcelTable.Cell(entry["error"]) });
                        else
                            rows.Add(new object[] { "Failed", ExcelTable.Cell(failed), "" });
                    }

                    rows.Add(new object[] { "Skipped", ExcelTable.Cell(report["skipped"]), "Recipients trimmed by the 5 recipient cap" });

                    return (object)ExcelTable.FromRows(rows);
                }

                var pairs = UtilitySupport.StatusPairs(response);
                if (response["queued"] != null)
                    pairs.Add(new KeyValuePair<string, object>("Queued", ExcelTable.Cell(response["queued"])));
                pairs.Add(new KeyValuePair<string, object>("Delivery",
                    "Queued for delivery, not confirmed. Pass TRUE to Wait For Delivery for a per recipient report."));

                return (object)ExcelTable.KeyValue(pairs);
            })!;
        }

        // ------------------------------------------------------------------
        // Telegram
        // ------------------------------------------------------------------

        /// <summary>
        /// Reads the Telegram bot configuration. The bot token comes back masked.
        /// </summary>
        [ExcelFunction(
            Name = "oa_telegram_config",
            Description = "Telegram bot configuration, with the bot token masked by the server. POST /telegram/webhook is deliberately not wrapped: it is called inbound by Telegram's servers and authenticates with the X-Telegram-Bot-Api-Secret-Token header rather than an OpenAlgo key, so reach it with oa_request if you must. Limit 30 calls per minute.",
            Category = MessagingCategory)]
        public static object oa_telegram_config()
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");

            return AsyncTaskUtil.RunTask(nameof(oa_telegram_config), new object[] { }, async () =>
            {
                var response = await OpenAlgoClient.GetAsync("telegram/config");

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                return (object)UtilitySupport.FlatKeyValue(response["data"], "No Telegram configuration returned");
            })!;
        }

        /// <summary>
        /// Updates the Telegram bot configuration from a key/value range, or from a
        /// single key and value.
        /// </summary>
        [ExcelFunction(
            Name = "oa_telegram_config_set",
            Description = "Updates Telegram bot settings. Accepts a two column key/value range or a single key and value. Keys: token, webhook_url, polling_mode, broadcast_enabled, rate_limit_per_minute (1 to 120). The token is never echoed back into the sheet.",
            Category = MessagingCategory)]
        public static object oa_telegram_config_set(
            [ExcelArgument(Name = "Settings Or Key", Description = "A two column range of setting names and values, or the name of a single setting.")] object settingsOrKey,
            [ExcelArgument(Name = "Value", Description = "Value for a single setting. Omit when passing a range.")] object? valueOptional = null)
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");

            var settings = UtilitySupport.ReadSettings(settingsOrKey, valueOptional, out string? readError);
            if (settings == null)
                return ExcelTable.Error(readError ?? "Could not read the settings.");

            var payload = new JObject();
            var applied = new List<KeyValuePair<string, object>>();

            foreach (var setting in settings)
            {
                if (!ConfigKeys.TryGetValue(setting.Key, out string? kind))
                    return ExcelTable.Error($"'{setting.Key}' is not a Telegram configuration key. Use token, webhook_url, polling_mode, broadcast_enabled or rate_limit_per_minute.");

                string field = setting.Key.Equals("bot_token", StringComparison.OrdinalIgnoreCase) ? "token" : setting.Key.ToLowerInvariant();

                switch (kind)
                {
                    case "bool":
                        if (!UtilitySupport.TryReadBool(setting.Value, out bool flag))
                            return ExcelTable.Error($"'{field}' must be TRUE or FALSE.");
                        payload[field] = flag;
                        applied.Add(new KeyValuePair<string, object>(field, flag));
                        break;
                    case "int":
                        int number = Arg.Int(setting.Value, -1);
                        if (number < 1 || number > 120)
                            return ExcelTable.Error($"'{field}' must be a whole number between 1 and 120.");
                        payload[field] = number;
                        applied.Add(new KeyValuePair<string, object>(field, (double)number));
                        break;
                    default:
                        string content = Arg.Str(setting.Value).Trim();
                        if (content.Length == 0)
                            return ExcelTable.Error($"'{field}' needs a value.");
                        payload[field] = content;
                        applied.Add(new KeyValuePair<string, object>(field,
                            field == "token" ? "set (not shown)" : content));
                        break;
                }
            }

            if (payload.Count == 0)
                return ExcelTable.Error("No settings to update.");

            string identity = payload.ToString(Formatting.None);

            return AsyncTaskUtil.RunTask(nameof(oa_telegram_config_set), new object[] { identity }, async () =>
            {
                var response = await OpenAlgoClient.PostAsync("telegram/config", payload);

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                var pairs = UtilitySupport.StatusPairs(response);
                pairs.AddRange(applied);

                return (object)ExcelTable.KeyValue(pairs, "Setting", "Value");
            })!;
        }

        /// <summary>
        /// Starts the Telegram bot using the stored configuration.
        /// </summary>
        [ExcelFunction(
            Name = "oa_telegram_start",
            Description = "Starts the Telegram bot in polling or webhook mode from the stored configuration. Fails when no bot token is configured. Automatic order alerts only fire while the bot is running. Limit 30 calls per minute.",
            Category = MessagingCategory)]
        public static object oa_telegram_start()
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");

            return AsyncTaskUtil.RunTask(nameof(oa_telegram_start), new object[] { }, async () =>
            {
                var response = await OpenAlgoClient.PostAsync("telegram/start", new JObject());

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                return (object)ExcelTable.KeyValue(UtilitySupport.StatusPairs(response));
            })!;
        }

        /// <summary>
        /// Stops the Telegram bot service.
        /// </summary>
        [ExcelFunction(
            Name = "oa_telegram_stop",
            Description = "Stops the Telegram bot service. Automatic order and Flow alerts are suppressed while it is stopped, although explicit sends through oa_telegram_notify still go out. Limit 30 calls per minute.",
            Category = MessagingCategory)]
        public static object oa_telegram_stop()
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");

            return AsyncTaskUtil.RunTask(nameof(oa_telegram_stop), new object[] { }, async () =>
            {
                var response = await OpenAlgoClient.PostAsync("telegram/stop", new JObject());

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                return (object)ExcelTable.KeyValue(UtilitySupport.StatusPairs(response));
            })!;
        }

        /// <summary>
        /// Lists the OpenAlgo users linked to a Telegram account.
        /// </summary>
        [ExcelFunction(
            Name = "oa_telegram_users",
            Description = "Linked Telegram users, one row per user, optionally filtered by broker or by whether notifications are enabled. Limit 30 calls per minute.",
            Category = MessagingCategory)]
        public static object oa_telegram_users(
            [ExcelArgument(Name = "Broker", Description = "Optional broker filter, for example zerodha.")] object? brokerOptional = null,
            [ExcelArgument(Name = "Notifications Enabled", Description = "Optional TRUE or FALSE filter on the user's notification setting.")] object? notificationsEnabledOptional = null)
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");

            string broker = Arg.Str(brokerOptional).Trim();

            string notifications = "";
            if (!Arg.IsMissing(notificationsEnabledOptional))
            {
                if (!UtilitySupport.TryReadBool(notificationsEnabledOptional, out bool flag))
                    return ExcelTable.Error("Notifications Enabled must be TRUE or FALSE.");
                notifications = flag ? "true" : "false";
            }

            return AsyncTaskUtil.RunTask(nameof(oa_telegram_users), new object[] { broker, notifications }, async () =>
            {
                var query = new List<(string, string?)>();
                if (broker.Length > 0) query.Add(("broker", broker));
                if (notifications.Length > 0) query.Add(("notifications_enabled", notifications));

                var response = await OpenAlgoClient.GetAsync("telegram/users", query.ToArray());

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                if (response["data"] is not JArray users)
                    return (object)ExcelTable.Error("No user list returned");

                return (object)ExcelTable.AutoRows(users, "No linked Telegram users");
            })!;
        }

        /// <summary>
        /// Sends a Telegram message to one linked user.
        /// </summary>
        [ExcelFunction(
            Name = "oa_telegram_notify",
            Description = "Sends a Telegram message to one linked OpenAlgo user. By default the call returns as soon as the message is queued, so success means queued, not delivered; pass TRUE to Wait For Delivery to attempt it immediately. Limit 30 calls per minute.",
            Category = MessagingCategory)]
        public static object oa_telegram_notify(
            [ExcelArgument(Name = "Username", Description = "OpenAlgo username that is already linked to a Telegram ID.")] object username,
            [ExcelArgument(Name = "Message", Description = "Message text, max 4096 characters.")] object message,
            [ExcelArgument(Name = "Wait For Delivery", Description = "TRUE attempts delivery before answering. FALSE (default) queues and returns immediately.")] object? waitForDeliveryOptional = null,
            [ExcelArgument(Name = "Priority", Description = "Queue priority from 1 to 10. Default 5.")] object? priorityOptional = null)
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");

            string user = Arg.Str(username).Trim();
            if (user.Length == 0)
                return ExcelTable.Error("Username is required and must already be linked to Telegram.");

            string text = Arg.Str(message);
            if (text.Trim().Length == 0)
                return ExcelTable.Error("Message is required.");
            if (text.Length > 4096)
                return ExcelTable.Error($"Message must be 4096 characters or fewer, this one is {text.Length}.");

            bool wait = false;
            if (!Arg.IsMissing(waitForDeliveryOptional) && !UtilitySupport.TryReadBool(waitForDeliveryOptional, out wait))
                return ExcelTable.Error("Wait For Delivery must be TRUE or FALSE.");

            int priority = Arg.IsMissing(priorityOptional) ? 5 : Arg.Int(priorityOptional, 5);
            if (priority < 1 || priority > 10)
                return ExcelTable.Error("Priority must be a whole number between 1 and 10.");

            return AsyncTaskUtil.RunTask(nameof(oa_telegram_notify), new object[] { user, text, wait, (double)priority }, async () =>
            {
                var payload = new JObject
                {
                    ["username"] = user,
                    ["message"] = text,
                    ["priority"] = priority,
                    ["wait_for_delivery"] = wait
                };

                var response = await OpenAlgoClient.PostAsync("telegram/notify", payload);

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                string serverMessage = response["message"]?.ToString() ?? "";
                bool delivered = serverMessage.IndexOf("sent successfully", StringComparison.OrdinalIgnoreCase) >= 0;

                var pairs = UtilitySupport.StatusPairs(response);
                pairs.Add(new KeyValuePair<string, object>("Username", user));
                pairs.Add(new KeyValuePair<string, object>("Delivery", delivered
                    ? "Delivered. The bot confirmed the send."
                    : "Queued for delivery. The server accepted the message but has not confirmed it reached Telegram."));

                return (object)ExcelTable.KeyValue(pairs);
            })!;
        }

        /// <summary>
        /// Broadcasts a Telegram message to every linked user matching the filters.
        /// </summary>
        [ExcelFunction(
            Name = "oa_telegram_broadcast",
            Description = "Broadcasts a message to linked Telegram users. KNOWN LIMITATION: the server validates the request but its dispatch is not implemented yet, so it answers with zero delivered and zero failed. That zero is the server's behaviour, not an add-in fault. Broadcast must be enabled in the bot config and is limited to 5 calls per minute.",
            Category = MessagingCategory)]
        public static object oa_telegram_broadcast(
            [ExcelArgument(Name = "Message", Description = "Message text, max 4096 characters.")] object message,
            [ExcelArgument(Name = "Filters", Description = "Optional audience filter: JSON text such as {\"broker\":\"zerodha\"} or a two column key/value range.")] object? filtersOptional = null)
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");

            string text = Arg.Str(message);
            if (text.Trim().Length == 0)
                return ExcelTable.Error("Message is required.");
            if (text.Length > 4096)
                return ExcelTable.Error($"Message must be 4096 characters or fewer, this one is {text.Length}.");

            JObject filters;
            if (Arg.IsMissing(filtersOptional))
            {
                filters = new JObject();
            }
            else
            {
                var parsed = UtilitySupport.ReadFilters(filtersOptional, out string? filterError);
                if (parsed == null)
                    return ExcelTable.Error(filterError ?? "Could not read the filters.");
                filters = parsed;
            }

            string filterJson = filters.ToString(Formatting.None);

            return AsyncTaskUtil.RunTask(nameof(oa_telegram_broadcast), new object[] { text, filterJson }, async () =>
            {
                var payload = new JObject
                {
                    ["message"] = text,
                    ["filters"] = filters
                };

                var response = await OpenAlgoClient.PostAsync("telegram/broadcast", payload);

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                var pairs = UtilitySupport.StatusPairs(response);
                pairs.Add(new KeyValuePair<string, object>("Success Count", ExcelTable.Cell(response["success_count"])));
                pairs.Add(new KeyValuePair<string, object>("Fail Count", ExcelTable.Cell(response["fail_count"])));
                pairs.Add(new KeyValuePair<string, object>("Note",
                    "The server's broadcast dispatch is not implemented yet, so these counts are expected to be zero."));

                return (object)ExcelTable.KeyValue(pairs);
            })!;
        }

        /// <summary>
        /// Command usage statistics for the Telegram bot.
        /// </summary>
        [ExcelFunction(
            Name = "oa_telegram_stats",
            Description = "Telegram bot command statistics: totals, a count per command and the most active users. Covers the last 1 to 365 days, default 7. Limit 30 calls per minute.",
            Category = MessagingCategory)]
        public static object oa_telegram_stats(
            [ExcelArgument(Name = "Days", Description = "Look back window from 1 to 365 days. Default 7.")] object? daysOptional = null)
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");

            int days = Arg.IsMissing(daysOptional) ? 7 : Arg.Int(daysOptional, 7);
            if (days < 1 || days > 365)
                return ExcelTable.Error("Days must be a whole number between 1 and 365.");

            return AsyncTaskUtil.RunTask(nameof(oa_telegram_stats), new object[] { (double)days }, async () =>
            {
                var response = await OpenAlgoClient.GetAsync("telegram/stats", ("days", days.ToString(CultureInfo.InvariantCulture)));

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                return (object)UtilitySupport.FlatKeyValue(response["data"], "No statistics returned");
            })!;
        }

        /// <summary>
        /// Reads the notification preferences of one linked Telegram user.
        /// </summary>
        [ExcelFunction(
            Name = "oa_telegram_preferences",
            Description = "Notification preferences for one linked Telegram user. The Telegram ID comes from the Telegram ID column of oa_telegram_users(). Limit 30 calls per minute.",
            Category = MessagingCategory)]
        public static object oa_telegram_preferences(
            [ExcelArgument(Name = "Telegram ID", Description = "Numeric Telegram user ID, as listed by oa_telegram_users().")] object telegramId)
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");

            if (Arg.IsMissing(telegramId))
                return ExcelTable.Error("Telegram ID is required.");

            long id = (long)Math.Round(Arg.Num(telegramId, 0));
            if (id <= 0)
                return ExcelTable.Error("Telegram ID must be a positive whole number.");

            return AsyncTaskUtil.RunTask(nameof(oa_telegram_preferences), new object[] { (double)id }, async () =>
            {
                var response = await OpenAlgoClient.GetAsync("telegram/preferences", ("telegram_id", id.ToString(CultureInfo.InvariantCulture)));

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                return (object)UtilitySupport.FlatKeyValue(response["data"], "No preferences returned");
            })!;
        }

        /// <summary>
        /// Updates one notification preference for a linked Telegram user.
        /// </summary>
        [ExcelFunction(
            Name = "oa_telegram_preferences_set",
            Description = "Updates one Telegram notification preference. Keys: order_notifications, trade_notifications, pnl_notifications and daily_summary take TRUE or FALSE; summary_time (HH:MM), language and timezone take text. Limit 30 calls per minute.",
            Category = MessagingCategory)]
        public static object oa_telegram_preferences_set(
            [ExcelArgument(Name = "Telegram ID", Description = "Numeric Telegram user ID, as listed by oa_telegram_users().")] object telegramId,
            [ExcelArgument(Name = "Key", Description = "order_notifications, trade_notifications, pnl_notifications, daily_summary, summary_time, language or timezone.")] object key,
            [ExcelArgument(Name = "Value", Description = "TRUE or FALSE for the notification switches, text for summary_time (HH:MM), language and timezone.")] object value)
        {
            if (!OpenAlgoClient.HasApiKey)
                return ExcelTable.Error("OpenAlgo API Key is not set. Use oa_api()");

            if (Arg.IsMissing(telegramId))
                return ExcelTable.Error("Telegram ID is required.");

            long id = (long)Math.Round(Arg.Num(telegramId, 0));
            if (id <= 0)
                return ExcelTable.Error("Telegram ID must be a positive whole number.");

            string name = Arg.Str(key).Trim().ToLowerInvariant();
            if (!PreferenceKeys.TryGetValue(name, out string? kind))
                return ExcelTable.Error($"'{name}' is not a Telegram preference. Use order_notifications, trade_notifications, pnl_notifications, daily_summary, summary_time, language or timezone.");

            var payload = new JObject { ["telegram_id"] = id };

            if (kind == "bool")
            {
                if (!UtilitySupport.TryReadBool(value, out bool flag))
                    return ExcelTable.Error($"'{name}' must be TRUE or FALSE.");
                payload[name] = flag;
            }
            else
            {
                string content = Arg.Str(value).Trim();
                if (content.Length == 0)
                    return ExcelTable.Error($"'{name}' needs a value.");
                payload[name] = content;
            }

            string identity = payload.ToString(Formatting.None);

            return AsyncTaskUtil.RunTask(nameof(oa_telegram_preferences_set), new object[] { identity }, async () =>
            {
                var response = await OpenAlgoClient.PostAsync("telegram/preferences", payload);

                var error = ExcelTable.ErrorOrNull(response);
                if (error != null) return (object)error;

                var pairs = UtilitySupport.StatusPairs(response);
                pairs.Add(new KeyValuePair<string, object>("Telegram ID", (double)id));
                pairs.Add(new KeyValuePair<string, object>(name, ExcelTable.Cell(payload[name])));

                return (object)ExcelTable.KeyValue(pairs);
            })!;
        }

        // ------------------------------------------------------------------
        // Helpers
        // ------------------------------------------------------------------

        /// <summary>
        /// Fills in exactly one of the four WhatsApp recipient forms: self, username,
        /// phone or phones. The API rejects any combination of them, so the choice is
        /// validated here before the call is made.
        /// </summary>
        private static bool BuildRecipient(object? recipient, string recipientType, JObject payload, out string? error)
        {
            error = null;

            var entries = new List<string>();
            foreach (var cell in Arg.Flatten(recipient))
            {
                string cellText = Arg.Str(cell).Trim();
                if (cellText.Length == 0)
                    continue;
                foreach (var part in cellText.Split(',', StringSplitOptions.RemoveEmptyEntries))
                {
                    string entry = part.Trim();
                    if (entry.Length > 0)
                        entries.Add(entry);
                }
            }

            string type = recipientType.Trim().ToLowerInvariant();
            if (type == "auto")
                type = "";

            if (type.Length > 0 && type != "self" && type != "username" && type != "phone" && type != "phones")
            {
                error = "Recipient Type must be self, username, phone or phones.";
                return false;
            }

            bool wantsSelf = type == "self"
                             || (type.Length == 0 && entries.Count == 0)
                             || (type.Length == 0 && entries.Count == 1 && entries[0].Equals("self", StringComparison.OrdinalIgnoreCase));

            if (wantsSelf)
            {
                if (type == "self" && entries.Count > 0 && !entries[0].Equals("self", StringComparison.OrdinalIgnoreCase))
                {
                    error = "Leave Recipient blank when Recipient Type is self. The API accepts exactly one recipient form.";
                    return false;
                }
                payload["self"] = true;
                return true;
            }

            if (entries.Count == 0)
            {
                error = $"Recipient is required when Recipient Type is {type}.";
                return false;
            }

            if (type == "username" || (type.Length == 0 && entries.Count == 1 && !UtilitySupport.IsPhoneNumber(entries[0])))
            {
                if (entries.Count > 1)
                {
                    error = "Only one username can be addressed per call. Use phone numbers for several recipients.";
                    return false;
                }
                payload["username"] = entries[0];
                return true;
            }

            var digits = new List<string>();
            foreach (var entry in entries)
            {
                if (!UtilitySupport.TryReadPhoneNumber(entry, out string number))
                {
                    error = $"'{entry}' is not a phone number in E.164 digits, for example 919876543210. Set Recipient Type to username to address a linked user.";
                    return false;
                }
                digits.Add(number);
            }

            if (type == "phone")
            {
                if (digits.Count > 1)
                {
                    error = "Recipient Type phone takes a single number. Use phones for up to 5 recipients.";
                    return false;
                }
                payload["phone"] = digits[0];
                return true;
            }

            if (digits.Count > 5)
            {
                error = $"WhatsApp notify accepts at most 5 recipients and drops the rest; {digits.Count} were given. Trim the list.";
                return false;
            }

            if (digits.Count == 1 && type.Length == 0)
            {
                payload["phone"] = digits[0];
                return true;
            }

            payload["phones"] = new JArray(digits.Cast<object>().ToArray());
            return true;
        }
    }

    /// <summary>
    /// Shared helpers for the analyzer, calendar, utility and messaging functions.
    /// </summary>
    internal static class UtilitySupport
    {
        /// <summary>
        /// Longest text an Excel cell holds is 32767 characters. Stay under it.
        /// </summary>
        private const int MaxCellLength = 32000;

        private static readonly TimeZoneInfo IndiaTimeZone = ResolveIndiaTimeZone();

        /// <summary>
        /// Indian market times are quoted in IST. Fall back to a fixed offset when the
        /// machine has no matching time zone database entry.
        /// </summary>
        private static TimeZoneInfo ResolveIndiaTimeZone()
        {
            foreach (string id in new[] { "India Standard Time", "Asia/Kolkata" })
            {
                try
                {
                    return TimeZoneInfo.FindSystemTimeZoneById(id);
                }
                catch
                {
                    // Try the next identifier.
                }
            }
            return TimeZoneInfo.CreateCustomTimeZone("OpenAlgo IST", TimeSpan.FromMinutes(330), "IST", "IST");
        }

        /// <summary>
        /// Formats an epoch millisecond token as IST text.
        /// </summary>
        public static string FormatEpoch(JToken? token, string format)
        {
            if (token == null || token.Type == JTokenType.Null)
                return "";

            double milliseconds;
            if (token.Type == JTokenType.Integer || token.Type == JTokenType.Float)
                milliseconds = token.ToObject<double>();
            else if (!double.TryParse(token.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out milliseconds))
                return token.ToString();

            var utc = DateTimeOffset.FromUnixTimeMilliseconds((long)milliseconds).UtcDateTime;
            return TimeZoneInfo.ConvertTimeFromUtc(utc, IndiaTimeZone).ToString(format, CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Reads a JSON array of strings into a list.
        /// </summary>
        public static List<string> StringList(JArray? array)
        {
            var values = new List<string>();
            if (array == null)
                return values;
            foreach (var item in array)
            {
                string text = item.ToString();
                if (!string.IsNullOrWhiteSpace(text))
                    values.Add(text);
            }
            return values;
        }

        /// <summary>
        /// Describes the special sessions a holiday carries, for example "MCX 18:30-23:30".
        /// </summary>
        public static string DescribeSessions(JArray? sessions)
        {
            if (sessions == null || sessions.Count == 0)
                return "";

            var parts = new List<string>();
            foreach (var session in sessions.OfType<JObject>())
            {
                string exchange = session["exchange"]?.ToString() ?? "";
                string start = FormatEpoch(session["start_time"], "HH:mm");
                string end = FormatEpoch(session["end_time"], "HH:mm");
                parts.Add(start.Length > 0 && end.Length > 0
                    ? $"{exchange} {start}-{end}"
                    : exchange);
            }
            return string.Join("; ", parts);
        }

        /// <summary>
        /// Status and message rows shared by the action style functions.
        /// </summary>
        public static List<KeyValuePair<string, object>> StatusPairs(JObject response)
        {
            var pairs = new List<KeyValuePair<string, object>>
            {
                new("Status", response["status"]?.ToString() ?? "unknown")
            };

            string? message = response["message"]?.ToString();
            if (!string.IsNullOrWhiteSpace(message))
                pairs.Add(new KeyValuePair<string, object>("Message", message!));

            return pairs;
        }

        /// <summary>
        /// Appends the properties of a payload that the caller has not already shown.
        /// </summary>
        public static void AppendRemaining(List<KeyValuePair<string, object>> pairs, JObject data, params string[] handled)
        {
            foreach (var property in data.Properties())
            {
                if (handled.Contains(property.Name, StringComparer.OrdinalIgnoreCase))
                    continue;
                pairs.Add(new KeyValuePair<string, object>(ExcelTable.Humanise(property.Name), ExcelTable.Cell(property.Value)));
            }
        }

        /// <summary>
        /// Key/value table for a payload of unknown depth. Nested objects and arrays are
        /// expanded into their own rows rather than being dumped as JSON in one cell.
        /// </summary>
        public static object[,] FlatKeyValue(JToken? data, string emptyMessage)
        {
            if (data == null || data.Type == JTokenType.Null)
                return ExcelTable.Error(emptyMessage);

            if (data is JArray array && array.Count > 0 && array.All(i => i is JObject))
                return ExcelTable.AutoRows(array, emptyMessage);

            var pairs = new List<KeyValuePair<string, object>>();
            FlattenInto(pairs, data, "");

            if (pairs.Count == 0)
                return ExcelTable.Error(emptyMessage);

            return ExcelTable.KeyValue(pairs);
        }

        /// <summary>
        /// Walks a JSON document into key/value rows, naming nested fields by their path.
        /// </summary>
        private static void FlattenInto(List<KeyValuePair<string, object>> pairs, JToken token, string prefix)
        {
            switch (token.Type)
            {
                case JTokenType.Object:
                    foreach (var property in ((JObject)token).Properties())
                        FlattenInto(pairs, property.Value, Join(prefix, ExcelTable.Humanise(property.Name)));
                    break;

                case JTokenType.Array:
                    var array = (JArray)token;
                    if (array.Count == 0)
                    {
                        pairs.Add(new KeyValuePair<string, object>(prefix, ""));
                    }
                    else if (array.All(i => i.Type != JTokenType.Object && i.Type != JTokenType.Array))
                    {
                        pairs.Add(new KeyValuePair<string, object>(prefix, string.Join(", ", array.Select(i => i.ToString()))));
                    }
                    else
                    {
                        for (int i = 0; i < array.Count; i++)
                            FlattenInto(pairs, array[i], Join(prefix, "#" + (i + 1)));
                    }
                    break;

                default:
                    pairs.Add(new KeyValuePair<string, object>(prefix.Length > 0 ? prefix : "Value", ExcelTable.Cell(token)));
                    break;
            }
        }

        private static string Join(string prefix, string name) =>
            prefix.Length == 0 ? name : prefix + " / " + name;

        /// <summary>
        /// Reads a boolean argument strictly: returns false when the value cannot be read
        /// as TRUE or FALSE, rather than guessing a default.
        /// </summary>
        public static bool TryReadBool(object? value, out bool result)
        {
            result = false;
            if (Arg.IsMissing(value))
                return false;

            bool asTrue = Arg.Bool(value, true);
            bool asFalse = Arg.Bool(value, false);
            if (asTrue != asFalse)
                return false;

            result = asTrue;
            return true;
        }

        /// <summary>
        /// True when the text looks like an E.164 phone number.
        /// </summary>
        public static bool IsPhoneNumber(string text) => TryReadPhoneNumber(text, out _);

        /// <summary>
        /// Normalises a phone number to plain E.164 digits.
        /// </summary>
        public static bool TryReadPhoneNumber(string text, out string digits)
        {
            digits = new string((text ?? "").Where(char.IsDigit).ToArray());

            string stripped = (text ?? "").Trim();
            foreach (char c in stripped)
            {
                if (!char.IsDigit(c) && c != '+' && c != ' ' && c != '-' && c != '(' && c != ')')
                    return false;
            }

            return digits.Length >= 8 && digits.Length <= 15;
        }

        /// <summary>
        /// Reads either a two column key/value range, or a single key with its value.
        /// </summary>
        public static List<KeyValuePair<string, object>>? ReadSettings(object source, object? value, out string? error)
        {
            error = null;
            var settings = new List<KeyValuePair<string, object>>();

            if (Arg.IsMissing(source))
            {
                error = "Pass a two column key/value range, or a setting name and a value.";
                return null;
            }

            if (source is object[,] grid)
            {
                int rowCount = grid.GetLength(0);
                int colCount = grid.GetLength(1);

                if (colCount < 2)
                {
                    error = "The settings range needs two columns: the setting name and its value.";
                    return null;
                }

                int firstRow = 0;
                string heading = Arg.Str(grid[0, 0]).Trim().ToLowerInvariant();
                if (rowCount > 1 && (heading == "key" || heading == "setting" || heading == "field" || heading == "name" || heading == "preference"))
                    firstRow = 1;

                for (int row = firstRow; row < rowCount; row++)
                {
                    string name = Arg.Str(grid[row, 0]).Trim();
                    if (name.Length == 0)
                        continue;
                    settings.Add(new KeyValuePair<string, object>(name, grid[row, 1]));
                }
            }
            else
            {
                string name = Arg.Str(source).Trim();
                if (name.Length == 0)
                {
                    error = "The setting name is empty.";
                    return null;
                }
                if (Arg.IsMissing(value))
                {
                    error = $"Pass a value for '{name}', or pass a two column range instead.";
                    return null;
                }
                settings.Add(new KeyValuePair<string, object>(name, value!));
            }

            if (settings.Count == 0)
            {
                error = "No settings were found in the range.";
                return null;
            }

            return settings;
        }

        /// <summary>
        /// Reads a filter argument that arrives either as JSON text or as a two column
        /// key/value range.
        /// </summary>
        public static JObject? ReadFilters(object? source, out string? error)
        {
            error = null;

            if (Arg.IsMissing(source))
                return new JObject();

            if (source is object[,])
            {
                var settings = ReadSettings(source!, null, out error);
                if (settings == null)
                    return null;

                var fromRange = new JObject();
                foreach (var setting in settings)
                    fromRange[setting.Key] = ValueToken(setting.Value);
                return fromRange;
            }

            string text = Arg.Str(source).Trim();
            if (text.Length == 0)
                return new JObject();

            try
            {
                return JObject.Parse(text);
            }
            catch (JsonException ex)
            {
                error = "Filters must be a JSON object such as {\"broker\":\"zerodha\"} or a two column range: " + ex.Message;
                return null;
            }
        }

        /// <summary>
        /// Converts a worksheet value into the JSON token to send. Text that parses as
        /// JSON is sent as JSON; anything else is sent as a string.
        /// </summary>
        public static JToken ValueToken(object? value)
        {
            if (value is double number)
                return new JValue(number);
            if (value is bool flag)
                return new JValue(flag);

            string text = Arg.Str(value).Trim();
            if (text.Length == 0)
                return new JValue("");

            try
            {
                var parsed = JToken.Parse(text);
                switch (parsed.Type)
                {
                    case JTokenType.Object:
                    case JTokenType.Array:
                    case JTokenType.Integer:
                    case JTokenType.Float:
                    case JTokenType.Boolean:
                    case JTokenType.Null:
                        return parsed;
                }
            }
            catch (JsonException)
            {
                // Not JSON, send it as text.
            }

            return new JValue(text);
        }

        /// <summary>
        /// Splits "telegram/stats?days=30" into the route and its query parameters.
        /// </summary>
        public static (string Route, (string, string?)[] Query) SplitQuery(string path)
        {
            int mark = path.IndexOf('?');
            if (mark < 0)
                return (path, Array.Empty<(string, string?)>());

            string route = path.Substring(0, mark);
            var query = new List<(string, string?)>();

            foreach (var part in path.Substring(mark + 1).Split('&', StringSplitOptions.RemoveEmptyEntries))
            {
                int equals = part.IndexOf('=');
                if (equals > 0)
                    query.Add((Uri.UnescapeDataString(part.Substring(0, equals)), Uri.UnescapeDataString(part.Substring(equals + 1))));
                else
                    query.Add((Uri.UnescapeDataString(part), ""));
            }

            return (route, query.ToArray());
        }

        /// <summary>
        /// Keeps long JSON inside the limit of a single Excel cell.
        /// </summary>
        public static string TruncateForCell(string text)
        {
            if (string.IsNullOrEmpty(text) || text.Length <= MaxCellLength)
                return text;
            return text.Substring(0, MaxCellLength) + " ...[truncated, " + text.Length + " characters]";
        }
    }
}
