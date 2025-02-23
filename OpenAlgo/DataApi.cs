using System;
using System.Globalization;
using ExcelDna.Integration;
using ExcelDna.Registration.Utils;
using Newtonsoft.Json.Linq;

namespace OpenAlgo
{
    public static class DataApi
    {
        /// <summary>
        /// Retrieves market quotes for a given symbol.
        /// </summary>
        [ExcelFunction(Name = "oa_quotes", Description = "Retrieves market quotes for a given symbol.")]
        public static object[,] oa_quotes(string symbol, string exchange)
        {
            if (string.IsNullOrWhiteSpace(OpenAlgoConfig.ApiKey))
                return new object[,] { { "Error: OpenAlgo API Key is not set. Use oa_api()" } };

            string endpoint = $"{OpenAlgoConfig.HostUrl}/api/{OpenAlgoConfig.Version}/quotes";
            var payload = new JObject { ["apikey"] = OpenAlgoConfig.ApiKey, ["symbol"] = symbol, ["exchange"] = exchange };

            return (object[,])AsyncTaskUtil.RunTask(nameof(oa_quotes), new object[] { symbol, exchange }, async () =>
            {
                string json = await Utilities.PostRequestAsync(endpoint, payload);
                JObject jsonResponse = JObject.Parse(json);

                if (!jsonResponse.TryGetValue("data", out JToken? dataToken) || dataToken == null)
                    return new object[,] { { "Error: No quote data found" } };

                JObject dataObject = (JObject)dataToken;
                object[,] resultArray = new object[dataObject.Count + 1, 2];

                resultArray[0, 0] = $"{symbol} ({exchange})";
                resultArray[0, 1] = "Value";

                int rowIndex = 1;
                foreach (var property in dataObject.Properties())
                {
                    resultArray[rowIndex, 0] = property.Name ?? "Unknown Key";
                    resultArray[rowIndex, 1] = property.Value?.ToString() ?? "N/A";
                    rowIndex++;
                }

                return resultArray;
            })!;
        }

        /// <summary>
        /// Retrieves order book depth for a given symbol.
        /// </summary>
        [ExcelFunction(Name = "oa_depth", Description = "Retrieves order book depth for a given symbol.")]
        public static object[,] oa_depth(string symbol, string exchange)
        {
            if (string.IsNullOrWhiteSpace(OpenAlgoConfig.ApiKey))
                return new object[,] { { "Error: OpenAlgo API Key is not set. Use oa_api()" } };

            string endpoint = $"{OpenAlgoConfig.HostUrl}/api/{OpenAlgoConfig.Version}/depth";
            var payload = new JObject { ["apikey"] = OpenAlgoConfig.ApiKey, ["symbol"] = symbol, ["exchange"] = exchange };

            return (object[,])AsyncTaskUtil.RunTask(nameof(oa_depth), new object[] { symbol, exchange }, async () =>
            {
                string json = await Utilities.PostRequestAsync(endpoint, payload);
                JObject jsonResponse = JObject.Parse(json);

                if (!jsonResponse.TryGetValue("data", out JToken? dataToken) || dataToken == null)
                    return new object[,] { { "Error: No depth data found" } };

                JObject dataObject = (JObject)dataToken;
                double ltp = dataObject["ltp"]?.ToObject<double?>() ?? 0;
                int volume = dataObject["volume"]?.ToObject<int?>() ?? 0;

                JArray asks = dataObject["asks"] as JArray ?? new JArray();
                JArray bids = dataObject["bids"] as JArray ?? new JArray();
                int rowCount = Math.Max(asks.Count, bids.Count);

                object[,] resultArray = new object[rowCount + 2, 4];

                resultArray[0, 0] = "LTP";
                resultArray[0, 1] = ltp;
                resultArray[0, 2] = "Volume";
                resultArray[0, 3] = volume;
            

                resultArray[1, 0] = "Ask Price";
                resultArray[1, 1] = "Ask Quantity";
                resultArray[1, 2] = "Bid Price";
                resultArray[1, 3] = "Bid Quantity";

                for (int i = 0; i < rowCount; i++)
                {
                    resultArray[i + 2, 0] = asks.ElementAtOrDefault(i)?["price"]?.ToObject<double?>() ?? 0;
                    resultArray[i + 2, 1] = asks.ElementAtOrDefault(i)?["quantity"]?.ToObject<int?>() ?? 0;
                    resultArray[i + 2, 2] = bids.ElementAtOrDefault(i)?["price"]?.ToObject<double?>() ?? 0;
                    resultArray[i + 2, 3] = bids.ElementAtOrDefault(i)?["quantity"]?.ToObject<int?>() ?? 0;
                }

                return resultArray;
            })!;
        }

        /// <summary>
        /// Retrieves historical market data with timestamps converted to IST.
        /// </summary>
        [ExcelFunction(Name = "oa_history", Description = "Retrieves historical market data.")]
        public static object[,] oa_history(string symbol, string exchange, string interval, string startDate, string endDate)
        {
            if (string.IsNullOrWhiteSpace(OpenAlgoConfig.ApiKey))
                return new object[,] { { "Error: OpenAlgo API Key is not set. Use oa_api()" } };

            string endpoint = $"{OpenAlgoConfig.HostUrl}/api/{OpenAlgoConfig.Version}/history";
            var payload = new JObject
            {
                ["apikey"] = OpenAlgoConfig.ApiKey,
                ["symbol"] = symbol,
                ["exchange"] = exchange,
                ["interval"] = interval,
                ["start_date"] = startDate,
                ["end_date"] = endDate
            };

            return (object[,])AsyncTaskUtil.RunTask(nameof(oa_history), new object[] { symbol, exchange, interval, startDate, endDate }, async () =>
            {
                string json = await Utilities.PostRequestAsync(endpoint, payload);
                JObject jsonResponse = JObject.Parse(json);

                if (!jsonResponse.TryGetValue("data", out JToken? dataToken) || dataToken == null)
                    return new object[,] { { "Error: No historical data found" } };

                JArray historyArray = (JArray)dataToken;
                object[,] resultArray = new object[historyArray.Count + 1, 8];
                string[] headers = { "Ticker", "Date", "Time (IST)", "Open", "High", "Low", "Close", "Volume" };

                for (int col = 0; col < headers.Length; col++)
                    resultArray[0, col] = headers[col];

                for (int i = 0; i < historyArray.Count; i++)
                {
                    JObject data = (JObject)historyArray[i];
                    long unixTimestamp = data["timestamp"]?.ToObject<long>() ?? 0;
                    DateTime istDateTime = DateTimeOffset.FromUnixTimeSeconds(unixTimestamp)
                        .UtcDateTime.AddHours(5).AddMinutes(30); // Convert to IST

                    resultArray[i + 1, 0] = symbol;
                    resultArray[i + 1, 1] = istDateTime.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                    resultArray[i + 1, 2] = istDateTime.ToString("HH:mm:ss", CultureInfo.InvariantCulture);
                    resultArray[i + 1, 3] = data["open"]?.ToObject<double?>() ?? 0;
                    resultArray[i + 1, 4] = data["high"]?.ToObject<double?>() ?? 0;
                    resultArray[i + 1, 5] = data["low"]?.ToObject<double?>() ?? 0;
                    resultArray[i + 1, 6] = data["close"]?.ToObject<double?>() ?? 0;
                    resultArray[i + 1, 7] = data["volume"]?.ToObject<int?>() ?? 0;
                }

                return resultArray;
            })!;
        }
        [ExcelFunction(Name = "oa_intervals", Description = "Retrieves available time intervals.")]
        public static object[,] oa_intervals()
        {
            if (string.IsNullOrWhiteSpace(OpenAlgoConfig.ApiKey))
                return new object[,] { { "Error: OpenAlgo API Key is not set. Use oa_api()" } };

            string endpoint = $"{OpenAlgoConfig.HostUrl}/api/{OpenAlgoConfig.Version}/intervals";
            var payload = new JObject { ["apikey"] = OpenAlgoConfig.ApiKey };

            return (object[,])AsyncTaskUtil.RunTask(nameof(oa_intervals), new object[] { }, async () =>
            {
                string json = await Utilities.PostRequestAsync(endpoint, payload);
                JObject jsonResponse = JObject.Parse(json);

                if (!jsonResponse.TryGetValue("data", out JToken? dataToken) || dataToken == null)
                    return new object[,] { { "Error: No interval data found" } };

                JObject dataObject = (JObject)dataToken;
                var categories = new[] { "Seconds", "Minutes", "Hours", "Days", "Weeks", "Months" };

                // Calculate total number of rows required
                int totalRows = categories.Sum(cat => (dataObject[cat.ToLower()] as JArray)?.Count ?? 0) + 1;
                object[,] resultArray = new object[totalRows, 2];

                // Headers
                resultArray[0, 0] = "Category";
                resultArray[0, 1] = "Interval";

                int rowIndex = 1;
                foreach (var category in categories)
                {
                    JArray intervalsArray = dataObject[category.ToLower()] as JArray ?? new JArray();
                    foreach (var interval in intervalsArray)
                    {
                        resultArray[rowIndex, 0] = category;
                        resultArray[rowIndex, 1] = interval.ToString();
                        rowIndex++;
                    }
                }

                return resultArray;
            })!;
        }

    }
}
