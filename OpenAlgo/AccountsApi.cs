using ExcelDna.Integration;
using ExcelDna.Registration.Utils;
using Newtonsoft.Json.Linq;

namespace OpenAlgo
{
    public static class AccountsApi
    {
        [ExcelFunction(Description = "Retrieve funds from OpenAlgo API.")]
        public static object[,] oa_funds()
        {
            if (string.IsNullOrWhiteSpace(OpenAlgoConfig.ApiKey))
                return new object[,] { { "Error: OpenAlgo API Key is not set. Use oa_api()" } };

            string endpoint = $"{OpenAlgoConfig.HostUrl}/api/{OpenAlgoConfig.Version}/funds";
            var payload = new JObject { ["apikey"] = OpenAlgoConfig.ApiKey };

            return (object[,])AsyncTaskUtil.RunTask(nameof(oa_funds), new object[] { }, async () =>
            {
                string json = await Utilities.PostRequestAsync(endpoint, payload);
                JObject jsonResponse = JObject.Parse(json);

                if (!jsonResponse.TryGetValue("data", out JToken? dataToken) || dataToken == null || dataToken.Type == JTokenType.Null)
                    return new object[,] { { "Error: No data found" } };

                JObject dataObject = (JObject)dataToken;
                object[,] resultArray = new object[dataObject.Count, 2];

                int rowIndex = 0;
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
        /// Retrieves the OrderBook from OpenAlgo API.
        /// </summary>
        [ExcelFunction(Name = "oa_orderbook", Description = "Retrieve the order book from OpenAlgo API.")]
        public static object[,] oa_orderbook()
        {
            if (string.IsNullOrWhiteSpace(OpenAlgoConfig.ApiKey))
                return new object[,] { { "Error: API Key is not set. Use SetOpenAlgoConfig()" } };

            string endpoint = $"{OpenAlgoConfig.HostUrl}/api/{OpenAlgoConfig.Version}/orderbook";
            var payload = new JObject { ["apikey"] = OpenAlgoConfig.ApiKey };

            return (object[,])AsyncTaskUtil.RunTask(nameof(oa_orderbook), new object[] { }, async () =>
            {
                try
                {
                    string json = await Utilities.PostRequestAsync(endpoint, payload);
                    JObject jsonResponse = JObject.Parse(json);

                    // ✅ Check if response contains "data"
                    if (!jsonResponse.TryGetValue("data", out JToken? dataToken) || dataToken == null || dataToken.Type == JTokenType.Null)
                        return new object[,] { { "Error: No data found" } };

                    JObject dataObject = (JObject)dataToken;

                    // ✅ Extract Orders
                    if (!dataObject.TryGetValue("orders", out JToken? ordersToken) || ordersToken == null || ordersToken.Type != JTokenType.Array)
                        return new object[,] { { "Error: No orders found" } };

                    JArray ordersArray = (JArray)ordersToken;
                    int orderCount = ordersArray.Count;
                    int columnCount = 11; // Number of columns (Action, Exchange, etc.)

                    // ✅ If no orders exist, return an error message
                    if (orderCount == 0)
                        return new object[,] { { "Error: No orders available" } };

                    // ✅ Prepare Excel table with headers
                    object[,] resultArray = new object[orderCount + 1, columnCount]; // Only order rows + header row
                    string[] headers = { "Symbol", "Action", "Exchange", "Quantity", "Order Status", "Order ID", "Price", "Price Type", "Trigger Price", "Product", "Timestamp" };

                    // 🔹 Add headers to the first row
                    for (int col = 0; col < headers.Length; col++)
                    {
                        resultArray[0, col] = headers[col];
                    }

                    // 🔹 Add order data
                    for (int i = 0; i < orderCount; i++)
                    {
                        JObject order = (JObject)ordersArray[i];

                        resultArray[i + 1, 0] = order["symbol"]?.ToString() ?? "";
                        resultArray[i + 1, 1] = order["action"]?.ToString() ?? "";
                        resultArray[i + 1, 2] = order["exchange"]?.ToString() ?? "";
                        resultArray[i + 1, 3] = order["quantity"]?.ToString() ?? "";
                        resultArray[i + 1, 4] = order["order_status"]?.ToString() ?? "";
                        resultArray[i + 1, 5] = order["orderid"]?.ToString() ?? "";
                        resultArray[i + 1, 6] = order["price"]?.ToObject<double?>() ?? 0;
                        resultArray[i + 1, 7] = order["pricetype"]?.ToString() ?? "";
                        resultArray[i + 1, 8] = order["trigger_price"]?.ToObject<double?>() ?? 0;
                        resultArray[i + 1, 9] = order["product"]?.ToString() ?? "";
                        resultArray[i + 1, 10] = order["timestamp"]?.ToString() ?? "";
                        
                    }

                    return resultArray;
                }
                catch (IndexOutOfRangeException ex)
                {
                    return new object[,] { { "Error:", "Index out of range - " + ex.Message } };
                }
                catch (Exception ex)
                {
                    return new object[,] { { "Error:", ex.Message } };
                }
            })!;
        }

        [ExcelFunction(Name = "oa_tradebook", Description = "Retrieves the trade book from OpenAlgo API.")]
        public static object[,] oa_tradebook()
        {
            if (string.IsNullOrWhiteSpace(OpenAlgoConfig.ApiKey))
                return new object[,] { { "Error: API Key is not set. Use SetOpenAlgoConfig()" } };

            string endpoint = $"{OpenAlgoConfig.HostUrl}/api/{OpenAlgoConfig.Version}/tradebook";
            var payload = new JObject { ["apikey"] = OpenAlgoConfig.ApiKey };

            return (object[,])AsyncTaskUtil.RunTask(nameof(oa_tradebook), new object[] { }, async () =>
            {
                try
                {
                    string json = await Utilities.PostRequestAsync(endpoint, payload);
                    JObject jsonResponse = JObject.Parse(json);

                    if (!jsonResponse.TryGetValue("data", out JToken? dataToken) || dataToken == null || dataToken.Type != JTokenType.Array)
                        return new object[,] { { "Error: No trade data available" } };

                    JArray tradesArray = (JArray)dataToken;
                    int tradeCount = tradesArray.Count;
                    int columnCount = 9; // Number of trade attributes

                    // ✅ If no trades exist, return only the headers
                    object[,] resultArray = new object[Math.Max(tradeCount + 1, 1), columnCount];
                    string[] headers = { "Symbol", "Exchange", "Action", "Quantity", "Product", "Timestamp", "Trade Value", "Average Price", "Order ID" };

                    for (int col = 0; col < headers.Length; col++)
                        resultArray[0, col] = headers[col];

                    if (tradeCount == 0)
                        return resultArray; // Return only headers if no trades

                    for (int i = 0; i < tradeCount; i++)
                    {
                        JObject trade = (JObject)tradesArray[i];

                        resultArray[i + 1, 0] = trade["symbol"]?.ToString() ?? "";
                        resultArray[i + 1, 1] = trade["exchange"]?.ToString() ?? "";
                        resultArray[i + 1, 2] = trade["action"]?.ToString() ?? "";
                        resultArray[i + 1, 3] = trade["quantity"]?.ToObject<int?>() ?? 0;
                        resultArray[i + 1, 4] = trade["product"]?.ToString() ?? "";
                        resultArray[i + 1, 5] = trade["timestamp"]?.ToString() ?? "";
                        resultArray[i + 1, 6] = trade["trade_value"]?.ToObject<double?>() ?? 0;
                        resultArray[i + 1, 7] = trade["average_price"]?.ToObject<double?>() ?? 0;
                        resultArray[i + 1, 8] = trade["orderid"]?.ToString() ?? "";
 
                        
                    }

                    return resultArray;
                }
                catch (Exception ex)
                {
                    return new object[,] { { "Error:", ex.Message } };
                }
            })!;
        }


        [ExcelFunction(Name = "oa_positionbook", Description = "Retrieves the position book from OpenAlgo API.")]
        public static object[,] oa_positionbook()
        {
            if (string.IsNullOrWhiteSpace(OpenAlgoConfig.ApiKey))
                return new object[,] { { "Error: API Key is not set. Use SetOpenAlgoConfig()" } };

            string endpoint = $"{OpenAlgoConfig.HostUrl}/api/{OpenAlgoConfig.Version}/positionbook";
            var payload = new JObject { ["apikey"] = OpenAlgoConfig.ApiKey };

            return (object[,])AsyncTaskUtil.RunTask(nameof(oa_positionbook), new object[] { }, async () =>
            {
                try
                {
                    string json = await Utilities.PostRequestAsync(endpoint, payload);
                    JObject jsonResponse = JObject.Parse(json);

                    if (!jsonResponse.TryGetValue("data", out JToken? dataToken) || dataToken == null || dataToken.Type != JTokenType.Array)
                        return new object[,] { { "Error: No position data available" } };

                    JArray positionsArray = (JArray)dataToken;
                    int positionCount = positionsArray.Count;
                    int columnCount = 5; // Number of position attributes

                    // ✅ If no positions exist, return only the headers
                    object[,] resultArray = new object[Math.Max(positionCount + 1, 1), columnCount];
                    string[] headers = { "Symbol", "Exchange", "Quantity", "Product", "Average Price" };

                    for (int col = 0; col < headers.Length; col++)
                        resultArray[0, col] = headers[col];

                    if (positionCount == 0)
                        return resultArray; // Return only headers if no positions

                    for (int i = 0; i < positionCount; i++)
                    {
                        JObject position = (JObject)positionsArray[i];

                        resultArray[i + 1, 0] = position["symbol"]?.ToString() ?? "";
                        resultArray[i + 1, 1] = position["exchange"]?.ToString() ?? "";
                        resultArray[i + 1, 2] = position["quantity"]?.ToObject<int?>() ?? 0;
                        resultArray[i + 1, 3] = position["product"]?.ToString() ?? "";
                        resultArray[i + 1, 4] = position["average_price"]?.ToObject<double?>() ?? 0;

                        
                    }

                    return resultArray;
                }
                catch (Exception ex)
                {
                    return new object[,] { { "Error:", ex.Message } };
                }
            })!;
        }


        [ExcelFunction(Name = "oa_holdings", Description = "Retrieves holdings from OpenAlgo API.")]
        public static object[,] oa_holdings()
        {
            if (string.IsNullOrWhiteSpace(OpenAlgoConfig.ApiKey))
                return new object[,] { { "Error: API Key is not set. Use SetOpenAlgoConfig()" } };

            string endpoint = $"{OpenAlgoConfig.HostUrl}/api/{OpenAlgoConfig.Version}/holdings";
            var payload = new JObject { ["apikey"] = OpenAlgoConfig.ApiKey };

            return (object[,])AsyncTaskUtil.RunTask(nameof(oa_holdings), new object[] { }, async () =>
            {
                try
                {
                    string json = await Utilities.PostRequestAsync(endpoint, payload);
                    JObject jsonResponse = JObject.Parse(json);

                    // ✅ Check if response contains "data" and "holdings"
                    if (!jsonResponse.TryGetValue("data", out JToken? dataToken) || dataToken == null || dataToken.Type != JTokenType.Object)
                        return new object[,] { { "Error: No holdings data available" } };

                    JObject dataObject = (JObject)dataToken;
                    if (!dataObject.TryGetValue("holdings", out JToken? holdingsToken) || holdingsToken == null || holdingsToken.Type != JTokenType.Array)
                        return new object[,] { { "Error: No holdings found" } };

                    JArray holdingsArray = (JArray)holdingsToken;
                    int holdingsCount = holdingsArray.Count;
                    int columnCount = 6; // Number of attributes

                    // ✅ If no holdings exist, return only headers
                    object[,] resultArray = new object[Math.Max(holdingsCount + 1, 1), columnCount];
                    string[] headers = { "Symbol", "Exchange", "Quantity", "Product", "Pnl", "Pnl Percent" };

                    for (int col = 0; col < headers.Length; col++)
                        resultArray[0, col] = headers[col];

                    if (holdingsCount == 0)
                        return resultArray; // Return only headers if no holdings

                    // 🔹 Add holdings data
                    for (int i = 0; i < holdingsCount; i++)
                    {
                        JObject holding = (JObject)holdingsArray[i];

                        resultArray[i + 1, 0] = holding["symbol"]?.ToString() ?? "";
                        resultArray[i + 1, 1] = holding["exchange"]?.ToString() ?? "";
                        resultArray[i + 1, 2] = holding["quantity"]?.ToObject<int?>() ?? 0;
                        resultArray[i + 1, 3] = holding["product"]?.ToString() ?? "";
                        resultArray[i + 1, 4] = holding["pnl"]?.ToObject<double?>() ?? 0;
                        resultArray[i + 1, 5] = holding["pnlpercent"]?.ToObject<double?>() ?? 0;
                        
                        
                        
                    }

                    return resultArray;
                }
                catch (Exception ex)
                {
                    return new object[,] { { "Error:", ex.Message } };
                }
            })!;
        }


    }
}
