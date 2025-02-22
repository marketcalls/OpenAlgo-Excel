using ExcelDna.Integration;
using ExcelDna.Registration.Utils;
using Newtonsoft.Json.Linq;

namespace OpenAlgo
{
    public static class AccountsApi
    {
        [ExcelFunction(Description = "Retrieve funds from OpenAlgo API.")]
        public static object[,] Funds()
        {
            if (string.IsNullOrWhiteSpace(OpenAlgoConfig.ApiKey))
                return new object[,] { { "Error: API Key is not set. Use SetOpenAlgoConfig()" } };

            string endpoint = $"{OpenAlgoConfig.HostUrl}/api/{OpenAlgoConfig.Version}/funds";
            var payload = new JObject { ["apikey"] = OpenAlgoConfig.ApiKey };

            return (object[,])AsyncTaskUtil.RunTask(nameof(Funds), new object[] { }, async () =>
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
        [ExcelFunction(Description = "Retrieve the order book from OpenAlgo API.")]
        public static object[,] OrderBook()
        {
            if (string.IsNullOrWhiteSpace(OpenAlgoConfig.ApiKey))
                return new object[,] { { "Error: API Key is not set. Use SetOpenAlgoConfig()" } };

            string endpoint = $"{OpenAlgoConfig.HostUrl}/api/{OpenAlgoConfig.Version}/orderbook";
            var payload = new JObject { ["apikey"] = OpenAlgoConfig.ApiKey };

            return (object[,])AsyncTaskUtil.RunTask(nameof(OrderBook), new object[] { }, async () =>
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
                    string[] headers = { "Action", "Exchange", "Order Status", "Order ID", "Price", "Price Type", "Product", "Quantity", "Symbol", "Timestamp", "Trigger Price" };

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



    }
}
