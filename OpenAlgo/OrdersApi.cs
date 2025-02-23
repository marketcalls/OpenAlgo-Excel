using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelDna.Registration.Utils;
using Newtonsoft.Json.Linq;

namespace OpenAlgo
{
    public static class OrderApi
    {
        /// <summary>
        /// Places an order via OpenAlgo API.
        /// </summary>
        [ExcelFunction(
            Name = "oa_placeorder",
            Description = "Places an order through OpenAlgo API.")]
        public static object[,] oa_placeorder(
            [ExcelArgument(Name = "Strategy", Description = "Trading strategy name")] string strategy,
            [ExcelArgument(Name = "Symbol", Description = "Trading symbol")] string symbol,
            [ExcelArgument(Name = "Action", Description = "Order action (BUY/SELL)")] string action,
            [ExcelArgument(Name = "Exchange", Description = "Exchange code(NSE/BSE/NFO/MCX)")] string exchange,
            [ExcelArgument(Name = "PriceType", Description = "Price type (MARKET/LIMIT)")] string priceType,
            [ExcelArgument(Name = "Product", Description = "Product type (MIS/CNC/NRML)")] string product,
            [ExcelArgument(Name = "Quantity", Description = "Order quantity")] object? quantity = null,
            [ExcelArgument(Name = "Price", Description = "Order price (optional)")] object? price = null,
            [ExcelArgument(Name = "TriggerPrice", Description = "Trigger price (optional)")] object? triggerPrice = null,
            [ExcelArgument(Name = "DisclosedQuantity", Description = "Disclosed quantity (optional)")] object? disclosedQuantity = null)
        {
            if (string.IsNullOrWhiteSpace(OpenAlgoConfig.ApiKey))
                return new object[,] { { "Error: OpenAlgo API Key is not set. Use oa_api()" } };

            // ✅ Convert all values to strings
            string quantityStr = (quantity is ExcelMissing or null) ? "0" : quantity.ToString()!;
            string priceStr = (price is ExcelMissing or null) ? "0" : price.ToString()!;
            string triggerPriceStr = (triggerPrice is ExcelMissing or null) ? "0" : triggerPrice.ToString()!;
            string disclosedQuantityStr = (disclosedQuantity is ExcelMissing or null) ? "0" : disclosedQuantity.ToString()!;

            string endpoint = $"{OpenAlgoConfig.HostUrl}/api/{OpenAlgoConfig.Version}/placeorder";

            var payload = new JObject
            {
                ["apikey"] = OpenAlgoConfig.ApiKey,
                ["strategy"] = strategy,
                ["symbol"] = symbol,
                ["action"] = action,
                ["exchange"] = exchange,
                ["pricetype"] = priceType,
                ["product"] = product,
                ["quantity"] = quantityStr,
                ["price"] = priceStr,
                ["trigger_price"] = triggerPriceStr,
                ["disclosed_quantity"] = disclosedQuantityStr
            };

            return (object[,])AsyncTaskUtil.RunTask(nameof(oa_placeorder), new object[] { }, async () =>
            {
                try
                {
                    using (var client = new HttpClient())
                    {
                        var content = new StringContent(payload.ToString(), Encoding.UTF8, "application/json");
                        var response = await client.PostAsync(endpoint, content);
                        string responseBody = await response.Content.ReadAsStringAsync();

                        if (response.IsSuccessStatusCode)
                        {
                            var jsonResponse = JObject.Parse(responseBody);
                            string orderId = jsonResponse["orderid"]?.ToString() ?? "Unknown";

                            // ✅ Return Order ID in separate columns
                            return new object[,]
                            {
                                { "Order ID", orderId }
                            };
                        }
                        else
                        {
                            return new object[,] { { $"Error: {response.StatusCode} - {responseBody}" } };
                        }
                    }
                }
                catch (Exception ex)
                {
                    return new object[,] { { $"Exception: {ex.Message}" } };
                }
            })!;
        }

        /// <summary>
        /// Places a smart order via OpenAlgo API.
        /// </summary>
        [ExcelFunction(
            Name = "oa_placesmartorder",
            Description = "Places a smart order through OpenAlgo API.")]
        public static object[,] oa_placesmartorder(
            [ExcelArgument(Name = "Strategy", Description = "Trading strategy name")] string strategy,
            [ExcelArgument(Name = "Symbol", Description = "Trading symbol")] string symbol,
            [ExcelArgument(Name = "Action", Description = "Order action (BUY/SELL)")] string action,
            [ExcelArgument(Name = "Exchange", Description = "Exchange code(NSE/BSE/NFO/MCX)")] string exchange,
            [ExcelArgument(Name = "PriceType", Description = "Price type (MARKET/LIMIT)")] string priceType,
            [ExcelArgument(Name = "Product", Description = "Product type (MIS/CNC/NRML)")] string product,
            [ExcelArgument(Name = "Quantity", Description = "Order quantity")] object? quantity = null,
            [ExcelArgument(Name = "PositionSize", Description = "Desired position size")] object? positionSize = null,
            [ExcelArgument(Name = "Price", Description = "Order price (optional)")] object? price = null,
            [ExcelArgument(Name = "TriggerPrice", Description = "Trigger price (optional)")] object? triggerPrice = null,
            [ExcelArgument(Name = "DisclosedQuantity", Description = "Disclosed quantity (optional)")] object? disclosedQuantity = null)
        {
            if (string.IsNullOrWhiteSpace(OpenAlgoConfig.ApiKey))
                return new object[,] { { "Error: API Key is not set. Use SetOpenAlgoConfig()" } };

            // Convert all values to strings
            string quantityStr = (quantity is ExcelMissing or null) ? "0" : quantity.ToString()!;
            string positionSizeStr = (positionSize is ExcelMissing or null) ? "0" : positionSize.ToString()!;
            string priceStr = (price is ExcelMissing or null) ? "0" : price.ToString()!;
            string triggerPriceStr = (triggerPrice is ExcelMissing or null) ? "0" : triggerPrice.ToString()!;
            string disclosedQuantityStr = (disclosedQuantity is ExcelMissing or null) ? "0" : disclosedQuantity.ToString()!;

            string endpoint = $"{OpenAlgoConfig.HostUrl}/api/{OpenAlgoConfig.Version}/placesmartorder";

            var payload = new JObject
            {
                ["apikey"] = OpenAlgoConfig.ApiKey,
                ["strategy"] = strategy,
                ["symbol"] = symbol,
                ["action"] = action,
                ["exchange"] = exchange,
                ["pricetype"] = priceType,
                ["product"] = product,
                ["quantity"] = quantityStr,
                ["position_size"] = positionSizeStr,
                ["price"] = priceStr,
                ["trigger_price"] = triggerPriceStr,
                ["disclosed_quantity"] = disclosedQuantityStr
            };

            return (object[,])AsyncTaskUtil.RunTask(nameof(oa_placesmartorder), new object[] { }, async () =>
            {
                try
                {
                    using (var client = new HttpClient())
                    {
                        var content = new StringContent(payload.ToString(), Encoding.UTF8, "application/json");
                        var response = await client.PostAsync(endpoint, content);
                        string responseBody = await response.Content.ReadAsStringAsync();

                        if (response.IsSuccessStatusCode)
                        {
                            var jsonResponse = JObject.Parse(responseBody);
                            string orderId = jsonResponse["orderid"]?.ToString() ?? "Unknown";

                            // Return Order ID in separate columns
                            return new object[,]
                            {
                                { "Order ID", orderId }
                            };
                        }
                        else
                        {
                            return new object[,] { { $"Error: {response.StatusCode} - {responseBody}" } };
                        }
                    }
                }
                catch (Exception ex)
                {
                    return new object[,] { { $"Exception: {ex.Message}" } };
                }
            })!;
        }

        /// <summary>
        /// Places a basket of orders via OpenAlgo API.
        /// </summary>
        [ExcelFunction(
            Name = "oa_basketorder",
            Description = "Places a basket of orders through OpenAlgo API.")]
        public static object[,] oa_basketorder(
            [ExcelArgument(Name = "Strategy", Description = "Trading strategy name")] string strategy,
            [ExcelArgument(Name = "Orders", Description = "Array of order parameters")] object[,] orders)
        {
            if (string.IsNullOrWhiteSpace(OpenAlgoConfig.ApiKey))
                return new object[,] { { "Error: API Key is not set. Use SetOpenAlgoConfig()" } };

            string endpoint = $"{OpenAlgoConfig.HostUrl}/api/{OpenAlgoConfig.Version}/basketorder";

            var orderList = new JArray();

            for (int i = 0; i < orders.GetLength(0); i++)
            {
                var order = new JObject
                {
                    ["symbol"] = orders[i, 0]?.ToString() ?? string.Empty,
                    ["exchange"] = orders[i, 1]?.ToString() ?? string.Empty,
                    ["action"] = orders[i, 2]?.ToString() ?? string.Empty,
                    ["quantity"] = orders[i, 3]?.ToString() ?? "0",
                    ["pricetype"] = orders[i, 4]?.ToString() ?? "MARKET",
                    ["product"] = orders[i, 5]?.ToString() ?? "MIS",
                    ["price"] = orders[i, 6]?.ToString() ?? "0",
                    ["trigger_price"] = orders[i, 7]?.ToString() ?? "0",
                    ["disclosed_quantity"] = orders[i, 8]?.ToString() ?? "0"
                };
                orderList.Add(order);
            }

            var payload = new JObject
            {
                ["apikey"] = OpenAlgoConfig.ApiKey,
                ["strategy"] = strategy,
                ["orders"] = orderList
            };

            return (object[,])AsyncTaskUtil.RunTask(nameof(oa_basketorder), new object[] { }, async () =>
            {
                try
                {
                    using (var client = new HttpClient())
                    {
                        var content = new StringContent(payload.ToString(), Encoding.UTF8, "application/json");
                        var response = await client.PostAsync(endpoint, content);
                        string responseBody = await response.Content.ReadAsStringAsync();

                        if (response.IsSuccessStatusCode)
                        {
                            var jsonResponse = JObject.Parse(responseBody);
                            var results = jsonResponse["results"] as JArray;

                            if (results != null)
                            {
                                var output = new object[results.Count + 1, 3];
                                output[0, 0] = "Symbol";
                                output[0, 1] = "Order ID";
                                output[0, 2] = "Status";

                                for (int i = 0; i < results.Count; i++)
                                {
                                    var result = results[i];
                                    output[i + 1, 0] = result["symbol"]?.ToString() ?? string.Empty;
                                    output[i + 1, 1] = result["orderid"]?.ToString() ?? string.Empty;
                                    output[i + 1, 2] = result["status"]?.ToString() ?? string.Empty;
                                }

                                return output;
                            }
                            else
                            {
                                return new object[,] { { "Error: Invalid response format." } };
                            }
                        }
                        else
                        {
                            return new object[,] { { $"Error: {response.StatusCode} - {responseBody}" } };
                        }
                    }
                }
                catch (Exception ex)
                {
                    return new object[,] { { $"Exception: {ex.Message}" } };
                }
            })!;
        }
    }
}
