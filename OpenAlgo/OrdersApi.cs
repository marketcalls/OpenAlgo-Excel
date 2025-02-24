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
                return new object[,] { { "Error: OpenAlgo API Key is not set. Use oa_api()" } };

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
                return new object[,] { { "Error: OpenAlgo API Key is not set. Use oa_api()" } };

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

        /// <summary>
        /// Places a split order via OpenAlgo API.
        /// </summary>
        [ExcelFunction(
            Name = "oa_splitorder",
            Description = "Places a split order through OpenAlgo API.")]
        public static object[,] oa_splitorder(
            [ExcelArgument(Name = "Strategy", Description = "Trading strategy name")] string strategy,
            [ExcelArgument(Name = "Symbol", Description = "Trading symbol")] string symbol,
            [ExcelArgument(Name = "Action", Description = "Order action (BUY/SELL)")] string action,
            [ExcelArgument(Name = "Exchange", Description = "Exchange code")] string exchange,
            [ExcelArgument(Name = "Quantity", Description = "Total order quantity")] object? quantity = null,
            [ExcelArgument(Name = "SplitSize", Description = "Size of each split order")] object? splitsize = null,
            [ExcelArgument(Name = "PriceType", Description = "Price type (MARKET/LIMIT)")] string priceType = "MARKET",
            [ExcelArgument(Name = "Product", Description = "Product type (MIS/CNC/MIS)")] string product = "MIS",
            [ExcelArgument(Name = "Price", Description = "Order price (optional)")] object? price = null,
            [ExcelArgument(Name = "TriggerPrice", Description = "Trigger price (optional)")] object? triggerPrice = null,
            [ExcelArgument(Name = "DisclosedQuantity", Description = "Disclosed quantity (optional)")] object? disclosedQuantity = null)
        {
            if (string.IsNullOrWhiteSpace(OpenAlgoConfig.ApiKey))
                return new object[,] { { "Error: OpenAlgo API Key is not set. Use oa_api()" } };

            // Convert all values to strings
            string quantityStr = (quantity is ExcelMissing or null) ? "0" : quantity.ToString()!;
            string splitsizeStr = (splitsize is ExcelMissing or null) ? "0" : splitsize.ToString()!;
            string priceStr = (price is ExcelMissing or null) ? "0" : price.ToString()!;
            string triggerPriceStr = (triggerPrice is ExcelMissing or null) ? "0" : triggerPrice.ToString()!;
            string disclosedQuantityStr = (disclosedQuantity is ExcelMissing or null) ? "0" : disclosedQuantity.ToString()!;

            string endpoint = $"{OpenAlgoConfig.HostUrl}/api/{OpenAlgoConfig.Version}/splitorder";

            var payload = new JObject
            {
                ["apikey"] = OpenAlgoConfig.ApiKey,
                ["strategy"] = strategy,
                ["symbol"] = symbol,
                ["action"] = action,
                ["exchange"] = exchange,
                ["quantity"] = quantityStr,
                ["splitsize"] = splitsizeStr,
                ["pricetype"] = priceType,
                ["product"] = product,
                ["price"] = priceStr,
                ["trigger_price"] = triggerPriceStr,
                ["disclosed_quantity"] = disclosedQuantityStr
            };

            return (object[,])AsyncTaskUtil.RunTask(nameof(oa_splitorder), new object[] { }, async () =>
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
                                var output = new object[results.Count + 1, 4];
                                output[0, 0] = "Order Number";
                                output[0, 1] = "Order ID";
                                output[0, 2] = "Quantity";
                                output[0, 3] = "Status";

                                for (int i = 0; i < results.Count; i++)
                                {
                                    var result = results[i];
                                    output[i + 1, 0] = result["order_num"]?.ToString() ?? string.Empty;
                                    output[i + 1, 1] = result["orderid"]?.ToString() ?? string.Empty;
                                    output[i + 1, 2] = result["quantity"]?.ToString() ?? string.Empty;
                                    output[i + 1, 3] = result["status"]?.ToString() ?? string.Empty;
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

        /// <summary>
        /// Modifies an existing order via OpenAlgo API.
        /// </summary>
        [ExcelFunction(
            Name = "oa_modifyorder",
            Description = "Modifies an existing order through OpenAlgo API.")]
        public static object[,] oa_modifyorder(
            [ExcelArgument(Name = "Strategy", Description = "Trading strategy name")] string strategy,
            [ExcelArgument(Name = "OrderID", Description = "ID of the order to modify")] string orderId,
            [ExcelArgument(Name = "Symbol", Description = "Trading symbol")] string symbol,
            [ExcelArgument(Name = "Action", Description = "Order action (BUY/SELL)")] string action,
            [ExcelArgument(Name = "Exchange", Description = "Exchange code")] string exchange,
            [ExcelArgument(Name = "Quantity", Description = "Order quantity")] object? quantity = null,
            [ExcelArgument(Name = "PriceType", Description = "Price type (MARKET/LIMIT)")] string priceType = "MARKET",
            [ExcelArgument(Name = "Product", Description = "Product type (MIS/CNC)")] string product = "MIS",
            [ExcelArgument(Name = "Price", Description = "Order price (optional)")] object? price = null,
            [ExcelArgument(Name = "TriggerPrice", Description = "Trigger price (optional)")] object? triggerPrice = null,
            [ExcelArgument(Name = "DisclosedQuantity", Description = "Disclosed quantity (optional)")] object? disclosedQuantity = null)
        {
            if (string.IsNullOrWhiteSpace(OpenAlgoConfig.ApiKey))
                return new object[,] { { "Error: OpenAlgo API Key is not set. Use oa_api()" } };

            string quantityStr = (quantity is ExcelMissing or null) ? "0" : quantity.ToString()!;
            string priceStr = (price is ExcelMissing or null) ? "0" : price.ToString()!;
            string triggerPriceStr = (triggerPrice is ExcelMissing or null) ? "0" : triggerPrice.ToString()!;
            string disclosedQuantityStr = (disclosedQuantity is ExcelMissing or null) ? "0" : disclosedQuantity.ToString()!;

            string endpoint = $"{OpenAlgoConfig.HostUrl}/api/{OpenAlgoConfig.Version}/modifyorder";

            var payload = new JObject
            {
                ["apikey"] = OpenAlgoConfig.ApiKey,
                ["strategy"] = strategy,
                ["orderid"] = orderId,
                ["symbol"] = symbol,
                ["action"] = action,
                ["exchange"] = exchange,
                ["quantity"] = quantityStr,
                ["pricetype"] = priceType,
                ["product"] = product,
                ["price"] = priceStr,
                ["trigger_price"] = triggerPriceStr,
                ["disclosed_quantity"] = disclosedQuantityStr
            };

            return (object[,])AsyncTaskUtil.RunTask(nameof(oa_modifyorder), new object[] { }, async () =>
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
                            string status = jsonResponse["status"]?.ToString() ?? "Unknown";
                            string message = jsonResponse["message"]?.ToString() ?? "No message provided";

                            return new object[,]
                            {
                                { "Status", "Message" },
                                { status, message }
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
        /// Cancels an existing order via OpenAlgo API.
        /// </summary>
        [ExcelFunction(
            Name = "oa_cancelorder",
            Description = "Cancels an existing order through OpenAlgo API.")]
        public static object[,] oa_cancelorder(
            [ExcelArgument(Name = "Strategy", Description = "Trading strategy name")] string strategy,
            [ExcelArgument(Name = "OrderID", Description = "ID of the order to cancel")] string orderId)
        {
            if (string.IsNullOrWhiteSpace(OpenAlgoConfig.ApiKey))
                return new object[,] { { "Error: OpenAlgo API Key is not set. Use oa_api()" } };

            string endpoint = $"{OpenAlgoConfig.HostUrl}/api/{OpenAlgoConfig.Version}/cancelorder";

            var payload = new JObject
            {
                ["apikey"] = OpenAlgoConfig.ApiKey,
                ["strategy"] = strategy,
                ["orderid"] = orderId
            };

            return (object[,])AsyncTaskUtil.RunTask(nameof(oa_cancelorder), new object[] { }, async () =>
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
                            string status = jsonResponse["status"]?.ToString() ?? "Unknown";
                            string message = jsonResponse["message"]?.ToString() ?? "No message provided";

                            return new object[,]
                            {
                        { "Status", "Message" },
                        { status, message }
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
        /// Cancels all open orders for a given strategy via OpenAlgo API.
        /// </summary>
        [ExcelFunction(
            Name = "oa_cancelallorder",
            Description = "Cancels all open orders for a specified strategy through OpenAlgo API.")]
        public static object[,] oa_cancelallorder(
            [ExcelArgument(Name = "Strategy", Description = "Trading strategy name")] string strategy)
        {
            if (string.IsNullOrWhiteSpace(OpenAlgoConfig.ApiKey))
                return new object[,] { { "Error: OpenAlgo API Key is not set. Use oa_api()" } };

            string endpoint = $"{OpenAlgoConfig.HostUrl}/api/{OpenAlgoConfig.Version}/cancelallorder";

            var payload = new JObject
            {
                ["apikey"] = OpenAlgoConfig.ApiKey,
                ["strategy"] = strategy
            };

            return (object[,])AsyncTaskUtil.RunTask(nameof(oa_cancelallorder), new object[] { }, async () =>
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
                            var canceledOrders = jsonResponse["canceled_orders"] as JArray;
                            var failedCancellations = jsonResponse["failed_cancellations"] as JArray;
                            string message = jsonResponse["message"]?.ToString() ?? "No message provided";
                            string status = jsonResponse["status"]?.ToString() ?? "Unknown";

                            // Prepare output
                            int rowCount = Math.Max(canceledOrders?.Count ?? 0, failedCancellations?.Count ?? 0) + 2;
                            var output = new object[rowCount, 3];
                            output[0, 0] = "Status";
                            output[0, 1] = "Message";
                            output[0, 2] = "";
                            output[1, 0] = status;
                            output[1, 1] = message;
                            output[1, 2] = "";

                            if (canceledOrders != null && canceledOrders.Count > 0)
                            {
                                output[1, 2] = "Canceled Orders";
                                for (int i = 0; i < canceledOrders.Count; i++)
                                {
                                    output[i + 2, 2] = canceledOrders[i]?.ToString() ?? string.Empty;
                                }
                            }

                            if (failedCancellations != null && failedCancellations.Count > 0)
                            {
                                output[1, 2] = "Failed Cancellations";
                                for (int i = 0; i < failedCancellations.Count; i++)
                                {
                                    output[i + 2, 2] = failedCancellations[i]?.ToString() ?? string.Empty;
                                }
                            }

                            return output;
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
        /// Closes all open positions for a given strategy via OpenAlgo API.
        /// </summary>
        [ExcelFunction(
            Name = "oa_closeposition",
            Description = "Closes all open positions for a specified strategy through OpenAlgo API.")]
        public static object[,] oa_closeposition(
            [ExcelArgument(Name = "Strategy", Description = "Trading strategy name")] string strategy)
        {
            if (string.IsNullOrWhiteSpace(OpenAlgoConfig.ApiKey))
                return new object[,] { { "Error: OpenAlgo API Key is not set. Use oa_api()" } };

            string endpoint = $"{OpenAlgoConfig.HostUrl}/api/{OpenAlgoConfig.Version}/closeposition";

            var payload = new JObject
            {
                ["apikey"] = OpenAlgoConfig.ApiKey,
                ["strategy"] = strategy
            };

            return (object[,])AsyncTaskUtil.RunTask(nameof(oa_closeposition), new object[] { }, async () =>
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
                            var closedPositions = jsonResponse["closed_positions"] as JArray;
                            var failedClosures = jsonResponse["failed_closures"] as JArray;
                            string message = jsonResponse["message"]?.ToString() ?? "No message provided";
                            string status = jsonResponse["status"]?.ToString() ?? "Unknown";

                            // Prepare output
                            int rowCount = Math.Max(closedPositions?.Count ?? 0, failedClosures?.Count ?? 0) + 2;
                            var output = new object[rowCount, 3];
                            output[0, 0] = "Status";
                            output[0, 1] = "Message";
                            output[0, 2] = "";
                            output[1, 0] = status;
                            output[1, 1] = message;
                            output[1, 2] = "";

                            if (closedPositions != null && closedPositions.Count > 0)
                            {
                                output[1, 2] = "Closed Positions";
                                for (int i = 0; i < closedPositions.Count; i++)
                                {
                                    output[i + 2, 2] = closedPositions[i]?.ToString() ?? string.Empty;
                                }
                            }

                            if (failedClosures != null && failedClosures.Count > 0)
                            {
                                output[1, 2] = "Failed Closures";
                                for (int i = 0; i < failedClosures.Count; i++)
                                {
                                    output[i + 2, 2] = failedClosures[i]?.ToString() ?? string.Empty;
                                }
                            }

                            return output;
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
        /// Retrieves the status of a specific order via OpenAlgo API.
        /// </summary>
        [ExcelFunction(
            Name = "oa_orderstatus",
            Description = "Retrieves the status of a specific order through OpenAlgo API.")]
        public static object[,] oa_orderstatus(
            [ExcelArgument(Name = "Strategy", Description = "Trading strategy name associated with the order")] string strategy,
            [ExcelArgument(Name = "OrderID", Description = "ID of the order to retrieve status for")] string orderId)
        {
            if (string.IsNullOrWhiteSpace(OpenAlgoConfig.ApiKey))
                return new object[,] { { "Error: OpenAlgo API Key is not set. Use oa_api()" } };

            string endpoint = $"{OpenAlgoConfig.HostUrl}/api/{OpenAlgoConfig.Version}/orderstatus";

            var payload = new JObject
            {
                ["apikey"] = OpenAlgoConfig.ApiKey,
                ["strategy"] = strategy,
                ["orderid"] = orderId
            };

            return (object[,])AsyncTaskUtil.RunTask(nameof(oa_orderstatus), new object[] { }, async () =>
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
                            string status = jsonResponse["status"]?.ToString() ?? "Unknown";
                            string message = jsonResponse["message"]?.ToString() ?? "No message provided";
                            var orderDetails = jsonResponse["order_details"] as JObject;

                            if (orderDetails != null)
                            {
                                var output = new object[orderDetails.Count + 2, 2];
                                output[0, 0] = "Status";
                                output[0, 1] = "Message";
                                output[1, 0] = status;
                                output[1, 1] = message;

                                int row = 2;
                                foreach (var detail in orderDetails)
                                {
                                    output[row, 0] = detail.Key;
                                    output[row, 1] = detail.Value?.ToString() ?? string.Empty;
                                    row++;
                                }

                                return output;
                            }
                            else
                            {
                                return new object[,] { { "Error: Order details not found." } };
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

        /// <summary>
        /// Retrieves open positions for a given strategy via OpenAlgo API.
        /// </summary>
        [ExcelFunction(
            Name = "oa_openposition",
            Description = "Retrieves open positions for a specified strategy through OpenAlgo API.")]
        public static object[,] oa_openposition(
            [ExcelArgument(Name = "Strategy", Description = "Trading strategy name")] string strategy,
            [ExcelArgument(Name = "Symbol", Description = "Trading symbol")] string symbol,
            [ExcelArgument(Name = "Exchange", Description = "Exchange code")] string exchange,
            [ExcelArgument(Name = "Product", Description = "Product type (e.g., CNC, MIS)")] string product)
        {
            if (string.IsNullOrWhiteSpace(OpenAlgoConfig.ApiKey))
                return new object[,] { { "Error: OpenAlgo API Key is not set. Use oa_api()" } };

            string endpoint = $"{OpenAlgoConfig.HostUrl}/api/{OpenAlgoConfig.Version}/openposition";

            var payload = new JObject
            {
                ["apikey"] = OpenAlgoConfig.ApiKey,
                ["strategy"] = strategy,
                ["symbol"] = symbol,
                ["exchange"] = exchange,
                ["product"] = product
            };

            return (object[,])AsyncTaskUtil.RunTask(nameof(oa_openposition), new object[] { }, async () =>
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
                            string status = jsonResponse["status"]?.ToString() ?? "Unknown";
                            string message = jsonResponse["message"]?.ToString() ?? "No message provided";
                            var positionDetails = jsonResponse["position_details"] as JObject;

                            if (positionDetails != null)
                            {
                                var output = new object[positionDetails.Count + 2, 2];
                                output[0, 0] = "Status";
                                output[0, 1] = "Message";
                                output[1, 0] = status;
                                output[1, 1] = message;

                                int row = 2;
                                foreach (var detail in positionDetails)
                                {
                                    output[row, 0] = detail.Key;
                                    output[row, 1] = detail.Value?.ToString() ?? string.Empty;
                                    row++;
                                }

                                return output;
                            }
                            else
                            {
                                return new object[,] { { "Error: Position details not found." } };
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
