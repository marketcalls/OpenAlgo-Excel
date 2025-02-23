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
            [ExcelArgument(Name = "Exchange", Description = "Exchange code")] string exchange,
            [ExcelArgument(Name = "PriceType", Description = "Price type (MARKET/LIMIT)")] string priceType,
            [ExcelArgument(Name = "Product", Description = "Product type (MIS/CNC)")] string product,
            [ExcelArgument(Name = "Quantity", Description = "Order quantity")] object? quantity = null,
            [ExcelArgument(Name = "Price", Description = "Order price (optional)")] object? price = null,
            [ExcelArgument(Name = "TriggerPrice", Description = "Trigger price (optional)")] object? triggerPrice = null,
            [ExcelArgument(Name = "DisclosedQuantity", Description = "Disclosed quantity (optional)")] object? disclosedQuantity = null)
        {
            if (string.IsNullOrWhiteSpace(OpenAlgoConfig.ApiKey))
                return new object[,] { { "Error: API Key is not set. Use SetOpenAlgoConfig()" } };

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
    }
}
