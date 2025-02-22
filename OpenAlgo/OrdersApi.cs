using ExcelDna.Integration;
using ExcelDna.Registration.Utils;
using Newtonsoft.Json.Linq;

namespace OpenAlgo
{
    public static class OrdersApi
    {
        [ExcelFunction(Description = "Place an order using OpenAlgo API.")]
        public static object PlaceOrder(
            [ExcelArgument(Name = "Symbol", Description = "Symbol of the asset")] string symbol,
            [ExcelArgument(Name = "Quantity", Description = "Number of units")] int quantity,
            [ExcelArgument(Name = "OrderType", Description = "Order Type (Market/Limit)")] string orderType,
            [ExcelArgument(Name = "Price", Description = "Price (for limit orders)")] double price)
        {
            if (string.IsNullOrWhiteSpace(OpenAlgoConfig.ApiKey))
                return "Error: API Key is not set. Use SetOpenAlgoConfig()";

            string endpoint = $"{OpenAlgoConfig.HostUrl}/api/{OpenAlgoConfig.Version}/placeorder";
            var payload = new JObject
            {
                ["apikey"] = OpenAlgoConfig.ApiKey,
                ["symbol"] = symbol,
                ["quantity"] = quantity,
                ["orderType"] = orderType,
                ["price"] = price
            };

            return AsyncTaskUtil.RunTask(nameof(PlaceOrder), new object[] { }, async () =>
            {
                string response = await Utilities.PostRequestAsync(endpoint, payload);
                return response;
            })!;
        }
    }
}
