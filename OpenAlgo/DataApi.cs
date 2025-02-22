using ExcelDna.Integration;
using ExcelDna.Registration.Utils;
using Newtonsoft.Json.Linq;

namespace OpenAlgo
{
    public static class DataApi
    {
        [ExcelFunction(Description = "Retrieve quotes from OpenAlgo API.")]
        public static object[,] GetQuotes(
            [ExcelArgument(Name = "Symbol", Description = "Symbol of the asset")] string symbol)
        {
            if (string.IsNullOrWhiteSpace(OpenAlgoConfig.ApiKey))
                return new object[,] { { "Error: API Key is not set. Use SetOpenAlgoConfig()" } };

            string endpoint = $"{OpenAlgoConfig.HostUrl}/api/{OpenAlgoConfig.Version}/quotes";
            var payload = new JObject
            {
                ["apikey"] = OpenAlgoConfig.ApiKey,
                ["symbol"] = symbol
            };

            return (object[,])AsyncTaskUtil.RunTask(nameof(GetQuotes), new object[] { }, async () =>
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
    }
}
