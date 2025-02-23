using ExcelDna.Integration;

namespace OpenAlgo
{
    public static class OpenAlgoFunctions
    {
        /// <summary>
        /// Allows setting the global API Key, API Version, and Host URL for OpenAlgo.
        /// </summary>
        [ExcelFunction(Description = "Set the OpenAlgo API Key, API Version, and Host URL globally.")]
        public static string oa_api(
            [ExcelArgument(Name = "API Key", Description = "API Key for authentication (Mandatory)")] string apiKey,
            [ExcelArgument(Name = "API Version", Description = "API Version (default: v1)")] object versionOptional,
            [ExcelArgument(Name = "Host URL", Description = "Base API URL (default: http://127.0.0.1:5000)")] object hostUrlOptional)
        {
            if (string.IsNullOrWhiteSpace(apiKey))
            {
                return "Error: API Key is required.";
            }

            OpenAlgoConfig.ApiKey = apiKey;
            OpenAlgoConfig.Version = versionOptional is ExcelMissing || versionOptional == null ? "v1" : versionOptional.ToString()!;
            OpenAlgoConfig.HostUrl = hostUrlOptional is ExcelMissing || hostUrlOptional == null ? "http://127.0.0.1:5000" : hostUrlOptional.ToString()!;

            return $"Configuration updated: API Key Set, Version = {OpenAlgoConfig.Version}, Host = {OpenAlgoConfig.HostUrl}";
        }
    }
}
