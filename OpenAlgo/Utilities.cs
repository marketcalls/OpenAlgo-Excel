using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace OpenAlgo
{
    public static class Utilities
    {
        /// <summary>
        /// Sends an HTTP POST request to the OpenAlgo API.
        ///
        /// Retained for the original call sites that build an absolute endpoint URL and a
        /// payload that already carries the API key. New code should call
        /// <see cref="OpenAlgoClient.PostAsync"/> with a relative path instead.
        /// </summary>
        public static async Task<string> PostRequestAsync(string endpoint, JObject payload)
        {
            string path = ToRelativePath(endpoint);
            JObject response = await OpenAlgoClient.PostAsync(path, payload).ConfigureAwait(false);
            return response.ToString();
        }

        /// <summary>
        /// Strips the configured base URL from an absolute endpoint so it can be routed
        /// through the shared client.
        /// </summary>
        private static string ToRelativePath(string endpoint)
        {
            string prefix = OpenAlgoClient.BaseUrl;
            if (endpoint.StartsWith(prefix, System.StringComparison.OrdinalIgnoreCase))
                return endpoint.Substring(prefix.Length).TrimStart('/');

            int marker = endpoint.IndexOf("/api/", System.StringComparison.OrdinalIgnoreCase);
            if (marker >= 0)
            {
                string tail = endpoint.Substring(marker + "/api/".Length);
                int slash = tail.IndexOf('/');
                if (slash >= 0)
                    return tail.Substring(slash + 1);
            }

            return endpoint.TrimStart('/');
        }
    }
}
