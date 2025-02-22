using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace OpenAlgo
{
    public static class Utilities
    {
        /// <summary>
        /// Sends an HTTP POST request to the OpenAlgo API.
        /// </summary>
        public static async Task<string> PostRequestAsync(string endpoint, JObject payload)
        {
            using (var client = new HttpClient())
            {
                var content = new StringContent(payload.ToString(), Encoding.UTF8, "application/json");
                HttpResponseMessage response = await client.PostAsync(endpoint, content);

                return response.IsSuccessStatusCode
                    ? await response.Content.ReadAsStringAsync()
                    : $"Error: {response.StatusCode}";
            }
        }
    }
}
