using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace OpenAlgo
{
    /// <summary>
    /// Shared HTTP transport for every OpenAlgo REST call made by the add-in.
    ///
    /// A single static HttpClient is reused for the lifetime of the Excel process.
    /// Creating one HttpClient per call (the previous behaviour) leaks a TCP socket
    /// per request and exhausts ephemeral ports when many cells recalculate.
    /// </summary>
    public static class OpenAlgoClient
    {
        private static readonly HttpClient Http = CreateClient();

        private static HttpClient CreateClient()
        {
            var handler = new HttpClientHandler
            {
                // Local OpenAlgo servers are frequently fronted by a self-signed
                // certificate. Only relax validation for loopback addresses.
                ServerCertificateCustomValidationCallback = (request, cert, chain, errors) =>
                {
                    if (errors == System.Net.Security.SslPolicyErrors.None)
                        return true;
                    var host = request.RequestUri?.Host ?? "";
                    return host == "127.0.0.1" || host == "localhost" || host == "::1";
                }
            };

            var client = new HttpClient(handler)
            {
                Timeout = TimeSpan.FromSeconds(OpenAlgoConfig.TimeoutSeconds)
            };
            client.DefaultRequestHeaders.Add("User-Agent", "OpenAlgo-Excel/" + OpenAlgoConfig.AddInVersion);
            return client;
        }

        /// <summary>
        /// Base REST address, for example http://127.0.0.1:5000/api/v1
        /// </summary>
        public static string BaseUrl =>
            $"{OpenAlgoConfig.HostUrl.TrimEnd('/')}/api/{OpenAlgoConfig.Version}";

        /// <summary>
        /// Builds an absolute endpoint URL from a relative path such as "quotes".
        /// </summary>
        public static string Endpoint(string path) => $"{BaseUrl}/{path.TrimStart('/')}";

        /// <summary>
        /// True when the user has supplied an API key through oa_api().
        /// </summary>
        public static bool HasApiKey => !string.IsNullOrWhiteSpace(OpenAlgoConfig.ApiKey);

        /// <summary>
        /// Adds the configured API key to a payload and returns it.
        /// </summary>
        public static JObject WithApiKey(JObject payload)
        {
            payload["apikey"] = OpenAlgoConfig.ApiKey;
            return payload;
        }

        /// <summary>
        /// POSTs a JSON payload and always returns a JObject.
        ///
        /// Transport failures, non-JSON bodies and HTTP error codes are normalised into
        /// { "status": "error", "message": ... } so callers never have to guard a parse.
        /// </summary>
        public static async Task<JObject> PostAsync(string path, JObject payload)
        {
            if (!HasApiKey)
                return Error("OpenAlgo API Key is not set. Use oa_api() first.");

            WithApiKey(payload);

            try
            {
                // OpenAlgo authenticates POST requests with "apikey" in the JSON body.
                // There is no API key header on this API.
                using (var content = new StringContent(payload.ToString(), Encoding.UTF8, "application/json"))
                using (var response = await Http.PostAsync(Endpoint(path), content).ConfigureAwait(false))
                {
                    string body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    return Interpret(response, body, path);
                }
            }
            catch (TaskCanceledException)
            {
                return Error($"Request to /{path.TrimStart('/')} timed out after {OpenAlgoConfig.TimeoutSeconds}s.");
            }
            catch (Exception ex)
            {
                return Error($"{path.TrimStart('/')}: {Flatten(ex)}");
            }
        }

        /// <summary>
        /// GETs a JSON document with the API key and any extra query parameters appended.
        /// </summary>
        public static async Task<JObject> GetAsync(string path, params (string Key, string? Value)[] query)
        {
            if (!HasApiKey)
                return Error("OpenAlgo API Key is not set. Use oa_api() first.");

            try
            {
                // GET requests carry the key as the "apikey" query parameter.
                using (var response = await Http.GetAsync(BuildUrl(path, query)).ConfigureAwait(false))
                {
                    string body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    return Interpret(response, body, path);
                }
            }
            catch (TaskCanceledException)
            {
                return Error($"Request to /{path.TrimStart('/')} timed out after {OpenAlgoConfig.TimeoutSeconds}s.");
            }
            catch (Exception ex)
            {
                return Error($"{path.TrimStart('/')}: {Flatten(ex)}");
            }
        }

        /// <summary>
        /// GETs a resource as raw text. Used by endpoints that can return CSV or plain text.
        /// On failure the Body is null and Error explains why.
        /// </summary>
        public static async Task<(string? Body, string? Error)> GetRawAsync(string path, params (string Key, string? Value)[] query)
        {
            if (!HasApiKey)
                return (null, "OpenAlgo API Key is not set. Use oa_api() first.");

            try
            {
                using (var response = await Http.GetAsync(BuildUrl(path, query)).ConfigureAwait(false))
                {
                    string body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    if (response.IsSuccessStatusCode)
                        return (body, null);
                    return (null, $"HTTP {(int)response.StatusCode} {response.StatusCode}: {Truncate(body)}");
                }
            }
            catch (TaskCanceledException)
            {
                return (null, $"Request to /{path.TrimStart('/')} timed out after {OpenAlgoConfig.TimeoutSeconds}s.");
            }
            catch (Exception ex)
            {
                return (null, $"{path.TrimStart('/')}: {Flatten(ex)}");
            }
        }

        private static string BuildUrl(string path, (string Key, string? Value)[] query)
        {
            var parts = new List<string> { "apikey=" + Uri.EscapeDataString(OpenAlgoConfig.ApiKey) };
            foreach (var (key, value) in query)
            {
                if (string.IsNullOrWhiteSpace(key) || value == null)
                    continue;
                parts.Add(Uri.EscapeDataString(key) + "=" + Uri.EscapeDataString(value));
            }
            return Endpoint(path) + "?" + string.Join("&", parts);
        }

        /// <summary>
        /// Turns an HTTP response into a JObject, preserving any server supplied message.
        /// </summary>
        private static JObject Interpret(HttpResponseMessage response, string body, string path)
        {
            JObject? parsed = null;
            if (!string.IsNullOrWhiteSpace(body))
            {
                try
                {
                    var token = JToken.Parse(body);
                    parsed = token as JObject ?? new JObject { ["status"] = "success", ["data"] = token };
                }
                catch
                {
                    parsed = null;
                }
            }

            if (response.IsSuccessStatusCode)
            {
                if (parsed != null)
                    return parsed;
                return Error($"/{path.TrimStart('/')} returned a non-JSON body: {Truncate(body)}");
            }

            // The server usually explains the failure in the body. Keep that text.
            string message = parsed?["message"]?.ToString()
                             ?? parsed?["error"]?.ToString()
                             ?? Truncate(body);
            if (string.IsNullOrWhiteSpace(message))
                message = response.ReasonPhrase ?? "request failed";

            return Error($"HTTP {(int)response.StatusCode}: {message}");
        }

        /// <summary>
        /// Builds a normalised error document.
        /// </summary>
        public static JObject Error(string message) =>
            new JObject { ["status"] = "error", ["message"] = message };

        /// <summary>
        /// True when a response document reports a failure.
        /// </summary>
        public static bool IsError(JObject response)
        {
            string status = response["status"]?.ToString() ?? "";
            return status.Equals("error", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Extracts the human readable message from a response document.
        /// </summary>
        public static string MessageOf(JObject response) =>
            response["message"]?.ToString()
            ?? response["error"]?.ToString()
            ?? "Unknown error";

        private static string Flatten(Exception ex)
        {
            var inner = ex;
            while (inner.InnerException != null)
                inner = inner.InnerException;
            return inner.Message;
        }

        private static string Truncate(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            text = text.Replace("\r", " ").Replace("\n", " ").Trim();
            return text.Length <= 300 ? text : text.Substring(0, 300) + "...";
        }
    }
}
