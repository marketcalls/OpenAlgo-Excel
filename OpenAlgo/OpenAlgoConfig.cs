using System;
using System.IO;
using Newtonsoft.Json.Linq;

namespace OpenAlgo
{
    public static class OpenAlgoConfig
    {
        public static string ApiKey { get; set; } = "";
        public static string Version { get; set; } = "v1";
        public static string HostUrl { get; set; } = "http://127.0.0.1:5000";

        private static readonly string ConfigFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "OpenAlgo",
            "config.json"
        );

        /// <summary>
        /// Saves current config to disk so it persists across Excel restarts
        /// </summary>
        public static void Save()
        {
            try
            {
                var dir = Path.GetDirectoryName(ConfigFilePath);
                if (dir != null && !Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                var config = new JObject
                {
                    ["api_key"] = ApiKey,
                    ["version"] = Version,
                    ["host_url"] = HostUrl
                };
                File.WriteAllText(ConfigFilePath, config.ToString());
            }
            catch
            {
                // Silently fail if save fails
            }
        }

        /// <summary>
        /// Loads saved config from disk. Called on add-in startup.
        /// </summary>
        public static void Load()
        {
            try
            {
                if (!File.Exists(ConfigFilePath))
                    return;

                var json = JObject.Parse(File.ReadAllText(ConfigFilePath));
                var key = json["api_key"]?.ToString();
                var ver = json["version"]?.ToString();
                var host = json["host_url"]?.ToString();

                if (!string.IsNullOrWhiteSpace(key))
                    ApiKey = key;
                if (!string.IsNullOrWhiteSpace(ver))
                    Version = ver;
                if (!string.IsNullOrWhiteSpace(host))
                    HostUrl = host;
            }
            catch
            {
                // Silently fail if load fails
            }
        }
    }
}
