using System;
using System.IO;
using Newtonsoft.Json.Linq;

namespace OpenAlgo
{
    public static class OpenAlgoConfig
    {
        /// <summary>
        /// Version reported in the User-Agent header and by oa_version().
        /// </summary>
        public const string AddInVersion = "1.0.5";

        public static string ApiKey { get; set; } = "";
        public static string Version { get; set; } = "v1";
        public static string HostUrl { get; set; } = "http://127.0.0.1:5000";

        /// <summary>
        /// WebSocket endpoint used by the oa_ws_* functions.
        /// </summary>
        public static string WebSocketUrl { get; set; } = "ws://127.0.0.1:8765";

        /// <summary>
        /// REST request timeout in seconds.
        /// </summary>
        public static int TimeoutSeconds { get; set; } = 30;

        /// <summary>
        /// Minimum gap in milliseconds between two pushed updates for the same streaming
        /// cell. Ticks that arrive sooner update the cache but do not wake Excel, which
        /// keeps the UI responsive on fast moving symbols.
        /// </summary>
        public static int StreamThrottleMs { get; set; } = 250;

        /// <summary>
        /// When false, order placing functions refuse to fire. Recalculating a sheet
        /// re-evaluates every order formula on it, so live trading has to be armed
        /// deliberately through oa_trading_enabled().
        ///
        /// This defaults to false and is deliberately not persisted. Opening a saved
        /// workbook that already contains order formulas must never arm trading on the
        /// user's behalf: a full rebuild (Ctrl+Alt+F9) would then re-place every order
        /// on the sheet. Arming is a per session decision the user makes each time.
        /// </summary>
        public static bool TradingEnabled { get; set; } = false;

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
                    ["host_url"] = HostUrl,
                    ["ws_url"] = WebSocketUrl,
                    ["timeout_seconds"] = TimeoutSeconds,
                    ["stream_throttle_ms"] = StreamThrottleMs
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
                var ws = json["ws_url"]?.ToString();

                if (!string.IsNullOrWhiteSpace(key))
                    ApiKey = key;
                if (!string.IsNullOrWhiteSpace(ver))
                    Version = ver;
                if (!string.IsNullOrWhiteSpace(host))
                    HostUrl = host;
                if (!string.IsNullOrWhiteSpace(ws))
                    WebSocketUrl = ws;

                int timeout = json["timeout_seconds"]?.ToObject<int?>() ?? 0;
                if (timeout > 0)
                    TimeoutSeconds = timeout;

                int throttle = json["stream_throttle_ms"]?.ToObject<int?>() ?? -1;
                if (throttle >= 0)
                    StreamThrottleMs = throttle;
            }
            catch
            {
                // Silently fail if load fails
            }
        }
    }
}
