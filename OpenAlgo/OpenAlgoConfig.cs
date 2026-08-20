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
        /// cell. Ticks that arrive sooner update the cache but do not wake Excel.
        ///
        /// Defaults to 0, meaning every tick is pushed, which is the DDE like behaviour
        /// most users expect from a live quote sheet. Raise it with oa_ws_throttle() if a
        /// very large sheet on a fast feed starts to feel heavy.
        /// </summary>
        public static int StreamThrottleMs { get; set; } = 0;

        /// <summary>
        /// Value written to Excel's Application.RTD.ThrottleInterval, in milliseconds.
        ///
        /// This is Excel's own limit on how often it collects values from any RTD server,
        /// and it is separate from StreamThrottleMs above. Excel ships with 2000, so a
        /// streaming cell updates only once every two seconds however fast the add-in
        /// pushes. That default is invisible from inside the add-in, which makes live
        /// data look stalled.
        ///
        /// 0 means collect as soon as data arrives. Excel treats -1 as "only on manual
        /// recalculation", so -1 effectively freezes streaming.
        /// </summary>
        public static int RtdThrottleMs { get; set; } = 0;

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
                    ["stream_throttle_ms"] = StreamThrottleMs,
                    ["rtd_throttle_ms"] = RtdThrottleMs
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

                // "rtd_throttle_ms" was added at the same time the streaming defaults were
                // reworked. Its absence marks a config file written by a build that
                // predates the fix, whose stream_throttle_ms of 250 was our default rather
                // than a considered user choice. Combined with Excel's own 2000 ms RTD
                // interval that made streaming look frozen, so do not carry it forward.
                bool preRtdFixConfig = json["rtd_throttle_ms"] == null;

                int throttle = json["stream_throttle_ms"]?.ToObject<int?>() ?? -1;
                if (throttle >= 0 && !preRtdFixConfig)
                    StreamThrottleMs = throttle;

                int rtd = json["rtd_throttle_ms"]?.ToObject<int?>() ?? -2;
                if (rtd >= -1)
                    RtdThrottleMs = rtd;

                if (preRtdFixConfig)
                    Save();
            }
            catch
            {
                // Silently fail if load fails
            }
        }
    }
}
