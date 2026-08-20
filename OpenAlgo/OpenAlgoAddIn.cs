using System;
using ExcelDna.Integration;
using ExcelDna.IntelliSense;

namespace OpenAlgo
{
    public class OpenAlgoAddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            // Load saved API key and config from previous session
            OpenAlgoConfig.Load();

            // Lower Excel's own RTD collection interval. It ships at 2000 ms, which caps
            // every streaming cell at one update every two seconds no matter how fast the
            // feed is. Broker feeds run at roughly 1 to 11 updates per second, so the
            // default silently discards most of them and live data looks frozen.
            ExcelRtdSettings.Apply();

            // Register IntelliSense
            IntelliSenseServer.Install();
        }

        public void AutoClose()
        {
            // Close the streaming connection before the add-in unloads. Without this the
            // socket, its receive loop and any pending throttle flush timers stay alive,
            // which matters when the add-in is unloaded and loaded again inside one Excel
            // session: the stale receive loop would keep writing into the cache.
            //
            // Excel is shutting down here, so the wait is bounded. A connection that will
            // not close in time is abandoned rather than holding up the UI.
            try
            {
                WebSocketManager.Instance.DisconnectAsync().Wait(TimeSpan.FromSeconds(2));
            }
            catch (Exception)
            {
                // Never let shutdown cleanup surface an error to the user.
            }

            // Unregister IntelliSense on close
            IntelliSenseServer.Uninstall();
        }
    }
}
