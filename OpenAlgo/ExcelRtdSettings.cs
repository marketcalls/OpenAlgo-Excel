using System;
using ExcelDna.Integration;

namespace OpenAlgo
{
    /// <summary>
    /// Controls Excel's own RealTimeData refresh rate.
    ///
    /// There are two independent throttles between a tick arriving and a cell changing:
    ///
    ///   1. OpenAlgoConfig.StreamThrottleMs, applied by this add-in, limits how often a
    ///      topic is pushed. Set with oa_ws_throttle().
    ///   2. Application.RTD.ThrottleInterval, applied by Excel, limits how often Excel
    ///      collects pushed values from any RTD server. Set with oa_rtd_interval().
    ///
    /// Excel's default for the second is 2000 ms, so a cell updates once every two
    /// seconds no matter how fast data arrives. That default is invisible from the
    /// add-in's side, which makes it look like streaming has stalled. The value is a
    /// per user Excel setting, not a workbook setting, and Excel persists it.
    ///
    /// Earlier versions of this add-in sidestepped the interval by calling
    /// Application.Calculate() on a timer, which is what made Excel unusable and broke
    /// copy and paste (GitHub issue #4). Lowering the interval is the supported way to
    /// get the same responsiveness without touching the calculation engine.
    /// </summary>
    public static class ExcelRtdSettings
    {
        private static int _lastApplied = -1;

        /// <summary>
        /// Applies the configured interval to Excel. Safe to call from any thread and at
        /// any time: the work is queued onto the main thread, and a failure is ignored
        /// because the COM object is not always reachable during start up.
        /// </summary>
        public static void Apply()
        {
            Apply(OpenAlgoConfig.RtdThrottleMs);
        }

        /// <summary>
        /// Applies a specific interval in milliseconds. 0 means update as soon as data
        /// arrives. Excel treats -1 as "only refresh on manual recalculation".
        /// </summary>
        public static void Apply(int milliseconds)
        {
            if (milliseconds < -1)
                milliseconds = 0;

            // Excel raises the interval back to its own value on some transitions, so
            // reapply rather than skipping when the number has not changed since the
            // last successful write.
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    dynamic? app = ExcelDnaUtil.Application;
                    if (app == null)
                        return;

                    app.RTD.ThrottleInterval = milliseconds;
                    _lastApplied = milliseconds;
                }
                catch (Exception)
                {
                    // Excel is not always ready to answer this during AutoOpen, and some
                    // sandboxed hosts refuse the property. Streaming still works, just at
                    // Excel's default rate, so this must never surface as an error.
                }
            });
        }

        /// <summary>
        /// Reads the interval currently set in Excel, or -2 when it cannot be read.
        /// </summary>
        public static int Read()
        {
            try
            {
                dynamic? app = ExcelDnaUtil.Application;
                if (app == null)
                    return -2;
                return (int)app.RTD.ThrottleInterval;
            }
            catch (Exception)
            {
                return _lastApplied;
            }
        }
    }
}
