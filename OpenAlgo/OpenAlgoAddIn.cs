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

            // Register IntelliSense
            IntelliSenseServer.Install();
        }

        public void AutoClose()
        {
            // ✅ Unregister IntelliSense on close
            IntelliSenseServer.Uninstall();
        }
    }
}
