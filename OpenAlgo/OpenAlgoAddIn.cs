using ExcelDna.Integration;
using ExcelDna.IntelliSense;

namespace OpenAlgo
{
    public class OpenAlgoAddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            // ✅ Register IntelliSense
            IntelliSenseServer.Install();
        }

        public void AutoClose()
        {
            // ✅ Unregister IntelliSense on close
            IntelliSenseServer.Uninstall();
        }
    }
}
