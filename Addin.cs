using ExcelDna.ComInterop;
using ExcelDna.Integration;

namespace XLog
{
  internal class Addin : IExcelAddIn
  {
    public void AutoOpen() {
      ComServer.DllRegisterServer();
    }
    public void AutoClose() {
      ComServer.DllUnregisterServer();
    }
  }
}
