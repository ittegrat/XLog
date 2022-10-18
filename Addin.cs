using System;
using System.Runtime.InteropServices;
using ExcelDna.ComInterop;
using ExcelDna.Integration;
using ExcelDna.Integration.Extensibility;
using NLog;

namespace XLog
{

  [ProgId("XLog.UnregisterHelper")]
  [Guid("1A6478CD-C63C-4F95-8EC7-0C77025E8B08")]
  public sealed class UnregisterHelper : ExcelComAddIn
  {
    static readonly NLog.Logger ilogger = LogManager.GetCurrentClassLogger();
    public UnregisterHelper() {
      FriendlyName = "XLog Unregister Helper";
      Description = "Unregister XLog COM objects on disconnection.";
    }
    public override void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom) {
      try {
        ComServer.DllUnregisterServer();
      }
      catch (Exception ex) {
        ilogger.Error(ex, "ComServer unregister failure");
      }
    }
  }

  internal class Addin : IExcelAddIn
  {

    static readonly NLog.Logger ilogger = LogManager.GetCurrentClassLogger();
    UnregisterHelper uHelper;

    public void AutoOpen() {
      try {
        ComServer.DllRegisterServer();
        try {
          uHelper = new UnregisterHelper();
          ExcelComAddInHelper.LoadComAddIn(uHelper);
        }
        catch (Exception ex) {
          ilogger.Error(ex, "UnregisterHelper not loaded");
        }
      }
      catch (Exception ex) {
        ilogger.Error(ex, "ComServer register failure");
      }
    }
    public void AutoClose() {
      try {
        ComServer.DllUnregisterServer();
      }
      catch (Exception ex) {
        ilogger.Error(ex, "ComServer unregister failure");
      }
    }

  }

}
