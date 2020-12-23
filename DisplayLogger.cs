using System;
using System.Runtime.InteropServices;

using NLog;
using NLog.Layouts;
using NLog.Targets;

namespace XLog
{

  [ComVisible(true)]
  [InterfaceType(ComInterfaceType.InterfaceIsDual)]
  [Guid("4A702E1D-38A7-45F7-AABB-993EF2834738")]
  public interface IDisplayLogger
  {

    // Logger properties and methods
    bool Initialized { get; }
    string Name { get; }

    string Layout { get; set; }

    void Fatal(string message);
    void Error(string message);
    void Warn(string message);
    void Info(string message);
    void Debug(string message);
    void Trace(string message);

    bool IsEnabled(string Level);
    void SetLogLevels(string MinLevel, string MaxLevel = "");

    // DisplayLogger properties and methods
    void Initialize(string WbName, string Context = "", bool CreateNew = false, string MinLogLevel = "Info",
      bool AutoShow = false
    );

  }

  [ComVisible(true)]
  [ClassInterface(ClassInterfaceType.None)]
  [ComDefaultInterface(typeof(IDisplayLogger))]
  [Guid("74710B8F-871B-44BD-942E-0CA5A5FD9221")]
  public class DisplayLogger : LoggerBase, IDisplayLogger
  {

    static readonly Logger ilogger = LogManager.GetCurrentClassLogger();

    MethodCallParameter mcp;

    protected override string TargetLayout { get { return ((SimpleLayout)mcp.Layout).Text; } set { mcp.Layout = value; } }

    public void Initialize(string wbName, string context, bool createNew, string minLogLevel, bool autoShow) {

      if (Initialized)
        throw new InvalidOperationException("Already initialized.");

      if (String.IsNullOrWhiteSpace(wbName))
        throw new ArgumentException("Invalid Workbook.Name.");
      wbName = wbName.Trim();
      context = context?.Trim();

      try {

        var config = GetConfig();
        var loggerId = GetLoggerId(wbName, context);

        if (config.FindRuleByName(loggerId) != null) {

          if (!createNew) {
            mcp = ((MethodCallTarget)config.FindTargetByName(loggerId)).Parameters[0];
            logger = LogManager.GetLogger(loggerId);
            return;
          }

          config.RemoveRuleByName(loggerId);
          config.RemoveTarget(loggerId);

        }

        var target = new MethodCallTarget(loggerId) {
          ClassName = typeof(ExcelDna.Logging.LogDisplay).AssemblyQualifiedName,
          MethodName = autoShow ? "SetText" : "RecordMessage",
          Parameters = { new MethodCallParameter(Configuration.DisplayLoggerLayout) }
        };
        mcp = target.Parameters[0];

        ConfigLogger(wbName, context, target, minLogLevel);

      }
      catch (Exception ex) {
        ilogger.Error(ex, "Internal error");
        throw new InvalidOperationException(ex.Message);
      }

    }

  }
}
