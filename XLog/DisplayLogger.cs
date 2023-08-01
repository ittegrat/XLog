using System;
using System.Runtime.InteropServices;
using NLog;
using NLog.Layouts;
using NLog.Targets;

namespace XLog
{

  [ComVisible(true)]
  [InterfaceType(ComInterfaceType.InterfaceIsDual)]
  [Guid("BB985023-8FBE-46C8-ABC6-1519712167B7")]
  public interface IDisplayLogger : ILogger
  {

    // Logger properties and methods
    new bool Initialized { get; }
    new bool IsClone { get; }
    new string MinLogLevel { get; }
    new string MaxLogLevel { get; }
    new string Name { get; }

    new string Layout { get; set; }

    new void Fatal(string message);
    new void Error(string message);
    new void Warn(string message);
    new void Info(string message);
    new void Debug(string message);
    new void Trace(string message);

    new bool IsEnabled(string Level);
    new void SetLogLevels(string MinLevel, string MaxLevel = "");

    // DisplayLogger properties and methods
    void Initialize(string WbName, string Context = "", bool CreateNew = false, string MinLogLevel = "Info",
      bool AutoShow = false
    );

  }

  [ComVisible(true)]
  [ClassInterface(ClassInterfaceType.None)]
  [ComDefaultInterface(typeof(IDisplayLogger))]
  [Guid("3CA9AE18-EFB2-4525-82B0-978F42DA0DF6")]
  public class DisplayLogger : Logger, IDisplayLogger
  {

    static readonly NLog.Logger ilogger = LogManager.GetCurrentClassLogger();

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

        rule = config.FindRuleByName(loggerId);
        if (rule != null) {

          if (!createNew) {
            mcp = ((MethodCallTarget)config.FindTargetByName(loggerId)).Parameters[0];
            logger = GetLogger(loggerId, wbName, context);
            IsClone = true;
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

        ConfigLogger(loggerId, wbName, context, target, minLogLevel);

      }
      catch (Exception ex) {
        ilogger.Error(ex, "Internal error");
        throw new InvalidOperationException(ex.Message);
      }

    }

  }
}
