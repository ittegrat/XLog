using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using NLog;
using NLog.Config;
using NLog.Targets;

namespace XLog
{

  [ComVisible(true)]
  [InterfaceType(ComInterfaceType.InterfaceIsDual)]
  [Guid("CA497F3C-90F9-4A1F-A2D9-BAC0D412FC25")]
  public interface ILogger
  {

    bool Initialized { get; }
    bool IsClone { get; }
    string MinLogLevel { get; }
    string MaxLogLevel { get; }
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

  }

  [ComVisible(true)]
  [ClassInterface(ClassInterfaceType.None)]
  [ComDefaultInterface(typeof(ILogger))]
  [Guid("30872E2C-678B-4870-B194-38301BB21D11")]
  public abstract class Logger : ILogger
  {

    static readonly NLog.Logger ilogger = LogManager.GetCurrentClassLogger();

    protected const string NOT_INITIALIZED = "<Not initialized>";

    protected NLog.Logger logger;
    protected LoggingRule rule;

    public bool Initialized => logger != null;
    public bool IsClone { get; protected set; } = false;
    public string Name => Initialized ? logger.Name : NOT_INITIALIZED;

    public string MinLogLevel => Initialized ? (rule.Levels.Min() ?? LogLevel.Off).ToString() : NOT_INITIALIZED;
    public string MaxLogLevel => Initialized ? (rule.Levels.Max() ?? LogLevel.Off).ToString() : NOT_INITIALIZED;

    protected abstract string TargetLayout { get; set; }

    public string Layout {
      get {
        return Initialized ? TargetLayout : NOT_INITIALIZED;
      }
      set {
        if (!Initialized)
          throw new NotInitializedException();
        TargetLayout = value;
        LogManager.ReconfigExistingLoggers();
      }
    }

    public void Fatal(string message) { logger?.Fatal(message); }
    public void Error(string message) { logger?.Error(message); }
    public void Warn(string message) { logger?.Warn(message); }
    public void Info(string message) { logger?.Info(message); }
    public void Debug(string message) { logger?.Debug(message); }
    public void Trace(string message) { logger?.Trace(message); }

    public bool IsEnabled(string level) {

      if (!Initialized)
        throw new InvalidOperationException("Not initialized.");

      try {
        var ll = LogLevel.FromString(level);
        return logger.IsEnabled(ll);
      }
      catch (Exception ex) {
        ilogger.Error(ex, "Internal error");
        throw new InvalidOperationException(ex.Message);
      }

    }
    public void SetLogLevels(string MinLevel, string MaxLevel = "") {

      if (!Initialized)
        throw new InvalidOperationException("Not initialized.");

      try {

        var min = LogLevel.FromString(MinLevel);
        var max = LogLevel.FromString(MaxLevel != String.Empty ? MaxLevel : MinLevel);

        rule.DisableLoggingForLevels(LogLevel.Trace, LogLevel.Fatal);
        rule.EnableLoggingForLevels(min, max);
        LogManager.ReconfigExistingLoggers();

      }
      catch (Exception ex) {
        ilogger.Error(ex, "Internal error");
        throw new InvalidOperationException(ex.Message);
      }

    }

    protected void ConfigLogger(string loggerId, string wbName, string context, Target target, string minLogLevel) {

      var config = GetConfig();

      config.AddTarget(target);

      var logLevel = LogLevel.FromString(minLogLevel);
      rule = new LoggingRule(loggerId, logLevel, target) {
        RuleName = loggerId,
        Final = true
      };

      lock (config.LoggingRules)
        config.LoggingRules.Add(rule);

      logger = GetLogger(loggerId, wbName, context);

      LogManager.ReconfigExistingLoggers();

    }
    protected LoggingConfiguration GetConfig() {
      var config = LogManager.Configuration;
      if (config == null) {
        config = new LoggingConfiguration();
        LogManager.Configuration = config;
      }
      return config;
    }
    protected string GetLoggerId(string wbName, string context) {
      var sb = new StringBuilder(this.GetType().Name);
      sb.Append("::").Append(wbName);
      if (!String.IsNullOrWhiteSpace(context))
        sb.Append("::").Append(context);
      return sb.ToString();
    }
    protected NLog.Logger GetLogger(string loggerId, string wbName, string context) {
      var props = new Dictionary<string, object> { { "WbName", wbName } };
      if (!String.IsNullOrEmpty(context))
        props.Add("Context", context);
      return LogManager.GetLogger(loggerId).WithProperties(props);
    }

  }

}
