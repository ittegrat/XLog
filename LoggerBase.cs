using System;
using System.Text;

using NLog;
using NLog.Config;
using NLog.Targets;

namespace XLog
{
  public abstract class LoggerBase
  {

    static readonly Logger ilogger = LogManager.GetCurrentClassLogger();

    protected const string NOT_INITIALIZED = "<Not initialized>";

    protected Logger logger;

    public bool Initialized => logger != null;
    public string Name => Initialized ? logger.Name : NOT_INITIALIZED;

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

        var config = LogManager.Configuration;
        var rule = config.FindRuleByName(Name);
        rule.DisableLoggingForLevels(LogLevel.Trace, LogLevel.Fatal);
        rule.EnableLoggingForLevels(min, max);
        LogManager.ReconfigExistingLoggers();

      }
      catch (Exception ex) {
        ilogger.Error(ex, "Internal error");
        throw new InvalidOperationException(ex.Message);
      }

    }

    protected void ConfigLogger(string wbName, string context, Target target, string minLogLevel) {

      var config = GetConfig();
      var loggerId = GetLoggerId(wbName, context);

      config.AddTarget(target);

      var logLevel = LogLevel.FromString(minLogLevel);
      var rule = new LoggingRule(loggerId, logLevel, target) {
        RuleName = loggerId,
        Final = true
      };

      lock (config.LoggingRules)
        config.LoggingRules.Add(rule);

      logger = LogManager.GetLogger(loggerId);
      logger.SetProperty("WbName", wbName);
      if (!String.IsNullOrEmpty(context))
        logger.SetProperty("Context", context);

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

  }
}
