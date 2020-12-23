using System;
using System.IO;
using System.Runtime.InteropServices;

using NLog;
using NLog.Config;
using NLog.Targets;

namespace XLog
{

  [ComVisible(true)]
  [InterfaceType(ComInterfaceType.InterfaceIsDual)]
  [Guid("601B0F10-0E83-43E7-B5E4-251697C2032C")]
  public interface IFileLogger
  {

    bool ArchivalSet { get; }
    bool Initialized { get; }
    string LogFile { get; }
    string Suffix { get; }

    string Layout { get; set; }

    void Initialize(string WbFullName, string BaseDir = "", string FileName = "", string LogSuffix = "", int Instance = 0);
    void ArchivalByDate(string Every, string DateFormat, int MaxArchiveDays = 0, string ArchiveSuffix = "");
    void ArchivalByNumber(string Numbering, int MaxArchiveFiles = 0, string ArchiveSuffix = "");
    void SetLogLevels(string MinLevel, string MaxLevel = "");

    void Fatal(string message);
    void Error(string message);
    void Warn(string message);
    void Info(string message);
    void Debug(string message);
    void Trace(string message);

  }

  [ComVisible(true)]
  [ClassInterface(ClassInterfaceType.None)]
  [ComDefaultInterface(typeof(IFileLogger))]
  [Guid("F8D842DC-48BB-4FBC-A73A-527A332A42CD")]
  public class FileLogger : IFileLogger
  {

    string fileName;
    string loggerId;
    Logger logger;

    public bool ArchivalSet { get; private set; } = false;
    public string Suffix { get; private set; } = Configuration.DefaultSuffix;

    public bool Initialized => logger != null;
    public string LogFile {
      get {
        try {
          return GetTarget().FileName.ToString();
        }
        catch (NotInitializedException) { return "<Not initialized>"; }
      }
    }

    public string Layout {
      get {
        try {
          return GetTarget().Layout.ToString();
        }
        catch (NotInitializedException) { return "<Not initialized>"; }
      }
      set {
        GetTarget().Layout = value;
        LogManager.ReconfigExistingLoggers();
      }
    }

    public void Initialize(string WbFullName, string BaseDir = "", string FileName = "", string LogSuffix = "", int Instance = 0) {

      if (Initialized)
        throw new InvalidOperationException("Already initialized.");

      if (String.IsNullOrWhiteSpace(WbFullName))
        throw new ArgumentException("Invalid Workbook.FullName.");

      try {

        var config = LogManager.Configuration;
        if (config == null) {
          config = new LoggingConfiguration();
          LogManager.Configuration = config;
        }

        loggerId = Path.GetFileName(WbFullName) + (Instance > 0 ? $"_{Instance}" : String.Empty);

        if (config.FindRuleByName(loggerId) != null)
          config.RemoveRuleByName(loggerId);

        if (config.FindTargetByName(loggerId) != null)
          config.RemoveTarget(loggerId);

        var bd = String.IsNullOrWhiteSpace(BaseDir)
          ? (File.Exists(WbFullName) ? Path.GetDirectoryName(WbFullName) : Path.GetTempPath())
          : BaseDir.Trim();

        var fn = String.IsNullOrWhiteSpace(FileName)
          ? loggerId
          : FileName.Trim();

        if (!String.IsNullOrWhiteSpace(LogSuffix))
          Suffix = LogSuffix.Trim();

        fileName = Path.Combine(bd, $"{fn}{Suffix}");

        var target = new FileTarget(loggerId) {
          FileName = fileName,
          Layout = Configuration.DefaultLayout,
          DeleteOldFileOnStartup = Configuration.DeleteOldFileOnStartup,
          NetworkWrites = true
        };

        var rule = new LoggingRule(loggerId, LogLevel.Info, target) {
          RuleName = loggerId,
          Final = true
        };

        lock (config.LoggingRules)
          config.LoggingRules.Add(rule);

        config.AddTarget(target);

        LogManager.ReconfigExistingLoggers();

        logger = LogManager.GetLogger(loggerId);

      }
      catch (Exception ex) {
        throw new InvalidOperationException(ex.Message);
      }

    }
    public void ArchivalByDate(string Every, string DateFormat, int MaxArchiveDays = 0, string ArchiveSuffix = "") {

      if (ArchivalSet)
        throw new InvalidOperationException("Archival options already set.");

      if (!Enum.TryParse<FileArchivePeriod>(Every.Trim(), out var every))
        throw new ArgumentException($"Invalid ArchivePeriod '{Every}'.");

      if (every == FileArchivePeriod.None)
        throw new ArgumentException($"Unsupported ArchivePeriod '{every}'.");

      if (String.IsNullOrWhiteSpace(DateFormat))
        throw new ArgumentException("Invalid DateFormat.");

      var target = GetTarget();

      try {

        if (String.IsNullOrWhiteSpace(ArchiveSuffix))
          ArchiveSuffix = $".{{#}}{Suffix}";
        else
          ArchiveSuffix = ArchiveSuffix.Trim();

        target.DeleteOldFileOnStartup = false;
        target.ArchiveOldFileOnStartup = Configuration.ArchiveOldFileOnStartup;
        target.ArchiveFileName = $"{fileName.Substring(0, fileName.Length - Suffix.Length)}{ArchiveSuffix}";
        target.ArchiveNumbering = ArchiveNumberingMode.Date;
        target.ArchiveEvery = every;
        target.ArchiveDateFormat = DateFormat.Trim();
        target.MaxArchiveDays = MaxArchiveDays;

        LogManager.ReconfigExistingLoggers();

        ArchivalSet = true;

      }
      catch (Exception ex) {
        throw new InvalidOperationException(ex.Message);
      }

    }
    public void ArchivalByNumber(string Numbering, int MaxArchiveFiles = 0, string ArchiveSuffix = "") {

      if (ArchivalSet)
        throw new InvalidOperationException("Archival options already set.");

      if (!Enum.TryParse<ArchiveNumberingMode>(Numbering.Trim(), out var mode))
        throw new ArgumentException($"Invalid NumberingMode '{Numbering}'.");

      if ((mode != ArchiveNumberingMode.Rolling) && (mode != ArchiveNumberingMode.Sequence))
        throw new ArgumentException($"Unsupported ArchiveNumbering '{mode}'.");

      var target = GetTarget();

      try {

        if (String.IsNullOrWhiteSpace(ArchiveSuffix))
          ArchiveSuffix = $"{Configuration.DefaultNumberSuffix}{Suffix}";
        else
          ArchiveSuffix = ArchiveSuffix.Trim();

        target.DeleteOldFileOnStartup = false;
        target.ArchiveOldFileOnStartup = true;
        target.ArchiveFileName = $"{fileName.Substring(0, fileName.Length - Suffix.Length)}{ArchiveSuffix}";
        target.ArchiveNumbering = mode;
        target.MaxArchiveFiles = MaxArchiveFiles;

        LogManager.ReconfigExistingLoggers();

        ArchivalSet = true;

      }
      catch (Exception ex) {
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
        var rule = config.FindRuleByName(loggerId);
        rule.DisableLoggingForLevels(LogLevel.Trace, LogLevel.Fatal);
        rule.EnableLoggingForLevels(min, max);
        LogManager.ReconfigExistingLoggers();

      }
      catch (Exception ex) {
        throw new InvalidOperationException(ex.Message);
      }

    }

    public void Fatal(string message) { logger?.Fatal(message); }
    public void Error(string message) { logger?.Error(message); }
    public void Warn(string message) { logger?.Warn(message); }
    public void Info(string message) { logger?.Info(message); }
    public void Debug(string message) { logger?.Debug(message); }
    public void Trace(string message) { logger?.Trace(message); }

    FileTarget GetTarget() {

      if (!Initialized)
        throw new NotInitializedException();

      var target = LogManager.Configuration.FindTargetByName<FileTarget>(loggerId);

      if (target == null)
        throw new InvalidOperationException("Null FileTarget.");

      return target;

    }

  }
}
