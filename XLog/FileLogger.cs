using System;
using System.IO;
using System.Runtime.InteropServices;
using NLog;
using NLog.Layouts;
using NLog.Targets;

namespace XLog
{

  [ComVisible(true)]
  [InterfaceType(ComInterfaceType.InterfaceIsDual)]
  [Guid("F5E39EC0-9D7B-4C38-B9E1-B68133881971")]
  public interface IFileLogger : ILogger
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

    // FileLogger properties and methods
    bool ArchivalSet { get; }
    string LogFile { get; }

    void Initialize(string WbFullName, string Context = "", bool CreateNew = false, string MinLogLevel = "Info",
      string LogDir = "", string LogFileName = "", string LogSuffix = "", bool NewFile = true
    );
    void ArchivalByDate(string Every, string DateFormat, int MaxArchiveDays = 0, string ArchiveSuffix = "", bool NewArchive = true);
    void ArchivalByNumber(string Numbering, int MaxArchiveFiles = 0, string ArchiveSuffix = "");

  }

  [ComVisible(true)]
  [ClassInterface(ClassInterfaceType.None)]
  [ComDefaultInterface(typeof(IFileLogger))]
  [Guid("102731C9-D788-4A44-9C40-07439FEEAF2B")]
  public class FileLogger : Logger, IFileLogger
  {

    static readonly NLog.Logger ilogger = LogManager.GetCurrentClassLogger();

    string logDir;
    string logFileName;
    string logSuffix = Configuration.FileLoggerLogSuffix;
    FileTarget target;

    protected override string TargetLayout { get { return ((SimpleLayout)target.Layout).Text; } set { target.Layout = value; } }

    public bool ArchivalSet => target?.ArchiveFileName != null;
    public string LogFile => Initialized ? ((SimpleLayout)target.FileName).Text : NOT_INITIALIZED;

    public void Initialize(string wbFullName, string context, bool createNew, string minLogLevel, string logDir, string logFileName, string logSuffix, bool newFile) {

      if (Initialized)
        throw new InvalidOperationException("Already initialized.");

      if (String.IsNullOrWhiteSpace(wbFullName))
        throw new ArgumentException("Invalid Workbook.FullName.");
      wbFullName = wbFullName.Trim();
      context = context?.Trim();

      try {

        var wbName = Path.GetFileName(wbFullName);

        var config = GetConfig();
        var loggerId = GetLoggerId(wbName, context);

        rule = config.FindRuleByName(loggerId);
        if (rule != null) {

          if (!createNew) {
            target = config.FindTargetByName<FileTarget>(loggerId);
            logger = GetLogger(loggerId, wbName, context);
            IsClone = true;
            return;
          }

          config.RemoveRuleByName(loggerId);
          config.RemoveTarget(loggerId);

        }

        this.logDir = String.IsNullOrWhiteSpace(logDir)
          ? (File.Exists(wbFullName) ? Path.GetDirectoryName(wbFullName) : Path.GetTempPath())
          : logDir.Trim();

        this.logFileName = String.IsNullOrWhiteSpace(logFileName)
          ? wbName
          : logFileName.Trim();

        if (!String.IsNullOrWhiteSpace(logSuffix))
          this.logSuffix = logSuffix.Trim();

        target = new FileTarget(loggerId) {
          FileName = Path.Combine(this.logDir, $"{this.logFileName}{this.logSuffix}"),
          Layout = Configuration.FileLoggerLayout,
          DeleteOldFileOnStartup = newFile,
          NetworkWrites = true
        };

        ConfigLogger(loggerId, wbName, context, target, minLogLevel);

      }
      catch (Exception ex) {
        ilogger.Error(ex, "Internal error");
        throw new InvalidOperationException(ex.Message);
      }

    }
    public void ArchivalByDate(string every, string dateFormat, int maxArchiveDays, string archiveSuffix, bool newArchive) {

      if (!Initialized)
        throw new NotInitializedException();

      if (ArchivalSet)
        throw new InvalidOperationException("Archival options already set.");

      if (!Enum.TryParse<FileArchivePeriod>(every.Trim(), out var period))
        throw new ArgumentException($"Invalid ArchivePeriod '{every}'.");

      if (period == FileArchivePeriod.None)
        throw new ArgumentException($"Unsupported ArchivePeriod '{period}'.");

      if (String.IsNullOrWhiteSpace(dateFormat))
        throw new ArgumentException("Invalid DateFormat.");

      try {

        if (String.IsNullOrWhiteSpace(archiveSuffix))
          archiveSuffix = $".{{#}}{logSuffix}";
        else
          archiveSuffix = archiveSuffix.Trim();

        target.DeleteOldFileOnStartup = false;
        target.ArchiveOldFileOnStartup = newArchive;
        target.ArchiveFileName = Path.Combine(logDir, $"{logFileName}{archiveSuffix}");
        target.ArchiveNumbering = ArchiveNumberingMode.Date;
        target.ArchiveEvery = period;
        target.ArchiveDateFormat = dateFormat.Trim();
        target.MaxArchiveDays = maxArchiveDays;

        LogManager.ReconfigExistingLoggers();

      }
      catch (Exception ex) {
        ilogger.Error(ex, "Internal error");
        throw new InvalidOperationException(ex.Message);
      }

    }
    public void ArchivalByNumber(string NumberingMode, int maxArchiveFiles, string archiveSuffix) {

      if (!Initialized)
        throw new NotInitializedException();

      if (ArchivalSet)
        throw new InvalidOperationException("Archival options already set.");

      if (!Enum.TryParse<ArchiveNumberingMode>(NumberingMode.Trim(), out var mode))
        throw new ArgumentException($"Invalid NumberingMode '{NumberingMode}'.");

      if ((mode != ArchiveNumberingMode.Rolling) && (mode != ArchiveNumberingMode.Sequence))
        throw new ArgumentException($"Unsupported ArchiveNumbering '{mode}'.");

      try {

        if (String.IsNullOrWhiteSpace(archiveSuffix))
          archiveSuffix = $"{Configuration.FileLoggerNumberSuffix}{logSuffix}";
        else
          archiveSuffix = archiveSuffix.Trim();

        target.DeleteOldFileOnStartup = false;
        target.ArchiveOldFileOnStartup = true;
        target.ArchiveFileName = Path.Combine(logDir, $"{logFileName}{archiveSuffix}");
        target.ArchiveNumbering = mode;
        target.MaxArchiveFiles = maxArchiveFiles;

        LogManager.ReconfigExistingLoggers();

      }
      catch (Exception ex) {
        ilogger.Error(ex, "Internal error");
        throw new InvalidOperationException(ex.Message);
      }

    }

  }
}
