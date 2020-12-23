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
  [Guid("601B0F10-0E83-43E7-B5E4-251697C2032C")]
  public interface IFileLogger
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
  [Guid("F8D842DC-48BB-4FBC-A73A-527A332A42CD")]
  public class FileLogger : LoggerBase, IFileLogger
  {

    static readonly Logger ilogger = LogManager.GetCurrentClassLogger();

    string logDir;
    string logFileName;
    string logSuffix = Configuration.FileLoggerLogSuffix;
    FileTarget target;

    protected override string TargetLayout { get { return ((SimpleLayout)target.Layout).Text; } set { target.Layout = value; } }

    public bool ArchivalSet => target.ArchiveFileName != null;
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

        if (config.FindRuleByName(loggerId) != null) {

          if (!createNew) {
            target = config.FindTargetByName<FileTarget>(loggerId);
            logger = LogManager.GetLogger(loggerId);
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

        ConfigLogger(wbName, context, target, minLogLevel);

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
