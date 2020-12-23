using System;
using System.Collections;
using System.Configuration;

using NLog;

namespace XLog
{
  internal static class Configuration
  {

    static readonly Logger ilogger = LogManager.GetCurrentClassLogger();

    const string CONFIG_SECTION = "XLog";

    public static string DisplayLoggerLayout { get; }
    public static string FileLoggerLayout { get; }
    public static string FileLoggerNumberSuffix { get; }
    public static string FileLoggerLogSuffix { get; }

    static Configuration() {
      Hashtable section = null;
      try {
        ConfigurationManager.RefreshSection(CONFIG_SECTION);
        section = (Hashtable)ConfigurationManager.GetSection(CONFIG_SECTION);
      }
      catch (Exception ex) {
        ilogger.Error(ex, "Internal error");
      }

      T GetValue<T>(string key, T @default) {
        if (section != null && section.ContainsKey(key))
          return (T)Convert.ChangeType(section[key], typeof(T));
        return @default;
      }

      DisplayLoggerLayout = GetValue("displaylogger.layout", @"${date:format=H\:mm\:ss}|${level:uppercase=true}|${event-properties:WbName}${when:when='${event-properties:Context}'!='':inner=|${event-properties:Context}}|${message}");

      FileLoggerLayout = GetValue("filelogger.layout", "${longdate}|${level:uppercase=true}|${when:when='${event-properties:Context}'!='':inner=|${event-properties:Context}}|${message}");
      FileLoggerNumberSuffix = GetValue("filelogger.numbersuffix", ".{###}");
      FileLoggerLogSuffix = GetValue("filelogger.logsuffix", ".log");

    }

  }
}
